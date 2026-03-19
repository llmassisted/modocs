package com.modocs.feature.docx

import android.content.Context
import android.net.Uri
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.withContext
import org.xmlpull.v1.XmlPullParser
import org.xmlpull.v1.XmlPullParserFactory
import java.io.ByteArrayInputStream
import java.io.ByteArrayOutputStream
import java.io.InputStream
import java.util.zip.ZipInputStream

/**
 * Parses DOCX files (ZIP archives containing OOXML) into [DocxDocument].
 *
 * A DOCX ZIP typically contains:
 * - word/document.xml — the main document content
 * - word/styles.xml — named styles (Heading1, Normal, etc.)
 * - word/numbering.xml — list/numbering definitions
 * - word/_rels/document.xml.rels — relationships (images, hyperlinks)
 * - word/media/ — embedded images
 */
object DocxParser {

    // OOXML namespaces
    private const val NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    private const val NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    private const val NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    private const val NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
    private const val NS_PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"

    /**
     * Parse a DOCX file from a content URI.
     */
    suspend fun parse(context: Context, uri: Uri): DocxDocument = withContext(Dispatchers.IO) {
        val inputStream = context.contentResolver.openInputStream(uri)
            ?: throw IllegalStateException("Cannot open file")
        parse(inputStream)
    }

    /**
     * Parse a DOCX file from an InputStream.
     */
    fun parse(inputStream: InputStream): DocxDocument {
        // Step 1: Read all ZIP entries into memory
        val entries = mutableMapOf<String, ByteArray>()
        ZipInputStream(inputStream.buffered()).use { zip ->
            var entry = zip.nextEntry
            while (entry != null) {
                if (!entry.isDirectory) {
                    val baos = ByteArrayOutputStream()
                    zip.copyTo(baos)
                    entries[entry.name] = baos.toByteArray()
                }
                zip.closeEntry()
                entry = zip.nextEntry
            }
        }

        // Step 2: Parse relationships (for image references)
        val relationships = parseRelationships(entries["word/_rels/document.xml.rels"])

        // Step 3: Parse styles
        val styles = entries["word/styles.xml"]?.let { parseStyles(it) } ?: emptyMap()

        // Step 4: Parse numbering
        val numbering = entries["word/numbering.xml"]?.let { parseNumbering(it) } ?: emptyMap()

        // Step 5: Extract images
        val images = mutableMapOf<String, ByteArray>()
        for ((relId, target) in relationships) {
            val imagePath = if (target.startsWith("/")) target.removePrefix("/")
            else "word/$target"
            entries[imagePath]?.let { images[relId] = it }
        }

        // Step 6: Parse main document
        val documentXml = entries["word/document.xml"]
            ?: throw IllegalStateException("Missing word/document.xml")
        val (body, pageSetup) = parseDocumentBodyAndSetup(documentXml, styles, relationships)

        return DocxDocument(
            body = body.toMutableList(),
            styles = styles,
            numbering = numbering,
            images = images,
            rawEntries = entries,
            pageSetup = pageSetup,
        )
    }

    // --- Relationships ---

    private fun parseRelationships(bytes: ByteArray?): Map<String, String> {
        if (bytes == null) return emptyMap()
        val rels = mutableMapOf<String, String>()
        val parser = createParser(bytes)

        while (parser.next() != XmlPullParser.END_DOCUMENT) {
            if (parser.eventType == XmlPullParser.START_TAG && parser.name == "Relationship") {
                val id = parser.getAttributeValue(null, "Id") ?: continue
                val target = parser.getAttributeValue(null, "Target") ?: continue
                rels[id] = target
            }
        }
        return rels
    }

    // --- Styles ---

    private fun parseStyles(bytes: ByteArray): Map<String, DocxStyle> {
        val styles = mutableMapOf<String, DocxStyle>()
        val parser = createParser(bytes)

        var currentStyleId: String? = null
        var currentStyleName = ""
        var currentBasedOn: String? = null
        var currentRunProps = RunProperties()
        var currentParaProps = ParagraphProperties()
        var isDefault = false
        var inStyle = false
        var inStyleRPr = false
        var inStylePPr = false
        var depth = 0

        while (parser.next() != XmlPullParser.END_DOCUMENT) {
            when (parser.eventType) {
                XmlPullParser.START_TAG -> {
                    val localName = parser.name ?: continue
                    when {
                        localName == "style" && !inStyle -> {
                            inStyle = true
                            depth = 0
                            currentStyleId = parser.getAttr("styleId")
                            isDefault = parser.getAttr("default") == "1"
                            currentStyleName = ""
                            currentBasedOn = null
                            currentRunProps = RunProperties()
                            currentParaProps = ParagraphProperties()
                        }
                        inStyle && localName == "name" && depth == 0 -> {
                            currentStyleName = parser.getAttr("val") ?: ""
                        }
                        inStyle && localName == "basedOn" && depth == 0 -> {
                            currentBasedOn = parser.getAttr("val")
                        }
                        inStyle && localName == "rPr" && !inStyleRPr && !inStylePPr -> {
                            inStyleRPr = true
                        }
                        inStyle && localName == "pPr" && !inStylePPr && !inStyleRPr -> {
                            inStylePPr = true
                        }
                        inStyleRPr -> {
                            currentRunProps = parseRunProperty(parser, localName, currentRunProps)
                        }
                        inStylePPr -> {
                            currentParaProps = parseParagraphProperty(parser, localName, currentParaProps)
                        }
                    }
                    if (inStyle) depth++
                }
                XmlPullParser.END_TAG -> {
                    val localName = parser.name ?: continue
                    if (inStyle) depth--
                    when {
                        localName == "style" && inStyle && depth <= 0 -> {
                            if (currentStyleId != null) {
                                styles[currentStyleId!!] = DocxStyle(
                                    styleId = currentStyleId!!,
                                    name = currentStyleName,
                                    basedOn = currentBasedOn,
                                    runProperties = currentRunProps,
                                    paragraphProperties = currentParaProps,
                                    isDefault = isDefault,
                                )
                            }
                            inStyle = false
                            inStyleRPr = false
                            inStylePPr = false
                        }
                        localName == "rPr" && inStyleRPr -> inStyleRPr = false
                        localName == "pPr" && inStylePPr -> inStylePPr = false
                    }
                }
            }
        }

        return styles
    }

    // --- Numbering ---

    private fun parseNumbering(bytes: ByteArray): Map<String, NumberingDefinition> {
        val abstractNums = mutableMapOf<String, MutableMap<Int, NumberingLevel>>()
        val numToAbstract = mutableMapOf<String, String>()
        val parser = createParser(bytes)

        var inAbstractNum = false
        var currentAbstractNumId: String? = null
        var inLvl = false
        var currentLevel = 0
        var currentFormat = NumberFormat.BULLET
        var currentText = "•"
        var currentIndent = 720
        var inNum = false
        var currentNumId: String? = null

        while (parser.next() != XmlPullParser.END_DOCUMENT) {
            when (parser.eventType) {
                XmlPullParser.START_TAG -> {
                    val localName = parser.name ?: continue
                    when {
                        localName == "abstractNum" -> {
                            inAbstractNum = true
                            currentAbstractNumId = parser.getAttr("abstractNumId")
                            if (currentAbstractNumId != null) {
                                abstractNums[currentAbstractNumId!!] = mutableMapOf()
                            }
                        }
                        localName == "lvl" && inAbstractNum -> {
                            inLvl = true
                            currentLevel = parser.getAttr("ilvl")?.toIntOrNull() ?: 0
                            currentFormat = NumberFormat.BULLET
                            currentText = "•"
                            currentIndent = 720 * (currentLevel + 1)
                        }
                        localName == "numFmt" && inLvl -> {
                            currentFormat = when (parser.getAttr("val")) {
                                "decimal" -> NumberFormat.DECIMAL
                                "lowerLetter" -> NumberFormat.LOWER_ALPHA
                                "upperLetter" -> NumberFormat.UPPER_ALPHA
                                "lowerRoman" -> NumberFormat.LOWER_ROMAN
                                "upperRoman" -> NumberFormat.UPPER_ROMAN
                                "bullet" -> NumberFormat.BULLET
                                else -> NumberFormat.BULLET
                            }
                        }
                        localName == "lvlText" && inLvl -> {
                            currentText = parser.getAttr("val") ?: "•"
                        }
                        localName == "ind" && inLvl -> {
                            currentIndent = parser.getAttr("left")?.toIntOrNull() ?: currentIndent
                        }
                        localName == "num" && !inAbstractNum -> {
                            inNum = true
                            currentNumId = parser.getAttr("numId")
                        }
                        localName == "abstractNumId" && inNum -> {
                            numToAbstract[currentNumId ?: ""] = parser.getAttr("val") ?: ""
                        }
                    }
                }
                XmlPullParser.END_TAG -> {
                    val localName = parser.name ?: continue
                    when {
                        localName == "lvl" && inLvl -> {
                            inLvl = false
                            currentAbstractNumId?.let { id ->
                                abstractNums[id]?.put(
                                    currentLevel,
                                    NumberingLevel(currentLevel, currentFormat, currentText, currentIndent),
                                )
                            }
                        }
                        localName == "abstractNum" -> inAbstractNum = false
                        localName == "num" -> inNum = false
                    }
                }
            }
        }

        // Build final numbering map: numId -> NumberingDefinition
        val result = mutableMapOf<String, NumberingDefinition>()
        for ((numId, abstractId) in numToAbstract) {
            val levels = abstractNums[abstractId] ?: continue
            result[numId] = NumberingDefinition(abstractId, levels)
        }
        return result
    }

    // --- Main document body ---

    private fun parseDocumentBodyAndSetup(
        bytes: ByteArray,
        styles: Map<String, DocxStyle>,
        relationships: Map<String, String>,
    ): Pair<List<DocxElement>, PageSetup> {
        val elements = mutableListOf<DocxElement>()
        var pageSetup = PageSetup()
        val parser = createParser(bytes)

        // Navigate to w:body
        var foundBody = false
        while (parser.next() != XmlPullParser.END_DOCUMENT) {
            if (parser.eventType == XmlPullParser.START_TAG && parser.name == "body") {
                foundBody = true
                break
            }
        }
        if (!foundBody) return elements to pageSetup

        // Parse body children
        var depth = 1
        while (depth > 0 && parser.next() != XmlPullParser.END_DOCUMENT) {
            when (parser.eventType) {
                XmlPullParser.START_TAG -> {
                    depth++
                    when (parser.name) {
                        "p" -> {
                            val para = parseParagraph(parser, styles, relationships)
                            elements.add(para)
                            depth-- // parseParagraph consumes the end tag
                        }
                        "tbl" -> {
                            val table = parseTable(parser, styles, relationships)
                            elements.add(table)
                            depth-- // parseTable consumes the end tag
                        }
                        "sectPr" -> {
                            pageSetup = parseSectionProperties(parser)
                            depth-- // parseSectionProperties consumes the end tag
                        }
                    }
                }
                XmlPullParser.END_TAG -> depth--
            }
        }

        return elements to pageSetup
    }

    // --- Section properties (page size and margins) ---

    private fun parseSectionProperties(parser: XmlPullParser): PageSetup {
        var setup = PageSetup()
        var depth = 1

        while (depth > 0 && parser.next() != XmlPullParser.END_DOCUMENT) {
            when (parser.eventType) {
                XmlPullParser.START_TAG -> {
                    depth++
                    when (parser.name) {
                        "pgSz" -> {
                            val w = parser.getAttr("w")?.toIntOrNull()
                            val h = parser.getAttr("h")?.toIntOrNull()
                            setup = setup.copy(
                                pageWidthTwips = w ?: setup.pageWidthTwips,
                                pageHeightTwips = h ?: setup.pageHeightTwips,
                            )
                        }
                        "pgMar" -> {
                            val top = parser.getAttr("top")?.toIntOrNull()
                            val bottom = parser.getAttr("bottom")?.toIntOrNull()
                            val left = parser.getAttr("left")?.toIntOrNull()
                            val right = parser.getAttr("right")?.toIntOrNull()
                            setup = setup.copy(
                                marginTopTwips = top ?: setup.marginTopTwips,
                                marginBottomTwips = bottom ?: setup.marginBottomTwips,
                                marginLeftTwips = left ?: setup.marginLeftTwips,
                                marginRightTwips = right ?: setup.marginRightTwips,
                            )
                        }
                    }
                }
                XmlPullParser.END_TAG -> depth--
            }
        }

        return setup
    }

    // --- Paragraph parsing ---

    private fun parseParagraph(
        parser: XmlPullParser,
        styles: Map<String, DocxStyle>,
        relationships: Map<String, String>,
    ): DocxParagraph {
        val runs = mutableListOf<DocxRun>()
        var paraProps = ParagraphProperties()
        var listInfo: ListInfo? = null
        var depth = 1

        while (depth > 0 && parser.next() != XmlPullParser.END_DOCUMENT) {
            when (parser.eventType) {
                XmlPullParser.START_TAG -> {
                    depth++
                    when (parser.name) {
                        "pPr" -> {
                            val (props, list) = parseParagraphProperties(parser, styles)
                            paraProps = props
                            listInfo = list
                            depth-- // consumed end tag
                        }
                        "r" -> {
                            runs.addAll(parseRun(parser, styles))
                            depth-- // consumed end tag
                        }
                        "hyperlink" -> {
                            // Parse runs inside hyperlink
                            runs.addAll(parseHyperlinkRuns(parser, styles, relationships))
                            depth-- // consumed end tag
                        }
                        "drawing" -> {
                            val image = parseDrawing(parser, relationships)
                            if (image != null) {
                                // Wrap image in a special run placeholder
                                runs.add(DocxRun("\uFFFC", RunProperties())) // object replacement char
                            }
                            depth-- // consumed end tag
                        }
                    }
                }
                XmlPullParser.END_TAG -> depth--
            }
        }

        // Apply style to paragraph properties if style is referenced
        val styledProps = applyStyleToParaProps(paraProps, styles)

        return DocxParagraph(runs, styledProps, listInfo)
    }

    private fun parseParagraphProperties(
        parser: XmlPullParser,
        styles: Map<String, DocxStyle>,
    ): Pair<ParagraphProperties, ListInfo?> {
        var props = ParagraphProperties()
        var listInfo: ListInfo? = null
        var depth = 1

        while (depth > 0 && parser.next() != XmlPullParser.END_DOCUMENT) {
            when (parser.eventType) {
                XmlPullParser.START_TAG -> {
                    depth++
                    val localName = parser.name ?: continue
                    when {
                        localName == "pStyle" -> {
                            val styleId = parser.getAttr("val")
                            props = props.copy(styleId = styleId)
                            // Detect heading level from style name
                            val level = styleId?.let { detectHeadingLevel(it, styles) } ?: 0
                            if (level > 0) props = props.copy(headingLevel = level)
                        }
                        localName == "numPr" -> {
                            listInfo = parseNumPr(parser)
                            depth-- // consumed end tag
                        }
                        else -> {
                            props = parseParagraphProperty(parser, localName, props)
                        }
                    }
                }
                XmlPullParser.END_TAG -> depth--
            }
        }

        return props to listInfo
    }

    private fun parseParagraphProperty(
        parser: XmlPullParser,
        localName: String,
        props: ParagraphProperties,
    ): ParagraphProperties {
        return when (localName) {
            "jc" -> {
                val alignment = when (parser.getAttr("val")) {
                    "center" -> ParagraphAlignment.CENTER
                    "right", "end" -> ParagraphAlignment.RIGHT
                    "both", "distribute" -> ParagraphAlignment.JUSTIFY
                    else -> ParagraphAlignment.LEFT
                }
                props.copy(alignment = alignment)
            }
            "spacing" -> {
                val before = parser.getAttr("before")?.toIntOrNull()
                val after = parser.getAttr("after")?.toIntOrNull()
                val line = parser.getAttr("line")?.toIntOrNull()
                val lineRule = parser.getAttr("lineRule") // "auto", "exact", "atLeast"
                val isExact = lineRule == "exact" || lineRule == "atLeast"
                props.copy(
                    spacingBeforePt = before?.let { it / 20f } ?: props.spacingBeforePt,
                    spacingBeforeExplicit = before != null || props.spacingBeforeExplicit,
                    spacingAfterPt = after?.let { it / 20f } ?: props.spacingAfterPt,
                    spacingAfterExplicit = after != null || props.spacingAfterExplicit,
                    lineSpacingMultiplier = if (!isExact) line?.let { it / 240f } ?: props.lineSpacingMultiplier else props.lineSpacingMultiplier,
                    lineSpacingPt = if (isExact) line?.let { it / 20f } ?: props.lineSpacingPt else props.lineSpacingPt,
                    lineSpacingExact = if (line != null) isExact else props.lineSpacingExact,
                    lineSpacingExplicitlySet = line != null || props.lineSpacingExplicitlySet,
                )
            }
            "ind" -> {
                val left = parser.getAttr("left")?.toIntOrNull()
                    ?: parser.getAttr("start")?.toIntOrNull()
                val right = parser.getAttr("right")?.toIntOrNull()
                    ?: parser.getAttr("end")?.toIntOrNull()
                val firstLine = parser.getAttr("firstLine")?.toIntOrNull()
                val hanging = parser.getAttr("hanging")?.toIntOrNull()
                props.copy(
                    indentLeftTwips = left ?: props.indentLeftTwips,
                    indentRightTwips = right ?: props.indentRightTwips,
                    indentFirstLineTwips = firstLine
                        ?: hanging?.let { -it }
                        ?: props.indentFirstLineTwips,
                )
            }
            "keepNext" -> props.copy(keepNext = true)
            "keepLines" -> props.copy(keepLines = true)
            "pageBreakBefore" -> props.copy(pageBreakBefore = true)
            "contextualSpacing" -> props.copy(contextualSpacing = true)
            "outlineLvl" -> {
                val level = parser.getAttr("val")?.toIntOrNull()?.plus(1) ?: 0
                props.copy(headingLevel = level)
            }
            else -> props
        }
    }

    private fun parseNumPr(parser: XmlPullParser): ListInfo {
        var numId = ""
        var level = 0
        var depth = 1

        while (depth > 0 && parser.next() != XmlPullParser.END_DOCUMENT) {
            when (parser.eventType) {
                XmlPullParser.START_TAG -> {
                    depth++
                    when (parser.name) {
                        "ilvl" -> level = parser.getAttr("val")?.toIntOrNull() ?: 0
                        "numId" -> numId = parser.getAttr("val") ?: ""
                    }
                }
                XmlPullParser.END_TAG -> depth--
            }
        }

        return ListInfo(numId, level)
    }

    // --- Run parsing ---

    private fun parseRun(
        parser: XmlPullParser,
        styles: Map<String, DocxStyle>,
    ): List<DocxRun> {
        val runs = mutableListOf<DocxRun>()
        var runProps = RunProperties()
        var depth = 1

        while (depth > 0 && parser.next() != XmlPullParser.END_DOCUMENT) {
            when (parser.eventType) {
                XmlPullParser.START_TAG -> {
                    depth++
                    when (parser.name) {
                        "rPr" -> {
                            runProps = parseRunProperties(parser, styles)
                            depth-- // consumed end tag
                        }
                        "t" -> {
                            val text = parser.nextText() ?: ""
                            runs.add(DocxRun(text, runProps))
                            depth-- // nextText advances past end tag
                        }
                        "br" -> {
                            val type = parser.getAttr("type")
                            if (type == "page") {
                                runs.add(DocxRun("\u000C", runProps))
                            } else {
                                runs.add(DocxRun("\n", runProps))
                            }
                        }
                        "tab" -> {
                            runs.add(DocxRun("\t", runProps))
                        }
                        "cr" -> {
                            runs.add(DocxRun("\n", runProps))
                        }
                    }
                }
                XmlPullParser.END_TAG -> depth--
            }
        }

        return runs
    }

    private fun parseRunProperties(
        parser: XmlPullParser,
        styles: Map<String, DocxStyle>,
    ): RunProperties {
        var props = RunProperties()
        var depth = 1

        while (depth > 0 && parser.next() != XmlPullParser.END_DOCUMENT) {
            when (parser.eventType) {
                XmlPullParser.START_TAG -> {
                    depth++
                    val localName = parser.name ?: continue
                    when (localName) {
                        "rStyle" -> {
                            val styleId = parser.getAttr("val")
                            val style = styleId?.let { styles[it] }
                            if (style != null) {
                                props = mergeRunProps(style.runProperties, props)
                            }
                        }
                        else -> {
                            props = parseRunProperty(parser, localName, props)
                        }
                    }
                }
                XmlPullParser.END_TAG -> depth--
            }
        }

        return props
    }

    private fun parseRunProperty(
        parser: XmlPullParser,
        localName: String,
        props: RunProperties,
    ): RunProperties {
        return when (localName) {
            "b" -> {
                val off = parser.getAttr("val") == "0" || parser.getAttr("val") == "false"
                props.copy(bold = !off)
            }
            "bCs" -> props // ignore complex script bold
            "i" -> {
                val off = parser.getAttr("val") == "0" || parser.getAttr("val") == "false"
                props.copy(italic = !off)
            }
            "iCs" -> props
            "u" -> {
                val val_ = parser.getAttr("val")
                props.copy(underline = val_ != null && val_ != "none")
            }
            "strike" -> {
                val off = parser.getAttr("val") == "0" || parser.getAttr("val") == "false"
                props.copy(strikethrough = !off)
            }
            "rFonts" -> {
                val font = parser.getAttr("ascii")
                    ?: parser.getAttr("hAnsi")
                    ?: parser.getAttr("cs")
                props.copy(fontName = font)
            }
            "sz" -> {
                val size = parser.getAttr("val")?.toIntOrNull()
                props.copy(fontSizeHalfPt = size)
            }
            "szCs" -> props // ignore complex script size
            "color" -> {
                val color = DocxColors.parseColor(parser.getAttr("val"))
                props.copy(color = color)
            }
            "highlight" -> {
                val highlight = DocxColors.parseHighlight(parser.getAttr("val"))
                props.copy(highlight = highlight)
            }
            "vertAlign" -> {
                when (parser.getAttr("val")) {
                    "superscript" -> props.copy(superscript = true)
                    "subscript" -> props.copy(subscript = true)
                    else -> props
                }
            }
            else -> props
        }
    }

    // --- Hyperlink runs ---

    private fun parseHyperlinkRuns(
        parser: XmlPullParser,
        styles: Map<String, DocxStyle>,
        relationships: Map<String, String>,
    ): List<DocxRun> {
        val runs = mutableListOf<DocxRun>()
        var depth = 1

        while (depth > 0 && parser.next() != XmlPullParser.END_DOCUMENT) {
            when (parser.eventType) {
                XmlPullParser.START_TAG -> {
                    depth++
                    if (parser.name == "r") {
                        // Apply underline + blue color to indicate hyperlink
                        val innerRuns = parseRun(parser, styles)
                        runs.addAll(innerRuns.map {
                            it.copy(
                                properties = it.properties.copy(
                                    underline = true,
                                    color = it.properties.color ?: 0xFF0563C1.toInt(),
                                )
                            )
                        })
                        depth-- // parseRun consumed end tag
                    }
                }
                XmlPullParser.END_TAG -> depth--
            }
        }

        return runs
    }

    // --- Drawing / images ---

    private fun parseDrawing(
        parser: XmlPullParser,
        relationships: Map<String, String>,
    ): DocxImage? {
        var relId: String? = null
        var widthEmu = 0L
        var heightEmu = 0L
        var altText = ""
        var depth = 1

        while (depth > 0 && parser.next() != XmlPullParser.END_DOCUMENT) {
            when (parser.eventType) {
                XmlPullParser.START_TAG -> {
                    depth++
                    when (parser.name) {
                        "extent" -> {
                            widthEmu = parser.getAttributeValue(null, "cx")?.toLongOrNull() ?: 0
                            heightEmu = parser.getAttributeValue(null, "cy")?.toLongOrNull() ?: 0
                        }
                        "docPr" -> {
                            altText = parser.getAttributeValue(null, "descr") ?: ""
                        }
                        "blip" -> {
                            relId = parser.getAttributeValue(NS_R, "embed")
                                ?: parser.getAttr("embed")
                        }
                    }
                }
                XmlPullParser.END_TAG -> depth--
            }
        }

        return if (relId != null) {
            DocxImage(relId!!, widthEmu, heightEmu, altText)
        } else null
    }

    // --- Table parsing ---

    private fun parseTable(
        parser: XmlPullParser,
        styles: Map<String, DocxStyle>,
        relationships: Map<String, String>,
    ): DocxTable {
        val rows = mutableListOf<DocxTableRow>()
        var tableProps = TableProperties()
        val gridColWidths = mutableListOf<Int>()
        var depth = 1

        while (depth > 0 && parser.next() != XmlPullParser.END_DOCUMENT) {
            when (parser.eventType) {
                XmlPullParser.START_TAG -> {
                    depth++
                    when (parser.name) {
                        "tblPr" -> {
                            tableProps = parseTableProperties(parser)
                            depth--
                        }
                        "tblGrid" -> {
                            // Parse grid column definitions — stay inside tblGrid
                        }
                        "gridCol" -> {
                            parser.getAttr("w")?.toIntOrNull()?.let { gridColWidths.add(it) }
                        }
                        "tr" -> {
                            rows.add(parseTableRow(parser, styles, relationships))
                            depth--
                        }
                    }
                }
                XmlPullParser.END_TAG -> depth--
            }
        }

        return DocxTable(rows, tableProps, gridColWidths)
    }

    private fun parseTableProperties(parser: XmlPullParser): TableProperties {
        var props = TableProperties()
        var depth = 1

        while (depth > 0 && parser.next() != XmlPullParser.END_DOCUMENT) {
            when (parser.eventType) {
                XmlPullParser.START_TAG -> {
                    depth++
                    when (parser.name) {
                        "tblW" -> {
                            props = props.copy(
                                widthTwips = parser.getAttr("w")?.toIntOrNull()
                            )
                        }
                        "jc" -> {
                            val align = when (parser.getAttr("val")) {
                                "center" -> ParagraphAlignment.CENTER
                                "right", "end" -> ParagraphAlignment.RIGHT
                                else -> ParagraphAlignment.LEFT
                            }
                            props = props.copy(alignment = align)
                        }
                        "tblBorders" -> {
                            props = props.copy(borders = parseTableBorders(parser))
                            depth--
                        }
                    }
                }
                XmlPullParser.END_TAG -> depth--
            }
        }

        return props
    }

    private fun parseTableBorders(parser: XmlPullParser): TableBorders {
        var borders = TableBorders()
        var depth = 1

        while (depth > 0 && parser.next() != XmlPullParser.END_DOCUMENT) {
            when (parser.eventType) {
                XmlPullParser.START_TAG -> {
                    depth++
                    val border = parseSingleBorder(parser)
                    when (parser.name) {
                        "top" -> borders = borders.copy(top = border)
                        "bottom" -> borders = borders.copy(bottom = border)
                        "left", "start" -> borders = borders.copy(left = border)
                        "right", "end" -> borders = borders.copy(right = border)
                        "insideH" -> borders = borders.copy(insideH = border)
                        "insideV" -> borders = borders.copy(insideV = border)
                    }
                }
                XmlPullParser.END_TAG -> depth--
            }
        }

        return borders
    }

    private fun parseSingleBorder(parser: XmlPullParser): CellBorder {
        val valStr = parser.getAttr("val") ?: "single"
        val style = when (valStr) {
            "none", "nil" -> BorderStyle.NONE
            "double" -> BorderStyle.DOUBLE
            "dashed", "dashDotStroked" -> BorderStyle.DASHED
            "dotted" -> BorderStyle.DOTTED
            "dashSmallGap" -> BorderStyle.DASH_SMALL_GAP
            else -> BorderStyle.SINGLE
        }
        val sz = parser.getAttr("sz")?.toIntOrNull() ?: 4
        val color = DocxColors.parseColor(parser.getAttr("color"))
        return CellBorder(style, sz, color)
    }

    private fun parseTableRow(
        parser: XmlPullParser,
        styles: Map<String, DocxStyle>,
        relationships: Map<String, String>,
    ): DocxTableRow {
        val cells = mutableListOf<DocxTableCell>()
        var depth = 1

        while (depth > 0 && parser.next() != XmlPullParser.END_DOCUMENT) {
            when (parser.eventType) {
                XmlPullParser.START_TAG -> {
                    depth++
                    if (parser.name == "tc") {
                        cells.add(parseTableCell(parser, styles, relationships))
                        depth--
                    }
                }
                XmlPullParser.END_TAG -> depth--
            }
        }

        return DocxTableRow(cells)
    }

    private fun parseTableCell(
        parser: XmlPullParser,
        styles: Map<String, DocxStyle>,
        relationships: Map<String, String>,
    ): DocxTableCell {
        val paragraphs = mutableListOf<DocxParagraph>()
        var widthTwips: Int? = null
        var gridSpan = 1
        var vMerge = VMergeType.NONE
        var shading: Int? = null
        var cellBorders: CellBorders? = null
        var depth = 1

        while (depth > 0 && parser.next() != XmlPullParser.END_DOCUMENT) {
            when (parser.eventType) {
                XmlPullParser.START_TAG -> {
                    depth++
                    when (parser.name) {
                        "tcPr" -> {
                            val props = parseTableCellProperties(parser)
                            widthTwips = props.width
                            gridSpan = props.gridSpan
                            vMerge = props.vMerge
                            shading = props.shading
                            cellBorders = props.borders
                            depth--
                        }
                        "p" -> {
                            paragraphs.add(parseParagraph(parser, styles, relationships))
                            depth--
                        }
                    }
                }
                XmlPullParser.END_TAG -> depth--
            }
        }

        return DocxTableCell(paragraphs, widthTwips, gridSpan, vMerge, shading, cellBorders)
    }

    private data class CellProps(
        val width: Int?, val gridSpan: Int, val vMerge: VMergeType,
        val shading: Int?, val borders: CellBorders?,
    )

    private fun parseTableCellProperties(parser: XmlPullParser): CellProps {
        var width: Int? = null
        var gridSpan = 1
        var vMerge = VMergeType.NONE
        var shading: Int? = null
        var borders: CellBorders? = null
        var depth = 1

        while (depth > 0 && parser.next() != XmlPullParser.END_DOCUMENT) {
            when (parser.eventType) {
                XmlPullParser.START_TAG -> {
                    depth++
                    when (parser.name) {
                        "tcW" -> width = parser.getAttr("w")?.toIntOrNull()
                        "gridSpan" -> gridSpan = parser.getAttr("val")?.toIntOrNull() ?: 1
                        "vMerge" -> {
                            val val_ = parser.getAttr("val")
                            vMerge = if (val_ == "restart") VMergeType.RESTART else VMergeType.CONTINUE
                        }
                        "shd" -> {
                            shading = DocxColors.parseColor(parser.getAttr("fill"))
                        }
                        "tcBorders" -> {
                            borders = parseCellBorders(parser)
                            depth--
                        }
                    }
                }
                XmlPullParser.END_TAG -> depth--
            }
        }

        return CellProps(width, gridSpan, vMerge, shading, borders)
    }

    private fun parseCellBorders(parser: XmlPullParser): CellBorders {
        var borders = CellBorders()
        var depth = 1

        while (depth > 0 && parser.next() != XmlPullParser.END_DOCUMENT) {
            when (parser.eventType) {
                XmlPullParser.START_TAG -> {
                    depth++
                    val border = parseSingleBorder(parser)
                    when (parser.name) {
                        "top" -> borders = borders.copy(top = border)
                        "bottom" -> borders = borders.copy(bottom = border)
                        "left", "start" -> borders = borders.copy(left = border)
                        "right", "end" -> borders = borders.copy(right = border)
                    }
                }
                XmlPullParser.END_TAG -> depth--
            }
        }

        return borders
    }

    // --- Helpers ---

    private fun detectHeadingLevel(styleId: String, styles: Map<String, DocxStyle>): Int {
        // Check style ID pattern like "Heading1", "heading2"
        val match = Regex("(?i)heading(\\d)").find(styleId)
        if (match != null) return match.groupValues[1].toIntOrNull() ?: 0

        // Check style name
        val style = styles[styleId]
        if (style != null) {
            val nameMatch = Regex("(?i)heading\\s*(\\d)").find(style.name)
            if (nameMatch != null) return nameMatch.groupValues[1].toIntOrNull() ?: 0
        }

        return 0
    }

    private fun applyStyleToParaProps(
        props: ParagraphProperties,
        styles: Map<String, DocxStyle>,
    ): ParagraphProperties {
        val styleId = props.styleId ?: return props

        // Resolve style chain (basedOn inheritance)
        val resolvedStyleProps = resolveStyleChain(styleId, styles)

        return ParagraphProperties(
            alignment = if (props.alignment != ParagraphAlignment.LEFT) props.alignment else resolvedStyleProps.alignment,
            spacingBeforePt = if (props.spacingBeforeExplicit) props.spacingBeforePt else resolvedStyleProps.spacingBeforePt,
            spacingBeforeExplicit = props.spacingBeforeExplicit || resolvedStyleProps.spacingBeforeExplicit,
            spacingAfterPt = if (props.spacingAfterExplicit) props.spacingAfterPt else resolvedStyleProps.spacingAfterPt,
            spacingAfterExplicit = props.spacingAfterExplicit || resolvedStyleProps.spacingAfterExplicit,
            lineSpacingMultiplier = if (props.lineSpacingExplicitlySet) props.lineSpacingMultiplier else resolvedStyleProps.lineSpacingMultiplier,
            lineSpacingPt = if (props.lineSpacingExplicitlySet) props.lineSpacingPt else resolvedStyleProps.lineSpacingPt,
            lineSpacingExact = if (props.lineSpacingExplicitlySet) props.lineSpacingExact else resolvedStyleProps.lineSpacingExact,
            lineSpacingExplicitlySet = props.lineSpacingExplicitlySet || resolvedStyleProps.lineSpacingExplicitlySet,
            indentLeftTwips = if (props.indentLeftTwips > 0) props.indentLeftTwips else resolvedStyleProps.indentLeftTwips,
            indentRightTwips = if (props.indentRightTwips > 0) props.indentRightTwips else resolvedStyleProps.indentRightTwips,
            indentFirstLineTwips = if (props.indentFirstLineTwips != 0) props.indentFirstLineTwips else resolvedStyleProps.indentFirstLineTwips,
            headingLevel = if (props.headingLevel > 0) props.headingLevel else resolvedStyleProps.headingLevel,
            styleId = props.styleId,
            contextualSpacing = props.contextualSpacing || resolvedStyleProps.contextualSpacing,
            keepNext = props.keepNext || resolvedStyleProps.keepNext,
            keepLines = props.keepLines || resolvedStyleProps.keepLines,
            pageBreakBefore = props.pageBreakBefore || resolvedStyleProps.pageBreakBefore,
        )
    }

    /**
     * Resolves a style's paragraph properties by walking the basedOn chain.
     * Deeper ancestors are applied first, then overridden by child styles.
     */
    private fun resolveStyleChain(
        styleId: String,
        styles: Map<String, DocxStyle>,
        visited: MutableSet<String> = mutableSetOf(),
    ): ParagraphProperties {
        if (styleId in visited) return ParagraphProperties() // cycle guard
        visited.add(styleId)
        val style = styles[styleId] ?: return ParagraphProperties()

        // Start with base style's resolved properties (if any)
        val base = style.basedOn?.let { resolveStyleChain(it, styles, visited) }
            ?: ParagraphProperties()
        val s = style.paragraphProperties

        // Override base with this style's explicit values
        return ParagraphProperties(
            alignment = if (s.alignment != ParagraphAlignment.LEFT) s.alignment else base.alignment,
            spacingBeforePt = if (s.spacingBeforeExplicit) s.spacingBeforePt else base.spacingBeforePt,
            spacingBeforeExplicit = s.spacingBeforeExplicit || base.spacingBeforeExplicit,
            spacingAfterPt = if (s.spacingAfterExplicit) s.spacingAfterPt else base.spacingAfterPt,
            spacingAfterExplicit = s.spacingAfterExplicit || base.spacingAfterExplicit,
            lineSpacingMultiplier = if (s.lineSpacingExplicitlySet) s.lineSpacingMultiplier else base.lineSpacingMultiplier,
            lineSpacingPt = if (s.lineSpacingExplicitlySet) s.lineSpacingPt else base.lineSpacingPt,
            lineSpacingExact = if (s.lineSpacingExplicitlySet) s.lineSpacingExact else base.lineSpacingExact,
            lineSpacingExplicitlySet = s.lineSpacingExplicitlySet || base.lineSpacingExplicitlySet,
            indentLeftTwips = if (s.indentLeftTwips > 0) s.indentLeftTwips else base.indentLeftTwips,
            indentRightTwips = if (s.indentRightTwips > 0) s.indentRightTwips else base.indentRightTwips,
            indentFirstLineTwips = if (s.indentFirstLineTwips != 0) s.indentFirstLineTwips else base.indentFirstLineTwips,
            headingLevel = if (s.headingLevel > 0) s.headingLevel else base.headingLevel,
            styleId = styleId,
            contextualSpacing = s.contextualSpacing || base.contextualSpacing,
            keepNext = s.keepNext || base.keepNext,
            keepLines = s.keepLines || base.keepLines,
            pageBreakBefore = s.pageBreakBefore || base.pageBreakBefore,
        )
    }

    private fun mergeRunProps(base: RunProperties, override: RunProperties): RunProperties {
        return RunProperties(
            bold = override.bold || base.bold,
            italic = override.italic || base.italic,
            underline = override.underline || base.underline,
            strikethrough = override.strikethrough || base.strikethrough,
            fontName = override.fontName ?: base.fontName,
            fontSizeHalfPt = override.fontSizeHalfPt ?: base.fontSizeHalfPt,
            color = override.color ?: base.color,
            highlight = override.highlight ?: base.highlight,
            superscript = override.superscript || base.superscript,
            subscript = override.subscript || base.subscript,
        )
    }

    private fun createParser(bytes: ByteArray): XmlPullParser {
        val factory = XmlPullParserFactory.newInstance()
        factory.isNamespaceAware = true
        val parser = factory.newPullParser()
        parser.setInput(ByteArrayInputStream(bytes), "UTF-8")
        return parser
    }

    /** Helper to get attribute value ignoring namespace. */
    private fun XmlPullParser.getAttr(name: String): String? {
        // Try with wordprocessingml namespace first, then null namespace
        return getAttributeValue(NS_W, name)
            ?: getAttributeValue(null, name)
    }
}
