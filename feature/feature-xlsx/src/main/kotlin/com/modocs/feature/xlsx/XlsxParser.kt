package com.modocs.feature.xlsx

import android.content.Context
import android.net.Uri
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.withContext
import com.modocs.core.common.readZipEntriesCapped
import org.xmlpull.v1.XmlPullParser
import org.xmlpull.v1.XmlPullParserFactory
import java.io.ByteArrayInputStream
import java.io.InputStream

/**
 * Parses XLSX files (ZIP archives containing SpreadsheetML) into [XlsxDocument].
 *
 * An XLSX ZIP typically contains:
 * - xl/workbook.xml — workbook with sheet definitions
 * - xl/worksheets/sheet1.xml, sheet2.xml, etc. — sheet data
 * - xl/sharedStrings.xml — shared string table
 * - xl/styles.xml — cell styles, fonts, fills, number formats
 * - xl/_rels/workbook.xml.rels — relationships mapping rIds to sheet files
 */
object XlsxParser {

    private const val NS_SS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    private const val NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

    suspend fun parse(context: Context, uri: Uri): XlsxDocument = withContext(Dispatchers.IO) {
        val inputStream = context.contentResolver.openInputStream(uri)
            ?: throw IllegalStateException("Cannot open file")
        parse(inputStream)
    }

    fun parse(inputStream: InputStream): XlsxDocument {
        // Step 1: Read all ZIP entries (size-capped against zip bombs)
        val entries = readZipEntriesCapped(inputStream)

        // Step 2: Parse workbook relationships
        val rels = parseRelationships(entries["xl/_rels/workbook.xml.rels"])

        // Step 3: Parse shared strings
        val sharedStrings = entries["xl/sharedStrings.xml"]?.let { parseSharedStrings(it) }
            ?: emptyList()

        // Step 4: Parse styles
        val styles = entries["xl/styles.xml"]?.let { parseStyles(it) } ?: emptyList()

        // Step 5: Parse workbook (sheet names and rIds)
        val sheetInfos = entries["xl/workbook.xml"]?.let { parseWorkbook(it) } ?: emptyList()

        // Step 6: Parse each sheet, tracking ZIP paths
        val sheetPathMap = mutableMapOf<Int, String>()
        val sheets = mutableListOf<XlsxSheet>()
        var sheetIndex = 0
        for (info in sheetInfos) {
            val target = rels[info.rId] ?: continue
            val sheetPath = if (target.startsWith("/")) target.removePrefix("/")
            else "xl/$target"
            val sheetBytes = entries[sheetPath] ?: continue
            val sheet = parseSheet(info.name, sheetBytes, sharedStrings, styles)
            sheets.add(sheet)
            sheetPathMap[sheetIndex] = sheetPath
            sheetIndex++
        }

        return XlsxDocument(
            sheets = sheets,
            styles = styles,
            rawEntries = entries,
            sheetPaths = sheetPathMap,
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

    // --- Shared Strings ---

    private fun parseSharedStrings(bytes: ByteArray): List<String> {
        val strings = mutableListOf<String>()
        val parser = createParser(bytes)
        var inSi = false
        var currentText = StringBuilder()

        while (parser.next() != XmlPullParser.END_DOCUMENT) {
            when (parser.eventType) {
                XmlPullParser.START_TAG -> {
                    when (parser.name) {
                        "si" -> {
                            inSi = true
                            currentText = StringBuilder()
                        }
                    }
                }
                XmlPullParser.TEXT -> {
                    if (inSi) {
                        currentText.append(parser.text)
                    }
                }
                XmlPullParser.END_TAG -> {
                    when (parser.name) {
                        "si" -> {
                            strings.add(currentText.toString())
                            inSi = false
                        }
                    }
                }
            }
        }
        return strings
    }

    // --- Styles ---

    private fun parseStyles(bytes: ByteArray): List<XlsxCellStyle> {
        val parser = createParser(bytes)

        val fonts = mutableListOf<FontInfo>()
        val fills = mutableListOf<Int?>()
        val numberFormats = mutableMapOf<Int, String>()
        val cellXfs = mutableListOf<XfInfo>()

        var inFonts = false
        var inFont = false
        var inFills = false
        var inFill = false
        var inPatternFill = false
        var inCellXfs = false
        var inNumFmts = false

        var currentFont = FontInfo()
        var currentFillColor: Int? = null

        while (parser.next() != XmlPullParser.END_DOCUMENT) {
            when (parser.eventType) {
                XmlPullParser.START_TAG -> {
                    when (parser.name) {
                        "fonts" -> inFonts = true
                        "font" -> if (inFonts) { inFont = true; currentFont = FontInfo() }
                        "b" -> if (inFont) currentFont = currentFont.copy(bold = true)
                        "i" -> if (inFont) currentFont = currentFont.copy(italic = true)
                        "u" -> if (inFont) currentFont = currentFont.copy(underline = true)
                        "sz" -> if (inFont) {
                            parser.getAttributeValue(null, "val")?.toFloatOrNull()?.let {
                                currentFont = currentFont.copy(fontSize = it)
                            }
                        }
                        "color" -> if (inFont) {
                            val rgb = parser.getAttributeValue(null, "rgb")
                            val theme = parser.getAttributeValue(null, "theme")?.toIntOrNull()
                            val indexed = parser.getAttributeValue(null, "indexed")?.toIntOrNull()
                            val color = XlsxColors.parseColor(rgb)
                                ?: theme?.let { XlsxColors.themeColor(it) }
                                ?: indexed?.let { XlsxColors.indexedColor(it) }
                            if (color != null) currentFont = currentFont.copy(fontColor = color)
                        }
                        "fills" -> inFills = true
                        "fill" -> if (inFills) { inFill = true; currentFillColor = null }
                        "patternFill" -> if (inFill) {
                            inPatternFill = true
                            val patternType = parser.getAttributeValue(null, "patternType")
                            if (patternType == "none" || patternType == "gray125") {
                                currentFillColor = null
                            }
                        }
                        "fgColor" -> if (inPatternFill) {
                            val rgb = parser.getAttributeValue(null, "rgb")
                            val theme = parser.getAttributeValue(null, "theme")?.toIntOrNull()
                            val indexed = parser.getAttributeValue(null, "indexed")?.toIntOrNull()
                            currentFillColor = XlsxColors.parseColor(rgb)
                                ?: theme?.let { XlsxColors.themeColor(it) }
                                ?: indexed?.let { XlsxColors.indexedColor(it) }
                        }
                        "numFmts" -> inNumFmts = true
                        "numFmt" -> if (inNumFmts) {
                            val id = parser.getAttributeValue(null, "numFmtId")?.toIntOrNull()
                            val code = parser.getAttributeValue(null, "formatCode")
                            if (id != null && code != null) numberFormats[id] = code
                        }
                        "cellXfs" -> inCellXfs = true
                        "xf" -> if (inCellXfs) {
                            val fontId = parser.getAttributeValue(null, "fontId")?.toIntOrNull() ?: 0
                            val fillId = parser.getAttributeValue(null, "fillId")?.toIntOrNull() ?: 0
                            val numFmtId = parser.getAttributeValue(null, "numFmtId")?.toIntOrNull() ?: 0
                            val applyAlignment = parser.getAttributeValue(null, "applyAlignment")
                            cellXfs.add(XfInfo(fontId, fillId, numFmtId))
                        }
                        "alignment" -> if (inCellXfs && cellXfs.isNotEmpty()) {
                            val h = parser.getAttributeValue(null, "horizontal")
                            val v = parser.getAttributeValue(null, "vertical")
                            val wrap = parser.getAttributeValue(null, "wrapText")
                            val last = cellXfs.last()
                            cellXfs[cellXfs.lastIndex] = last.copy(
                                hAlign = h,
                                vAlign = v,
                                wrapText = wrap == "1" || wrap == "true",
                            )
                        }
                    }
                }
                XmlPullParser.END_TAG -> {
                    when (parser.name) {
                        "font" -> if (inFonts) { fonts.add(currentFont); inFont = false }
                        "fonts" -> inFonts = false
                        "fill" -> if (inFills) { fills.add(currentFillColor); inFill = false; inPatternFill = false }
                        "fills" -> inFills = false
                        "patternFill" -> inPatternFill = false
                        "cellXfs" -> inCellXfs = false
                        "numFmts" -> inNumFmts = false
                    }
                }
            }
        }

        // Combine into cell styles
        return cellXfs.map { xf ->
            val font = fonts.getOrNull(xf.fontId) ?: FontInfo()
            val fillColor = fills.getOrNull(xf.fillId)
            val fmtCode = numberFormats[xf.numFmtId]
            XlsxCellStyle(
                fontBold = font.bold,
                fontItalic = font.italic,
                fontUnderline = font.underline,
                fontSize = font.fontSize,
                fontColor = font.fontColor,
                fillColor = fillColor,
                horizontalAlignment = when (xf.hAlign) {
                    "left" -> CellAlignment.LEFT
                    "center" -> CellAlignment.CENTER
                    "right" -> CellAlignment.RIGHT
                    "fill" -> CellAlignment.FILL
                    "justify" -> CellAlignment.JUSTIFY
                    else -> CellAlignment.GENERAL
                },
                verticalAlignment = when (xf.vAlign) {
                    "top" -> CellVerticalAlignment.TOP
                    "center" -> CellVerticalAlignment.CENTER
                    else -> CellVerticalAlignment.BOTTOM
                },
                numberFormatId = xf.numFmtId,
                numberFormatCode = fmtCode,
                wrapText = xf.wrapText,
            )
        }
    }

    private data class FontInfo(
        val bold: Boolean = false,
        val italic: Boolean = false,
        val underline: Boolean = false,
        val fontSize: Float = 11f,
        val fontColor: Int? = null,
    )

    private data class XfInfo(
        val fontId: Int = 0,
        val fillId: Int = 0,
        val numFmtId: Int = 0,
        val hAlign: String? = null,
        val vAlign: String? = null,
        val wrapText: Boolean = false,
    )

    // --- Workbook ---

    private data class SheetInfo(val name: String, val rId: String)

    private fun parseWorkbook(bytes: ByteArray): List<SheetInfo> {
        val sheets = mutableListOf<SheetInfo>()
        val parser = createParser(bytes)

        while (parser.next() != XmlPullParser.END_DOCUMENT) {
            if (parser.eventType == XmlPullParser.START_TAG && parser.name == "sheet") {
                val name = parser.getAttributeValue(null, "name") ?: "Sheet"
                val rId = parser.getAttributeValue(NS_R, "id")
                    ?: parser.getAttributeValue(null, "r:id")
                    ?: continue
                sheets.add(SheetInfo(name, rId))
            }
        }
        return sheets
    }

    // --- Sheet ---

    private fun parseSheet(
        name: String,
        bytes: ByteArray,
        sharedStrings: List<String>,
        styles: List<XlsxCellStyle>,
    ): XlsxSheet {
        val parser = createParser(bytes)
        val rows = mutableListOf<XlsxRow>()
        val columnWidths = mutableMapOf<Int, Float>()
        val mergedCells = mutableListOf<CellRange>()
        var defaultRowHeight = 15f
        var frozenRows = 0
        var frozenCols = 0

        var inRow = false
        var inCell = false
        var inFormula = false
        var inValue = false
        var inInlineStr = false
        var currentRowIndex = 0
        var currentCells = mutableListOf<XlsxCell>()
        var currentRowHeight: Float? = null
        var currentCellRef = ""
        var currentCellType = ""
        var currentCellStyle = 0
        var currentValue = StringBuilder()
        var currentFormula = StringBuilder()

        while (parser.next() != XmlPullParser.END_DOCUMENT) {
            when (parser.eventType) {
                XmlPullParser.START_TAG -> {
                    when (parser.name) {
                        "sheetFormatPr" -> {
                            parser.getAttributeValue(null, "defaultRowHeight")
                                ?.toFloatOrNull()?.let { defaultRowHeight = it }
                        }
                        "col" -> {
                            val min = parser.getAttributeValue(null, "min")?.toIntOrNull()
                            val max = parser.getAttributeValue(null, "max")?.toIntOrNull()
                            val width = parser.getAttributeValue(null, "width")?.toFloatOrNull()
                            if (min != null && max != null && width != null) {
                                for (col in min..max) {
                                    columnWidths[col - 1] = width // 0-indexed
                                }
                            }
                        }
                        "pane" -> {
                            parser.getAttributeValue(null, "ySplit")?.toIntOrNull()
                                ?.let { frozenRows = it }
                            parser.getAttributeValue(null, "xSplit")?.toIntOrNull()
                                ?.let { frozenCols = it }
                        }
                        "row" -> {
                            inRow = true
                            currentRowIndex = (parser.getAttributeValue(null, "r")
                                ?.toIntOrNull() ?: (rows.size + 1)) - 1
                            currentCells = mutableListOf()
                            currentRowHeight = parser.getAttributeValue(null, "ht")
                                ?.toFloatOrNull()
                        }
                        "c" -> if (inRow) {
                            inCell = true
                            currentCellRef = parser.getAttributeValue(null, "r") ?: ""
                            currentCellType = parser.getAttributeValue(null, "t") ?: ""
                            currentCellStyle = parser.getAttributeValue(null, "s")
                                ?.toIntOrNull() ?: 0
                            currentValue = StringBuilder()
                            currentFormula = StringBuilder()
                        }
                        "v" -> if (inCell) inValue = true
                        "f" -> if (inCell) inFormula = true
                        "is" -> if (inCell) inInlineStr = true
                        "t" -> if (inInlineStr) inValue = true
                        "mergeCell" -> {
                            parser.getAttributeValue(null, "ref")?.let { ref ->
                                parseMergeRef(ref)?.let { mergedCells.add(it) }
                            }
                        }
                    }
                }
                XmlPullParser.TEXT -> {
                    if (inValue) currentValue.append(parser.text)
                    if (inFormula) currentFormula.append(parser.text)
                }
                XmlPullParser.END_TAG -> {
                    when (parser.name) {
                        "v", "t" -> inValue = false
                        "f" -> inFormula = false
                        "is" -> inInlineStr = false
                        "c" -> if (inCell) {
                            val colIndex = cellRefToColumn(currentCellRef)
                            val rawValue = currentValue.toString()
                            val resolvedValue = when (currentCellType) {
                                "s" -> {
                                    // Shared string index
                                    val idx = rawValue.toIntOrNull()
                                    if (idx != null && idx < sharedStrings.size) sharedStrings[idx]
                                    else rawValue
                                }
                                "inlineStr" -> rawValue
                                "b" -> if (rawValue == "1") "TRUE" else "FALSE"
                                "e" -> rawValue // Error
                                else -> rawValue // Number or general
                            }
                            val type = when (currentCellType) {
                                "s" -> CellType.STRING
                                "inlineStr" -> CellType.INLINE_STRING
                                "b" -> CellType.BOOLEAN
                                "e" -> CellType.ERROR
                                "n", "" -> {
                                    // Check if number format suggests a date
                                    val style = styles.getOrNull(currentCellStyle)
                                    if (style != null && isDateFormat(style.numberFormatId, style.numberFormatCode)) {
                                        CellType.DATE
                                    } else {
                                        CellType.NUMBER
                                    }
                                }
                                else -> CellType.STRING
                            }
                            val displayValue = if (type == CellType.NUMBER) {
                                formatNumber(resolvedValue)
                            } else if (type == CellType.DATE) {
                                formatDate(resolvedValue)
                            } else {
                                resolvedValue
                            }
                            currentCells.add(
                                XlsxCell(
                                    columnIndex = colIndex,
                                    value = displayValue,
                                    type = type,
                                    styleIndex = currentCellStyle,
                                    formula = currentFormula.toString().ifEmpty { null },
                                )
                            )
                            inCell = false
                        }
                        "row" -> if (inRow) {
                            rows.add(XlsxRow(currentRowIndex, currentCells, currentRowHeight))
                            inRow = false
                        }
                    }
                }
            }
        }

        return XlsxSheet(
            name = name,
            rows = rows,
            columnWidths = columnWidths,
            defaultRowHeight = defaultRowHeight,
            mergedCells = mergedCells,
            frozenColumns = frozenCols,
            frozenRows = frozenRows,
        )
    }

    // --- Helpers ---

    /** Convert cell reference like "B3" to 0-based column index (1). */
    private fun cellRefToColumn(ref: String): Int {
        var col = 0
        for (ch in ref) {
            if (ch.isLetter()) {
                col = col * 26 + (ch.uppercaseChar() - 'A' + 1)
            } else break
        }
        return col - 1
    }

    /** Parse merge reference like "A1:C3" into a CellRange. */
    private fun parseMergeRef(ref: String): CellRange? {
        val parts = ref.split(":")
        if (parts.size != 2) return null
        val (startCol, startRow) = parseCellRef(parts[0]) ?: return null
        val (endCol, endRow) = parseCellRef(parts[1]) ?: return null
        return CellRange(startRow, startCol, endRow, endCol)
    }

    /** Parse "B3" -> (colIndex=1, rowIndex=2). */
    private fun parseCellRef(ref: String): Pair<Int, Int>? {
        var col = 0
        var i = 0
        while (i < ref.length && ref[i].isLetter()) {
            col = col * 26 + (ref[i].uppercaseChar() - 'A' + 1)
            i++
        }
        val row = ref.substring(i).toIntOrNull() ?: return null
        return (col - 1) to (row - 1)
    }

    /** Check if a number format ID represents a date. */
    private fun isDateFormat(numFmtId: Int, customCode: String?): Boolean {
        // Built-in date format IDs
        if (numFmtId in 14..22 || numFmtId in 45..47) return true
        // Check custom format code for date patterns
        if (customCode != null) {
            val lower = customCode.lowercase()
            return lower.contains("yy") || lower.contains("mm") && lower.contains("dd")
        }
        return false
    }

    /** Format a numeric string (remove unnecessary decimals). */
    private fun formatNumber(value: String): String {
        val num = value.toDoubleOrNull() ?: return value
        return if (num == num.toLong().toDouble()) {
            num.toLong().toString()
        } else {
            // Keep reasonable precision
            val formatted = "%.10g".format(num)
            formatted.trimEnd('0').trimEnd('.')
        }
    }

    /** Convert Excel serial date number to a readable date string. */
    private fun formatDate(value: String): String {
        val serial = value.toDoubleOrNull() ?: return value
        if (serial < 1) return value
        // Excel epoch is 1900-01-01, with the Lotus 1-2-3 leap year bug
        val days = serial.toLong() - (if (serial >= 61) 2 else 1)
        return try {
            val epoch = java.time.LocalDate.of(1900, 1, 1).plusDays(days)
            epoch.toString()
        } catch (_: Exception) {
            value
        }
    }

    private fun createParser(bytes: ByteArray): XmlPullParser {
        val factory = XmlPullParserFactory.newInstance()
        factory.isNamespaceAware = true
        val parser = factory.newPullParser()
        parser.setInput(ByteArrayInputStream(bytes), "UTF-8")
        return parser
    }
}
