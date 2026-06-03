package com.modocs.feature.pptx

import android.content.Context
import android.graphics.Color as AndroidColor
import android.net.Uri
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.withContext
import com.modocs.core.common.readZipEntriesCapped
import org.xmlpull.v1.XmlPullParser
import org.xmlpull.v1.XmlPullParserFactory
import java.io.ByteArrayInputStream
import java.io.InputStream
import kotlin.math.roundToInt

/**
 * Parses PPTX files (ZIP archives containing PresentationML) into [PptxDocument].
 *
 * Handles:
 * - Theme color resolution (schemeClr with dk1, lt1, accent1-6, etc.)
 * - Color modifiers (lumMod, lumOff, tint, shade, alpha)
 * - Slide master/layout background inheritance
 * - Group shapes (grpSp) with coordinate transforms
 * - Default run properties (defRPr)
 * - Placeholder dimension resolution from layout/master
 * - Layout/master decorative shapes as background
 * - Gradient fill (first stop color)
 */
object PptxParser {

    private const val NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
    private const val NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    private const val NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"

    private const val REL_TYPE_SLIDE_LAYOUT =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
    private const val REL_TYPE_SLIDE_MASTER =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"

    suspend fun parse(context: Context, uri: Uri): PptxDocument = withContext(Dispatchers.IO) {
        val inputStream = context.contentResolver.openInputStream(uri)
            ?: throw IllegalStateException("Cannot open file")
        parse(inputStream)
    }

    fun parse(inputStream: InputStream): PptxDocument {
        // Step 1: Read all ZIP entries (size-capped against zip bombs)
        val entries = readZipEntriesCapped(inputStream)

        // Step 2: Parse theme colors
        val themeColors = parseThemeColors(entries)

        // Step 3: Parse presentation relationships
        val presRels = parseRelationships(entries["ppt/_rels/presentation.xml.rels"])

        // Step 4: Parse presentation.xml (slide size + slide order)
        val presBytes = entries["ppt/presentation.xml"]
            ?: throw IllegalStateException("Missing ppt/presentation.xml")
        val presInfo = parsePresentation(presBytes)

        // Step 5: Parse each slide
        val slides = presInfo.slideRIds.mapIndexedNotNull { index, rId ->
            val target = presRels[rId] ?: return@mapIndexedNotNull null
            val slidePath = if (target.startsWith("/")) target.removePrefix("/")
            else "ppt/$target"
            val slideBytes = entries[slidePath] ?: return@mapIndexedNotNull null

            // Parse per-slide relationships (for images + layout reference)
            val slideFileName = slidePath.substringAfterLast("/")
            val slideRelsPath = slidePath.replace(slideFileName, "_rels/$slideFileName.rels")
            val slideRels = parseRelationships(entries[slideRelsPath])
            val slideRelsTyped = parseRelationshipsWithType(entries[slideRelsPath])

            // Resolve images
            val images = mutableMapOf<String, ByteArray>()
            for ((relId, relTarget) in slideRels) {
                val imagePath = if (relTarget.startsWith("/")) relTarget.removePrefix("/")
                else {
                    val slideDir = slidePath.substringBeforeLast("/")
                    "$slideDir/$relTarget".replace("/./", "/")
                }
                val normalizedPath = normalizePath(imagePath)
                entries[normalizedPath]?.let { images[relId] = it }
            }

            // Resolve layout path
            val layoutRelInfo = slideRelsTyped.values.firstOrNull { it.type == REL_TYPE_SLIDE_LAYOUT }
            val layoutPath = layoutRelInfo?.let { rel ->
                val slideDir = slidePath.substringBeforeLast("/")
                if (rel.target.startsWith("/")) rel.target.removePrefix("/")
                else normalizePath("$slideDir/${rel.target}")
            }

            // Resolve master path from layout rels
            var masterPath: String? = null
            if (layoutPath != null) {
                val layoutFileName = layoutPath.substringAfterLast("/")
                val layoutRelsPath = layoutPath.replace(layoutFileName, "_rels/$layoutFileName.rels")
                val layoutRelsTyped = parseRelationshipsWithType(entries[layoutRelsPath])
                val masterRel = layoutRelsTyped.values.firstOrNull { it.type == REL_TYPE_SLIDE_MASTER }
                masterPath = masterRel?.let { rel ->
                    val layoutDir = layoutPath.substringBeforeLast("/")
                    if (rel.target.startsWith("/")) rel.target.removePrefix("/")
                    else normalizePath("$layoutDir/${rel.target}")
                }
            }

            // Parse master layout info (placeholders + decorative shapes)
            val masterInfo = if (masterPath != null) {
                entries[masterPath]?.let { parseLayoutInfo(it, themeColors) }
            } else null

            // Parse slide layout info (placeholders + decorative shapes)
            val layoutInfo = if (layoutPath != null) {
                entries[layoutPath]?.let { parseLayoutInfo(it, themeColors) }
            } else null

            // Merge placeholders: layout overrides master
            val masterPlaceholders = masterInfo?.placeholders ?: emptyList()
            val layoutPlaceholders = layoutInfo?.placeholders ?: emptyList()
            val mergedPlaceholders = mergePlaceholders(masterPlaceholders, layoutPlaceholders)

            // Collect background shapes: master shapes first, then layout shapes
            val bgShapes = mutableListOf<PptxShape>()
            masterInfo?.decorativeShapes?.let { bgShapes.addAll(it) }
            layoutInfo?.decorativeShapes?.let { bgShapes.addAll(it) }

            // Resolve background color from master -> layout -> slide chain
            var bgColor: Int? = null
            if (masterPath != null) {
                entries[masterPath]?.let { extractBackgroundColor(it, themeColors)?.let { c -> bgColor = c } }
            }
            if (layoutPath != null) {
                entries[layoutPath]?.let { extractBackgroundColor(it, themeColors)?.let { c -> bgColor = c } }
            }

            parseSlide(
                slideNumber = index + 1,
                bytes = slideBytes,
                images = images,
                themeColors = themeColors,
                inheritedBgColor = bgColor,
                layoutPlaceholders = mergedPlaceholders,
                backgroundShapes = bgShapes,
                slideWidth = presInfo.slideWidth,
                slideHeight = presInfo.slideHeight,
            )
        }

        return PptxDocument(
            slides = slides,
            slideWidth = presInfo.slideWidth,
            slideHeight = presInfo.slideHeight,
            rawEntries = entries,
            themeColors = themeColors,
        )
    }

    // ==================== Layout/Master Info ====================

    /**
     * Placeholder geometry from a layout or master.
     */
    private data class LayoutPlaceholder(
        val type: String?,
        val idx: Int?,
        val x: Long, val y: Long,
        val w: Long, val h: Long,
        /** Bullet character per paragraph level, inherited from lstStyle. */
        val bulletsByLevel: Map<Int, String> = emptyMap(),
    )

    private data class LayoutInfo(
        val placeholders: List<LayoutPlaceholder>,
        val decorativeShapes: List<PptxShape>,
    )

    /**
     * Parse a layout or master XML to extract placeholder positions and non-placeholder shapes.
     */
    private fun parseLayoutInfo(
        bytes: ByteArray,
        themeColors: Map<String, Int>,
    ): LayoutInfo {
        val placeholders = mutableListOf<LayoutPlaceholder>()
        val shapes = mutableListOf<PptxShape>()

        val parser = createParser(bytes)
        var inSpTree = false

        while (parser.next() != XmlPullParser.END_DOCUMENT) {
            when (parser.eventType) {
                XmlPullParser.START_TAG -> when (parser.name) {
                    "spTree" -> inSpTree = true
                    "sp" -> if (inSpTree) {
                        parseLayoutSp(parser, placeholders, shapes, themeColors)
                    }
                    // Skip group shapes in layout for simplicity — they're uncommon
                }
                XmlPullParser.END_TAG -> when (parser.name) {
                    "spTree" -> return LayoutInfo(placeholders, shapes)
                }
            }
        }

        return LayoutInfo(placeholders, shapes)
    }

    /**
     * Parse a single shape from a layout/master XML.
     * If it has <p:ph>, record as placeholder. Otherwise, add as decorative shape.
     */
    private fun parseLayoutSp(
        parser: XmlPullParser,
        placeholders: MutableList<LayoutPlaceholder>,
        shapes: MutableList<PptxShape>,
        themeColors: Map<String, Int>,
    ) {
        val startDepth = parser.depth
        var phType: String? = null
        var phIdx: Int? = null
        var hasPh = false
        var x = 0L; var y = 0L; var w = 0L; var h = 0L
        var fillColor: Int? = null
        var borderColor: Int? = null
        var borderWidth = 0L
        val paragraphs = mutableListOf<PptxParagraph>()
        val bulletsByLevel = mutableMapOf<Int, String>()
        var inSpPr = false
        var inXfrm = false
        var inLn = false
        var inTxBody = false

        while (parser.next() != XmlPullParser.END_DOCUMENT) {
            when (parser.eventType) {
                XmlPullParser.START_TAG -> when (parser.name) {
                    "ph" -> {
                        hasPh = true
                        phType = parser.getAttributeValue(null, "type")
                        phIdx = parser.getAttributeValue(null, "idx")?.toIntOrNull()
                    }
                    "spPr" -> inSpPr = true
                    "xfrm" -> if (inSpPr) inXfrm = true
                    "off" -> if (inXfrm && !inLn) {
                        x = parser.getAttributeValue(null, "x")?.toLongOrNull() ?: 0L
                        y = parser.getAttributeValue(null, "y")?.toLongOrNull() ?: 0L
                    }
                    "ext" -> if (inXfrm && !inLn && !inTxBody) {
                        w = parser.getAttributeValue(null, "cx")?.toLongOrNull() ?: 0L
                        h = parser.getAttributeValue(null, "cy")?.toLongOrNull() ?: 0L
                    }
                    "ln" -> if (inSpPr && !inTxBody) {
                        inLn = true
                        borderWidth = parser.getAttributeValue(null, "w")?.toLongOrNull() ?: 0L
                    }
                    "solidFill" -> {
                        if (inSpPr && !inLn && !inTxBody) {
                            fillColor = readSolidFillColor(parser, themeColors)
                        } else if (inLn && !inTxBody) {
                            borderColor = readSolidFillColor(parser, themeColors)
                        }
                    }
                    "txBody" -> {
                        inTxBody = true
                        parseTxBody(parser, paragraphs, themeColors, bulletsByLevel)
                        inTxBody = false
                    }
                }
                XmlPullParser.END_TAG -> {
                    when (parser.name) {
                        "spPr" -> { inSpPr = false; inXfrm = false; inLn = false }
                        "xfrm" -> inXfrm = false
                        "ln" -> inLn = false
                    }
                    if (parser.depth <= startDepth) {
                        if (hasPh) {
                            // Placeholder — record position and bullet styles
                            if (w > 0 || h > 0) {
                                placeholders.add(LayoutPlaceholder(phType, phIdx, x, y, w, h, bulletsByLevel))
                            }
                        } else if (w > 0 && h > 0) {
                            // Non-placeholder — decorative shape
                            if (paragraphs.isNotEmpty()) {
                                shapes.add(
                                    PptxTextBox(x, y, w, h, paragraphs.toList(), fillColor, borderColor, borderWidth)
                                )
                            } else if (fillColor != null || borderColor != null) {
                                shapes.add(
                                    PptxRectangle(x, y, w, h, fillColor, borderColor, borderWidth)
                                )
                            }
                        }
                        return
                    }
                }
            }
        }
    }

    /**
     * Parse a lstStyle element to extract bullet characters per level.
     * Parser is at the lstStyle START_TAG. Consumes through END_TAG.
     */
    private fun parseLstStyle(parser: XmlPullParser, bullets: MutableMap<Int, String>) {
        val startDepth = parser.depth

        while (parser.next() != XmlPullParser.END_DOCUMENT) {
            when (parser.eventType) {
                XmlPullParser.START_TAG -> {
                    val name = parser.name
                    // Match lvl1pPr through lvl9pPr
                    val match = Regex("lvl(\\d)pPr").matchEntire(name)
                    if (match != null) {
                        val level = match.groupValues[1].toInt() - 1 // 0-based
                        parseLevelBullet(parser, level, bullets)
                    }
                }
                XmlPullParser.END_TAG -> {
                    if (parser.depth <= startDepth) return
                }
            }
        }
    }

    /**
     * Parse a level paragraph properties element for bullet definition.
     * Parser is at the lvlNpPr START_TAG. Consumes through END_TAG.
     */
    private fun parseLevelBullet(parser: XmlPullParser, level: Int, bullets: MutableMap<Int, String>) {
        val startDepth = parser.depth

        while (parser.next() != XmlPullParser.END_DOCUMENT) {
            when (parser.eventType) {
                XmlPullParser.START_TAG -> {
                    when (parser.name) {
                        "buChar" -> {
                            parser.getAttributeValue(null, "char")?.let { bullets[level] = it }
                        }
                        "buAutoNum" -> {
                            // Auto-numbered bullets — use a default bullet character
                            if (level !in bullets) bullets[level] = "\u2022"
                        }
                    }
                }
                XmlPullParser.END_TAG -> {
                    if (parser.depth <= startDepth) return
                }
            }
        }
    }

    /**
     * Merge placeholder lists: layout overrides master by type or idx.
     */
    private fun mergePlaceholders(
        master: List<LayoutPlaceholder>,
        layout: List<LayoutPlaceholder>,
    ): List<LayoutPlaceholder> {
        val merged = master.toMutableList()
        for (lph in layout) {
            merged.removeAll { mph ->
                (lph.idx != null && mph.idx != null && lph.idx == mph.idx) ||
                    (lph.type != null && mph.type != null && lph.type == mph.type)
            }
            merged.add(lph)
        }
        return merged
    }

    /**
     * Find a matching layout placeholder for a slide shape's placeholder info.
     */
    private fun findLayoutPlaceholder(
        type: String?,
        idx: Int?,
        layoutPlaceholders: List<LayoutPlaceholder>,
    ): LayoutPlaceholder? {
        // 1. Match by idx (most specific)
        if (idx != null) {
            val byIdx = layoutPlaceholders.find { it.idx == idx }
            if (byIdx != null) return byIdx
        }
        // 2. Match by type
        if (type != null) {
            val byType = layoutPlaceholders.find { it.type == type }
            if (byType != null) return byType
        }
        // 3. Default: if ph has no type/idx, try body placeholder
        if (type == null && idx == null) {
            return layoutPlaceholders.find { it.type == "body" }
        }
        return null
    }

    // ==================== Theme Colors ====================

    private fun parseThemeColors(entries: Map<String, ByteArray>): Map<String, Int> {
        val themeKey = entries.keys.firstOrNull { it.matches(Regex("ppt/theme/theme\\d*\\.xml")) }
            ?: return emptyMap()
        val themeBytes = entries[themeKey] ?: return emptyMap()

        val colors = mutableMapOf<String, Int>()
        val parser = createParser(themeBytes)

        val schemeElements = setOf(
            "dk1", "lt1", "dk2", "lt2",
            "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
            "hlink", "folHlink",
        )

        var inClrScheme = false
        var currentSchemeElement: String? = null

        while (parser.next() != XmlPullParser.END_DOCUMENT) {
            when (parser.eventType) {
                XmlPullParser.START_TAG -> {
                    val tag = parser.name
                    when {
                        tag == "clrScheme" -> inClrScheme = true
                        inClrScheme && tag in schemeElements -> {
                            currentSchemeElement = tag
                        }
                        currentSchemeElement != null && tag == "srgbClr" -> {
                            val hex = parser.getAttributeValue(null, "val")
                            if (hex != null) {
                                PptxColors.parseColor(hex)?.let { colors[currentSchemeElement!!] = it }
                            }
                        }
                        currentSchemeElement != null && tag == "sysClr" -> {
                            val hex = parser.getAttributeValue(null, "lastClr")
                            if (hex != null) {
                                PptxColors.parseColor(hex)?.let { colors[currentSchemeElement!!] = it }
                            }
                        }
                    }
                }
                XmlPullParser.END_TAG -> {
                    val tag = parser.name
                    when {
                        tag == "clrScheme" -> {
                            inClrScheme = false
                            currentSchemeElement = null
                        }
                        inClrScheme && tag in schemeElements -> {
                            currentSchemeElement = null
                        }
                    }
                }
            }
        }

        return colors
    }

    // ==================== Color Resolution ====================

    private fun resolveColorFromElement(
        parser: XmlPullParser,
        themeColors: Map<String, Int>,
    ): Int? {
        val parentDepth = parser.depth
        var baseColor: Int? = null
        var lumMod: Float? = null
        var lumOff: Float? = null
        var tint: Float? = null
        var shade: Float? = null
        var alphaVal: Float? = null

        while (parser.next() != XmlPullParser.END_DOCUMENT) {
            when (parser.eventType) {
                XmlPullParser.START_TAG -> {
                    when (parser.name) {
                        "srgbClr" -> {
                            baseColor = PptxColors.parseColor(parser.getAttributeValue(null, "val"))
                            val mods = readColorModifiers(parser)
                            lumMod = mods.lumMod
                            lumOff = mods.lumOff
                            tint = mods.tint
                            shade = mods.shade
                            alphaVal = mods.alpha
                        }
                        "schemeClr" -> {
                            val schemeName = parser.getAttributeValue(null, "val")
                            baseColor = resolveSchemeColor(schemeName, themeColors)
                            val mods = readColorModifiers(parser)
                            lumMod = mods.lumMod
                            lumOff = mods.lumOff
                            tint = mods.tint
                            shade = mods.shade
                            alphaVal = mods.alpha
                        }
                        "sysClr" -> {
                            baseColor = PptxColors.parseColor(
                                parser.getAttributeValue(null, "lastClr")
                            )
                            val mods = readColorModifiers(parser)
                            lumMod = mods.lumMod
                            lumOff = mods.lumOff
                            tint = mods.tint
                            shade = mods.shade
                            alphaVal = mods.alpha
                        }
                    }
                }
                XmlPullParser.END_TAG -> {
                    if (parser.depth <= parentDepth) break
                }
            }
        }

        return baseColor?.let { applyColorModifiers(it, lumMod, lumOff, tint, shade, alphaVal) }
    }

    private data class ColorModifiers(
        val lumMod: Float? = null,
        val lumOff: Float? = null,
        val tint: Float? = null,
        val shade: Float? = null,
        val alpha: Float? = null,
    )

    private fun readColorModifiers(parser: XmlPullParser): ColorModifiers {
        val depth = parser.depth
        var lumMod: Float? = null
        var lumOff: Float? = null
        var tint: Float? = null
        var shade: Float? = null
        var alpha: Float? = null

        while (parser.next() != XmlPullParser.END_DOCUMENT) {
            when (parser.eventType) {
                XmlPullParser.START_TAG -> {
                    val valAttr = parser.getAttributeValue(null, "val")?.toFloatOrNull()
                    when (parser.name) {
                        "lumMod" -> lumMod = valAttr?.let { it / 100000f }
                        "lumOff" -> lumOff = valAttr?.let { it / 100000f }
                        "tint" -> tint = valAttr?.let { it / 100000f }
                        "shade" -> shade = valAttr?.let { it / 100000f }
                        "alpha" -> alpha = valAttr?.let { it / 100000f }
                    }
                }
                XmlPullParser.END_TAG -> {
                    if (parser.depth <= depth) break
                }
            }
        }

        return ColorModifiers(lumMod, lumOff, tint, shade, alpha)
    }

    private fun resolveSchemeColor(schemeName: String?, themeColors: Map<String, Int>): Int? {
        if (schemeName == null) return null
        val mappedName = when (schemeName) {
            "tx1" -> "dk1"
            "tx2" -> "dk2"
            "bg1" -> "lt1"
            "bg2" -> "lt2"
            else -> schemeName
        }
        return themeColors[mappedName]
    }

    private fun applyColorModifiers(
        baseColor: Int,
        lumMod: Float?,
        lumOff: Float?,
        tint: Float?,
        shade: Float?,
        alphaVal: Float?,
    ): Int {
        var r = AndroidColor.red(baseColor) / 255f
        var g = AndroidColor.green(baseColor) / 255f
        var b = AndroidColor.blue(baseColor) / 255f
        var a = AndroidColor.alpha(baseColor) / 255f

        if (tint != null) {
            r = r + (1f - r) * (1f - tint)
            g = g + (1f - g) * (1f - tint)
            b = b + (1f - b) * (1f - tint)
        }

        if (shade != null) {
            r *= shade
            g *= shade
            b *= shade
        }

        if (lumMod != null || lumOff != null) {
            val hsl = rgbToHsl(r, g, b)
            var l = hsl[2]
            if (lumMod != null) l *= lumMod
            if (lumOff != null) l += lumOff
            l = l.coerceIn(0f, 1f)
            hsl[2] = l
            val rgb = hslToRgb(hsl[0], hsl[1], hsl[2])
            r = rgb[0]
            g = rgb[1]
            b = rgb[2]
        }

        if (alphaVal != null) {
            a = alphaVal
        }

        return AndroidColor.argb(
            (a * 255).roundToInt().coerceIn(0, 255),
            (r * 255).roundToInt().coerceIn(0, 255),
            (g * 255).roundToInt().coerceIn(0, 255),
            (b * 255).roundToInt().coerceIn(0, 255),
        )
    }

    private fun readSolidFillColor(
        parser: XmlPullParser,
        themeColors: Map<String, Int>,
    ): Int? {
        return resolveColorFromElement(parser, themeColors)
    }

    private fun readRunColor(
        parser: XmlPullParser,
        themeColors: Map<String, Int>,
    ): Int? {
        return resolveColorFromElement(parser, themeColors)
    }

    // ==================== HSL Conversions ====================

    private fun rgbToHsl(r: Float, g: Float, b: Float): FloatArray {
        val max = maxOf(r, g, b)
        val min = minOf(r, g, b)
        val l = (max + min) / 2f
        if (max == min) return floatArrayOf(0f, 0f, l)

        val d = max - min
        val s = if (l > 0.5f) d / (2f - max - min) else d / (max + min)
        val h = when (max) {
            r -> ((g - b) / d + (if (g < b) 6f else 0f)) / 6f
            g -> ((b - r) / d + 2f) / 6f
            else -> ((r - g) / d + 4f) / 6f
        }
        return floatArrayOf(h, s, l)
    }

    private fun hslToRgb(h: Float, s: Float, l: Float): FloatArray {
        if (s == 0f) return floatArrayOf(l, l, l)

        val q = if (l < 0.5f) l * (1f + s) else l + s - l * s
        val p = 2f * l - q
        return floatArrayOf(
            hueToRgb(p, q, h + 1f / 3f),
            hueToRgb(p, q, h),
            hueToRgb(p, q, h - 1f / 3f),
        )
    }

    private fun hueToRgb(p: Float, q: Float, t0: Float): Float {
        var t = t0
        if (t < 0f) t += 1f
        if (t > 1f) t -= 1f
        return when {
            t < 1f / 6f -> p + (q - p) * 6f * t
            t < 1f / 2f -> q
            t < 2f / 3f -> p + (q - p) * (2f / 3f - t) * 6f
            else -> p
        }
    }

    // ==================== Background Extraction ====================

    private fun extractBackgroundColor(
        bytes: ByteArray,
        themeColors: Map<String, Int>,
    ): Int? {
        val parser = createParser(bytes)
        var inBg = false
        var inBgPr = false
        var inBgRef = false

        while (parser.next() != XmlPullParser.END_DOCUMENT) {
            when (parser.eventType) {
                XmlPullParser.START_TAG -> {
                    when (parser.name) {
                        "bg" -> inBg = true
                        "bgPr" -> if (inBg) inBgPr = true
                        "bgRef" -> if (inBg) {
                            inBgRef = true
                        }
                        "solidFill" -> {
                            if (inBgPr) {
                                return resolveColorFromElement(parser, themeColors)
                            }
                        }
                        "gradFill" -> {
                            // Use first gradient stop color as approximate background
                            if (inBgPr) {
                                val color = extractFirstGradientStop(parser, themeColors)
                                if (color != null) return color
                            }
                        }
                        "srgbClr" -> {
                            if (inBgRef) {
                                val color = PptxColors.parseColor(parser.getAttributeValue(null, "val"))
                                if (color != null) return color
                            }
                        }
                        "schemeClr" -> {
                            if (inBgRef) {
                                val schemeName = parser.getAttributeValue(null, "val")
                                val baseColor = resolveSchemeColor(schemeName, themeColors)
                                if (baseColor != null) {
                                    val mods = readColorModifiers(parser)
                                    return applyColorModifiers(
                                        baseColor, mods.lumMod, mods.lumOff,
                                        mods.tint, mods.shade, mods.alpha,
                                    )
                                }
                            }
                        }
                    }
                }
                XmlPullParser.END_TAG -> {
                    when (parser.name) {
                        "bg" -> inBg = false
                        "bgPr" -> inBgPr = false
                        "bgRef" -> inBgRef = false
                    }
                }
            }
        }
        return null
    }

    /**
     * Extract the first color stop from a gradFill element.
     * Parser is at the gradFill START_TAG. Consumes through END_TAG.
     */
    private fun extractFirstGradientStop(
        parser: XmlPullParser,
        themeColors: Map<String, Int>,
    ): Int? {
        val startDepth = parser.depth
        var foundColor: Int? = null

        while (parser.next() != XmlPullParser.END_DOCUMENT) {
            when (parser.eventType) {
                XmlPullParser.START_TAG -> {
                    when (parser.name) {
                        "gs" -> {
                            // Only grab the first gradient stop
                            if (foundColor == null) {
                                foundColor = resolveColorFromElement(parser, themeColors)
                            }
                        }
                    }
                }
                XmlPullParser.END_TAG -> {
                    if (parser.depth <= startDepth) return foundColor
                }
            }
        }
        return foundColor
    }

    // ==================== Relationships ====================

    private data class RelInfo(val target: String, val type: String)

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

    private fun parseRelationshipsWithType(bytes: ByteArray?): Map<String, RelInfo> {
        if (bytes == null) return emptyMap()
        val rels = mutableMapOf<String, RelInfo>()
        val parser = createParser(bytes)

        while (parser.next() != XmlPullParser.END_DOCUMENT) {
            if (parser.eventType == XmlPullParser.START_TAG && parser.name == "Relationship") {
                val id = parser.getAttributeValue(null, "Id") ?: continue
                val target = parser.getAttributeValue(null, "Target") ?: continue
                val type = parser.getAttributeValue(null, "Type") ?: ""
                rels[id] = RelInfo(target, type)
            }
        }
        return rels
    }

    // ==================== Presentation ====================

    private data class PresentationInfo(
        val slideRIds: List<String>,
        val slideWidth: Long = 9144000,
        val slideHeight: Long = 6858000,
    )

    private fun parsePresentation(bytes: ByteArray): PresentationInfo {
        val parser = createParser(bytes)
        val slideRIds = mutableListOf<String>()
        var slideWidth = 9144000L
        var slideHeight = 6858000L

        while (parser.next() != XmlPullParser.END_DOCUMENT) {
            if (parser.eventType == XmlPullParser.START_TAG) {
                when (parser.name) {
                    "sldSz" -> {
                        parser.getAttributeValue(null, "cx")?.toLongOrNull()
                            ?.let { slideWidth = it }
                        parser.getAttributeValue(null, "cy")?.toLongOrNull()
                            ?.let { slideHeight = it }
                    }
                    "sldId" -> {
                        val rId = parser.getAttributeValue(NS_R, "id")
                            ?: parser.getAttributeValue(null, "r:id")
                        if (rId != null) slideRIds.add(rId)
                    }
                }
            }
        }

        return PresentationInfo(slideRIds, slideWidth, slideHeight)
    }

    // ==================== Slide Parsing ====================

    private data class DefaultRunProps(
        val bold: Boolean? = null,
        val italic: Boolean? = null,
        val underline: Boolean? = null,
        val fontSizePt: Float? = null,
        val fontColor: Int? = null,
        val fontName: String? = null,
    )

    private data class GroupTransform(
        val groupX: Long,
        val groupY: Long,
        val groupW: Long,
        val groupH: Long,
        val childOffX: Long,
        val childOffY: Long,
        val childExtW: Long,
        val childExtH: Long,
    ) {
        fun transformX(childX: Long): Long {
            if (childExtW == 0L) return groupX
            return groupX + (childX - childOffX) * groupW / childExtW
        }

        fun transformY(childY: Long): Long {
            if (childExtH == 0L) return groupY
            return groupY + (childY - childOffY) * groupH / childExtH
        }

        fun transformW(childW: Long): Long {
            if (childExtW == 0L) return childW
            return childW * groupW / childExtW
        }

        fun transformH(childH: Long): Long {
            if (childExtH == 0L) return childH
            return childH * groupH / childExtH
        }
    }

    private fun parseSlide(
        slideNumber: Int,
        bytes: ByteArray,
        images: Map<String, ByteArray>,
        themeColors: Map<String, Int>,
        inheritedBgColor: Int?,
        layoutPlaceholders: List<LayoutPlaceholder>,
        backgroundShapes: List<PptxShape>,
        slideWidth: Long,
        slideHeight: Long,
    ): PptxSlide {
        // Extract background color from the slide itself
        val slideBgColor = extractBackgroundColor(bytes, themeColors)
        val backgroundColor = slideBgColor ?: inheritedBgColor

        // Parse slide shapes
        val shapes = mutableListOf<PptxShape>()
        val parser = createParser(bytes)
        parseShapesFromTree(parser, shapes, themeColors, images, emptyList(), layoutPlaceholders, slideWidth, slideHeight)

        return PptxSlide(
            slideNumber = slideNumber,
            shapes = shapes,
            backgroundColor = backgroundColor,
            images = images,
            backgroundShapes = backgroundShapes,
        )
    }

    private fun parseShapesFromTree(
        parser: XmlPullParser,
        shapes: MutableList<PptxShape>,
        themeColors: Map<String, Int>,
        images: Map<String, ByteArray>,
        groupTransforms: List<GroupTransform>,
        layoutPlaceholders: List<LayoutPlaceholder>,
        slideWidth: Long,
        slideHeight: Long,
    ) {
        var inSpTree = false

        while (parser.next() != XmlPullParser.END_DOCUMENT) {
            when (parser.eventType) {
                XmlPullParser.START_TAG -> {
                    when (parser.name) {
                        "spTree" -> inSpTree = true
                        "sp" -> {
                            if (inSpTree || groupTransforms.isNotEmpty()) {
                                parseShape(parser, shapes, themeColors, groupTransforms, layoutPlaceholders, slideWidth, slideHeight)
                            }
                        }
                        "pic" -> {
                            if (inSpTree || groupTransforms.isNotEmpty()) {
                                parsePicture(parser, shapes, images, groupTransforms)
                            }
                        }
                        "cxnSp" -> {
                            if (inSpTree || groupTransforms.isNotEmpty()) {
                                parseConnector(parser, shapes, themeColors, groupTransforms)
                            }
                        }
                        "grpSp" -> {
                            if (inSpTree || groupTransforms.isNotEmpty()) {
                                parseGroupShape(parser, shapes, themeColors, images, groupTransforms, layoutPlaceholders, slideWidth, slideHeight)
                            }
                        }
                    }
                }
                XmlPullParser.END_TAG -> {
                    when (parser.name) {
                        "spTree" -> return
                    }
                }
            }
        }
    }

    private fun parseGroupShape(
        parser: XmlPullParser,
        shapes: MutableList<PptxShape>,
        themeColors: Map<String, Int>,
        images: Map<String, ByteArray>,
        parentTransforms: List<GroupTransform>,
        layoutPlaceholders: List<LayoutPlaceholder>,
        slideWidth: Long,
        slideHeight: Long,
    ) {
        val startDepth = parser.depth
        var grpX = 0L; var grpY = 0L; var grpW = 0L; var grpH = 0L
        var chOffX = 0L; var chOffY = 0L; var chExtW = 0L; var chExtH = 0L
        var inGrpSpPr = false
        var inXfrm = false
        var gotOff = false
        var gotExt = false

        while (parser.next() != XmlPullParser.END_DOCUMENT) {
            when (parser.eventType) {
                XmlPullParser.START_TAG -> {
                    when (parser.name) {
                        "grpSpPr" -> inGrpSpPr = true
                        "xfrm" -> if (inGrpSpPr) inXfrm = true
                        "off" -> if (inXfrm) {
                            if (!gotOff) {
                                grpX = parser.getAttributeValue(null, "x")?.toLongOrNull() ?: 0L
                                grpY = parser.getAttributeValue(null, "y")?.toLongOrNull() ?: 0L
                                gotOff = true
                            } else {
                                chOffX = parser.getAttributeValue(null, "x")?.toLongOrNull() ?: 0L
                                chOffY = parser.getAttributeValue(null, "y")?.toLongOrNull() ?: 0L
                            }
                        }
                        "ext" -> if (inXfrm) {
                            if (!gotExt) {
                                grpW = parser.getAttributeValue(null, "cx")?.toLongOrNull() ?: 0L
                                grpH = parser.getAttributeValue(null, "cy")?.toLongOrNull() ?: 0L
                                gotExt = true
                            } else {
                                chExtW = parser.getAttributeValue(null, "cx")?.toLongOrNull() ?: 0L
                                chExtH = parser.getAttributeValue(null, "cy")?.toLongOrNull() ?: 0L
                            }
                        }
                        "chOff" -> if (inXfrm) {
                            chOffX = parser.getAttributeValue(null, "x")?.toLongOrNull() ?: 0L
                            chOffY = parser.getAttributeValue(null, "y")?.toLongOrNull() ?: 0L
                        }
                        "chExt" -> if (inXfrm) {
                            chExtW = parser.getAttributeValue(null, "cx")?.toLongOrNull() ?: 0L
                            chExtH = parser.getAttributeValue(null, "cy")?.toLongOrNull() ?: 0L
                        }
                        "sp" -> {
                            val transform = GroupTransform(grpX, grpY, grpW, grpH, chOffX, chOffY, chExtW, chExtH)
                            val transforms = parentTransforms + transform
                            parseShape(parser, shapes, themeColors, transforms, layoutPlaceholders, slideWidth, slideHeight)
                        }
                        "pic" -> {
                            val transform = GroupTransform(grpX, grpY, grpW, grpH, chOffX, chOffY, chExtW, chExtH)
                            val transforms = parentTransforms + transform
                            parsePicture(parser, shapes, images, transforms)
                        }
                        "cxnSp" -> {
                            val transform = GroupTransform(grpX, grpY, grpW, grpH, chOffX, chOffY, chExtW, chExtH)
                            val transforms = parentTransforms + transform
                            parseConnector(parser, shapes, themeColors, transforms)
                        }
                        "grpSp" -> {
                            val transform = GroupTransform(grpX, grpY, grpW, grpH, chOffX, chOffY, chExtW, chExtH)
                            val transforms = parentTransforms + transform
                            parseGroupShape(parser, shapes, themeColors, images, transforms, layoutPlaceholders, slideWidth, slideHeight)
                        }
                    }
                }
                XmlPullParser.END_TAG -> {
                    when (parser.name) {
                        "grpSpPr" -> {
                            inGrpSpPr = false
                            inXfrm = false
                        }
                        "xfrm" -> inXfrm = false
                    }
                    if (parser.depth <= startDepth) return
                }
            }
        }
    }

    private fun applyGroupTransforms(
        x: Long, y: Long, w: Long, h: Long,
        transforms: List<GroupTransform>,
    ): LongArray {
        var cx = x; var cy = y; var cw = w; var ch = h
        for (t in transforms) {
            val newX = t.transformX(cx)
            val newY = t.transformY(cy)
            val newW = t.transformW(cw)
            val newH = t.transformH(ch)
            cx = newX; cy = newY; cw = newW; ch = newH
        }
        return longArrayOf(cx, cy, cw, ch)
    }

    // ==================== Individual Shape Parsers ====================

    private fun parseShape(
        parser: XmlPullParser,
        shapes: MutableList<PptxShape>,
        themeColors: Map<String, Int>,
        groupTransforms: List<GroupTransform>,
        layoutPlaceholders: List<LayoutPlaceholder>,
        slideWidth: Long,
        slideHeight: Long,
    ) {
        val startDepth = parser.depth
        var shapeX = 0L; var shapeY = 0L; var shapeW = 0L; var shapeH = 0L
        var shapeRotation = 0f
        var fillColor: Int? = null
        var borderColor: Int? = null
        var borderWidth = 0L
        val paragraphs = mutableListOf<PptxParagraph>()

        var inSpPr = false
        var inTxBody = false
        var inLn = false

        // Placeholder detection
        var phType: String? = null
        var phIdx: Int? = null
        var hasPh = false

        while (parser.next() != XmlPullParser.END_DOCUMENT) {
            when (parser.eventType) {
                XmlPullParser.START_TAG -> {
                    when (parser.name) {
                        "ph" -> {
                            hasPh = true
                            phType = parser.getAttributeValue(null, "type")
                            phIdx = parser.getAttributeValue(null, "idx")?.toIntOrNull()
                        }
                        "spPr" -> inSpPr = true
                        "xfrm" -> {
                            if (inSpPr) {
                                parser.getAttributeValue(null, "rot")?.toLongOrNull()?.let {
                                    shapeRotation = it / 60000f
                                }
                            }
                        }
                        "off" -> {
                            if (inSpPr && !inLn) {
                                shapeX = parser.getAttributeValue(null, "x")?.toLongOrNull() ?: 0L
                                shapeY = parser.getAttributeValue(null, "y")?.toLongOrNull() ?: 0L
                            }
                        }
                        "ext" -> {
                            if (inSpPr && !inLn && !inTxBody) {
                                shapeW = parser.getAttributeValue(null, "cx")?.toLongOrNull() ?: 0L
                                shapeH = parser.getAttributeValue(null, "cy")?.toLongOrNull() ?: 0L
                            }
                        }
                        "solidFill" -> {
                            if (inSpPr && !inLn && !inTxBody) {
                                fillColor = readSolidFillColor(parser, themeColors)
                            } else if (inLn && !inTxBody) {
                                borderColor = readSolidFillColor(parser, themeColors)
                            }
                        }
                        "ln" -> {
                            if (inSpPr && !inTxBody) {
                                inLn = true
                                parser.getAttributeValue(null, "w")?.toLongOrNull()
                                    ?.let { borderWidth = it }
                            }
                        }
                        "txBody" -> {
                            inTxBody = true
                            parseTxBody(parser, paragraphs, themeColors)
                            inTxBody = false
                        }
                    }
                }
                XmlPullParser.END_TAG -> {
                    when (parser.name) {
                        "spPr" -> {
                            inSpPr = false
                            inLn = false
                        }
                        "ln" -> inLn = false
                    }
                    if (parser.depth <= startDepth) {
                        // Resolve placeholder dimensions and bullets from layout
                        val layoutPh = if (hasPh) findLayoutPlaceholder(phType, phIdx, layoutPlaceholders) else null
                        if (layoutPh != null && (shapeW == 0L || shapeH == 0L)) {
                            if (shapeX == 0L && shapeY == 0L) {
                                shapeX = layoutPh.x
                                shapeY = layoutPh.y
                            }
                            if (shapeW == 0L) shapeW = layoutPh.w
                            if (shapeH == 0L) shapeH = layoutPh.h
                        }

                        // Apply inherited bullets from layout placeholder's lstStyle
                        if (layoutPh != null && layoutPh.bulletsByLevel.isNotEmpty() && paragraphs.isNotEmpty()) {
                            for (i in paragraphs.indices) {
                                val p = paragraphs[i]
                                if (p.bulletChar == null && p.runs.isNotEmpty() && p.runs.any { it.text.isNotBlank() }) {
                                    val bullet = layoutPh.bulletsByLevel[p.level] ?: layoutPh.bulletsByLevel[0]
                                    if (bullet != null) {
                                        paragraphs[i] = p.copy(bulletChar = bullet)
                                    }
                                }
                            }
                        }

                        // Safety fallback: if shape has text but still 0 dimensions, use slide defaults
                        if (paragraphs.isNotEmpty() && shapeW == 0L) {
                            shapeW = (slideWidth * 85 / 100)
                            shapeX = (slideWidth - shapeW) / 2
                        }
                        if (paragraphs.isNotEmpty() && shapeH == 0L) {
                            shapeH = (slideHeight * 60 / 100)
                            if (shapeY == 0L) shapeY = slideHeight * 20 / 100
                        }

                        val coords = applyGroupTransforms(shapeX, shapeY, shapeW, shapeH, groupTransforms)
                        val fx = coords[0]; val fy = coords[1]
                        val fw = coords[2]; val fh = coords[3]

                        if (paragraphs.isNotEmpty()) {
                            shapes.add(
                                PptxTextBox(
                                    x = fx, y = fy, width = fw, height = fh,
                                    paragraphs = paragraphs.toList(),
                                    fillColor = fillColor,
                                    borderColor = borderColor,
                                    borderWidthEmu = borderWidth,
                                    rotation = shapeRotation,
                                )
                            )
                        } else if (fillColor != null || borderColor != null) {
                            shapes.add(
                                PptxRectangle(
                                    x = fx, y = fy, width = fw, height = fh,
                                    fillColor = fillColor,
                                    borderColor = borderColor,
                                    borderWidthEmu = borderWidth,
                                    rotation = shapeRotation,
                                )
                            )
                        }
                        return
                    }
                }
            }
        }
    }

    private fun parsePicture(
        parser: XmlPullParser,
        shapes: MutableList<PptxShape>,
        images: Map<String, ByteArray>,
        groupTransforms: List<GroupTransform>,
    ) {
        val startDepth = parser.depth
        var shapeX = 0L; var shapeY = 0L; var shapeW = 0L; var shapeH = 0L
        var shapeRotation = 0f
        var imageRelId: String? = null

        while (parser.next() != XmlPullParser.END_DOCUMENT) {
            when (parser.eventType) {
                XmlPullParser.START_TAG -> {
                    when (parser.name) {
                        "xfrm" -> {
                            parser.getAttributeValue(null, "rot")?.toLongOrNull()?.let {
                                shapeRotation = it / 60000f
                            }
                        }
                        "off" -> {
                            shapeX = parser.getAttributeValue(null, "x")?.toLongOrNull() ?: 0L
                            shapeY = parser.getAttributeValue(null, "y")?.toLongOrNull() ?: 0L
                        }
                        "ext" -> {
                            shapeW = parser.getAttributeValue(null, "cx")?.toLongOrNull() ?: 0L
                            shapeH = parser.getAttributeValue(null, "cy")?.toLongOrNull() ?: 0L
                        }
                        "blip" -> {
                            imageRelId = parser.getAttributeValue(NS_R, "embed")
                                ?: parser.getAttributeValue(null, "r:embed")
                        }
                    }
                }
                XmlPullParser.END_TAG -> {
                    if (parser.depth <= startDepth) {
                        if (imageRelId != null) {
                            val coords = applyGroupTransforms(shapeX, shapeY, shapeW, shapeH, groupTransforms)
                            shapes.add(
                                PptxImageShape(
                                    x = coords[0], y = coords[1],
                                    width = coords[2], height = coords[3],
                                    relationId = imageRelId!!,
                                    rotation = shapeRotation,
                                )
                            )
                        }
                        return
                    }
                }
            }
        }
    }

    private fun parseConnector(
        parser: XmlPullParser,
        shapes: MutableList<PptxShape>,
        themeColors: Map<String, Int>,
        groupTransforms: List<GroupTransform>,
    ) {
        val startDepth = parser.depth
        var shapeX = 0L; var shapeY = 0L; var shapeW = 0L; var shapeH = 0L
        var lineColor: Int? = null
        var lineWidth = 12700L
        var inSpPr = false
        var inLn = false

        while (parser.next() != XmlPullParser.END_DOCUMENT) {
            when (parser.eventType) {
                XmlPullParser.START_TAG -> {
                    when (parser.name) {
                        "spPr" -> inSpPr = true
                        "off" -> if (inSpPr) {
                            shapeX = parser.getAttributeValue(null, "x")?.toLongOrNull() ?: 0L
                            shapeY = parser.getAttributeValue(null, "y")?.toLongOrNull() ?: 0L
                        }
                        "ext" -> if (inSpPr) {
                            shapeW = parser.getAttributeValue(null, "cx")?.toLongOrNull() ?: 0L
                            shapeH = parser.getAttributeValue(null, "cy")?.toLongOrNull() ?: 0L
                        }
                        "ln" -> if (inSpPr) {
                            inLn = true
                            parser.getAttributeValue(null, "w")?.toLongOrNull()
                                ?.let { lineWidth = it }
                        }
                        "solidFill" -> if (inLn) {
                            lineColor = readSolidFillColor(parser, themeColors)
                        }
                    }
                }
                XmlPullParser.END_TAG -> {
                    when (parser.name) {
                        "spPr" -> { inSpPr = false; inLn = false }
                        "ln" -> inLn = false
                    }
                    if (parser.depth <= startDepth) {
                        val coords = applyGroupTransforms(shapeX, shapeY, shapeW, shapeH, groupTransforms)
                        shapes.add(
                            PptxLine(
                                x = coords[0], y = coords[1],
                                width = coords[2], height = coords[3],
                                color = lineColor ?: AndroidColor.BLACK,
                                lineWidthEmu = lineWidth,
                            )
                        )
                        return
                    }
                }
            }
        }
    }

    // ==================== Text Body Parsing ====================

    private fun parseTxBody(
        parser: XmlPullParser,
        paragraphs: MutableList<PptxParagraph>,
        themeColors: Map<String, Int>,
        outBulletsByLevel: MutableMap<Int, String>? = null,
    ) {
        val startDepth = parser.depth

        while (parser.next() != XmlPullParser.END_DOCUMENT) {
            when (parser.eventType) {
                XmlPullParser.START_TAG -> {
                    when (parser.name) {
                        "lstStyle" -> {
                            if (outBulletsByLevel != null) {
                                parseLstStyle(parser, outBulletsByLevel)
                            }
                        }
                        "p" -> {
                            parseParagraph(parser, paragraphs, themeColors)
                        }
                    }
                }
                XmlPullParser.END_TAG -> {
                    if (parser.depth <= startDepth) return
                }
            }
        }
    }

    private fun parseParagraph(
        parser: XmlPullParser,
        paragraphs: MutableList<PptxParagraph>,
        themeColors: Map<String, Int>,
    ) {
        val startDepth = parser.depth
        val runs = mutableListOf<PptxRun>()
        var alignment = PptxAlignment.LEFT
        var bulletChar: String? = null
        var level = 0
        var spacingBefore = 0f
        var spacingAfter = 0f
        var defRunProps: DefaultRunProps? = null

        while (parser.next() != XmlPullParser.END_DOCUMENT) {
            when (parser.eventType) {
                XmlPullParser.START_TAG -> {
                    when (parser.name) {
                        "pPr" -> {
                            parser.getAttributeValue(null, "lvl")?.toIntOrNull()
                                ?.let { level = it }
                            val algn = parser.getAttributeValue(null, "algn")
                            alignment = when (algn) {
                                "ctr" -> PptxAlignment.CENTER
                                "r" -> PptxAlignment.RIGHT
                                "just" -> PptxAlignment.JUSTIFY
                                else -> PptxAlignment.LEFT
                            }
                            val pprResult = parseParagraphProperties(parser, themeColors)
                            defRunProps = pprResult.defRunProps
                            bulletChar = pprResult.bulletChar
                            spacingBefore = pprResult.spacingBefore
                            spacingAfter = pprResult.spacingAfter
                        }
                        "r" -> {
                            parseRun(parser, runs, themeColors, defRunProps)
                        }
                        "br" -> {
                            runs.add(PptxRun(text = "\n"))
                            skipToEndTag(parser)
                        }
                    }
                }
                XmlPullParser.END_TAG -> {
                    if (parser.depth <= startDepth) {
                        paragraphs.add(
                            PptxParagraph(
                                runs = runs.toList(),
                                alignment = alignment,
                                spacingBeforePt = spacingBefore,
                                spacingAfterPt = spacingAfter,
                                bulletChar = bulletChar,
                                level = level,
                            )
                        )
                        return
                    }
                }
            }
        }
    }

    private data class ParagraphPropertiesResult(
        val defRunProps: DefaultRunProps?,
        val bulletChar: String?,
        val spacingBefore: Float,
        val spacingAfter: Float,
    )

    private fun parseParagraphProperties(
        parser: XmlPullParser,
        themeColors: Map<String, Int>,
    ): ParagraphPropertiesResult {
        val startDepth = parser.depth
        var defRunProps: DefaultRunProps? = null
        var bulletChar: String? = null
        var spacingBefore = 0f
        var spacingAfter = 0f
        var inSpcBef = false
        var inSpcAft = false

        while (parser.next() != XmlPullParser.END_DOCUMENT) {
            when (parser.eventType) {
                XmlPullParser.START_TAG -> {
                    when (parser.name) {
                        "buChar" -> {
                            bulletChar = parser.getAttributeValue(null, "char")
                        }
                        "buNone" -> {
                            bulletChar = null
                        }
                        "spcBef" -> inSpcBef = true
                        "spcAft" -> inSpcAft = true
                        "spcPts" -> {
                            val pts = parser.getAttributeValue(null, "val")
                                ?.toFloatOrNull()?.let { it / 100f }
                            if (pts != null) {
                                if (inSpcBef) spacingBefore = pts
                                else if (inSpcAft) spacingAfter = pts
                            }
                        }
                        "defRPr" -> {
                            defRunProps = parseRunProperties(parser, themeColors)
                        }
                    }
                }
                XmlPullParser.END_TAG -> {
                    when (parser.name) {
                        "spcBef" -> inSpcBef = false
                        "spcAft" -> inSpcAft = false
                    }
                    if (parser.depth <= startDepth) {
                        return ParagraphPropertiesResult(defRunProps, bulletChar, spacingBefore, spacingAfter)
                    }
                }
            }
        }

        return ParagraphPropertiesResult(defRunProps, bulletChar, spacingBefore, spacingAfter)
    }

    private fun parseRun(
        parser: XmlPullParser,
        runs: MutableList<PptxRun>,
        themeColors: Map<String, Int>,
        defRunProps: DefaultRunProps?,
    ) {
        val startDepth = parser.depth
        val textBuilder = StringBuilder()
        var runProps: DefaultRunProps? = null

        while (parser.next() != XmlPullParser.END_DOCUMENT) {
            when (parser.eventType) {
                XmlPullParser.START_TAG -> {
                    when (parser.name) {
                        "rPr" -> {
                            runProps = parseRunProperties(parser, themeColors)
                        }
                    }
                }
                XmlPullParser.TEXT -> {
                    textBuilder.append(parser.text)
                }
                XmlPullParser.END_TAG -> {
                    if (parser.depth <= startDepth) {
                        if (textBuilder.isNotEmpty()) {
                            val bold = runProps?.bold ?: defRunProps?.bold ?: false
                            val italic = runProps?.italic ?: defRunProps?.italic ?: false
                            val underline = runProps?.underline ?: defRunProps?.underline ?: false
                            val fontSize = runProps?.fontSizePt ?: defRunProps?.fontSizePt
                            val fontColor = runProps?.fontColor ?: defRunProps?.fontColor
                            val fontName = runProps?.fontName ?: defRunProps?.fontName

                            runs.add(
                                PptxRun(
                                    text = textBuilder.toString(),
                                    bold = bold,
                                    italic = italic,
                                    underline = underline,
                                    fontSizePt = fontSize,
                                    fontColor = fontColor,
                                    fontName = fontName,
                                )
                            )
                        }
                        return
                    }
                }
            }
        }
    }

    private fun parseRunProperties(
        parser: XmlPullParser,
        themeColors: Map<String, Int>,
    ): DefaultRunProps {
        val startDepth = parser.depth

        val bold = parser.getAttributeValue(null, "b")?.let { it == "1" }
        val italic = parser.getAttributeValue(null, "i")?.let { it == "1" }
        val underline = parser.getAttributeValue(null, "u")?.let { it != "none" }
        val fontSize = parser.getAttributeValue(null, "sz")?.toFloatOrNull()?.let { it / 100f }

        var fontColor: Int? = null
        var fontName: String? = null

        while (parser.next() != XmlPullParser.END_DOCUMENT) {
            when (parser.eventType) {
                XmlPullParser.START_TAG -> {
                    when (parser.name) {
                        "solidFill" -> {
                            fontColor = readRunColor(parser, themeColors)
                        }
                        "latin", "ea", "cs" -> {
                            if (fontName == null) {
                                parser.getAttributeValue(null, "typeface")?.let {
                                    if (!it.startsWith("+")) fontName = it
                                }
                            }
                        }
                    }
                }
                XmlPullParser.END_TAG -> {
                    if (parser.depth <= startDepth) {
                        return DefaultRunProps(
                            bold = bold,
                            italic = italic,
                            underline = underline,
                            fontSizePt = fontSize,
                            fontColor = fontColor,
                            fontName = fontName,
                        )
                    }
                }
            }
        }

        return DefaultRunProps(bold, italic, underline, fontSize, fontColor, fontName)
    }

    // ==================== Utilities ====================

    private fun skipToEndTag(parser: XmlPullParser) {
        val depth = parser.depth
        while (parser.next() != XmlPullParser.END_DOCUMENT) {
            if (parser.eventType == XmlPullParser.END_TAG && parser.depth <= depth) return
        }
    }

    private fun normalizePath(path: String): String {
        val parts = path.split("/").toMutableList()
        val result = mutableListOf<String>()
        for (part in parts) {
            when (part) {
                ".." -> if (result.isNotEmpty()) result.removeAt(result.lastIndex)
                "." -> { /* skip */ }
                else -> result.add(part)
            }
        }
        return result.joinToString("/")
    }

    private fun createParser(bytes: ByteArray): XmlPullParser {
        val factory = XmlPullParserFactory.newInstance()
        factory.isNamespaceAware = true
        val parser = factory.newPullParser()
        parser.setInput(ByteArrayInputStream(bytes), "UTF-8")
        return parser
    }
}
