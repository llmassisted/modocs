package com.modocs.feature.docx

import android.graphics.Color as AndroidColor

/**
 * In-memory representation of a parsed DOCX file.
 * Contains the document body elements plus resolved styles and images.
 */
/**
 * Page size and margin settings parsed from w:sectPr.
 * Values in twips (1 twip = 1/1440 inch = 1/20 point).
 */
data class PageSetup(
    val pageWidthTwips: Int = 12240,     // US Letter width 8.5" (default)
    val pageHeightTwips: Int = 15840,    // US Letter height 11" (default)
    val marginTopTwips: Int = 1440,      // 1 inch
    val marginBottomTwips: Int = 1440,
    val marginLeftTwips: Int = 1440,
    val marginRightTwips: Int = 1440,
) {
    /** Page width in points. */
    val pageWidthPt: Float get() = pageWidthTwips / 20f
    /** Page height in points. */
    val pageHeightPt: Float get() = pageHeightTwips / 20f
    /** Left margin in points. */
    val marginLeftPt: Float get() = marginLeftTwips / 20f
    /** Right margin in points. */
    val marginRightPt: Float get() = marginRightTwips / 20f
    /** Top margin in points. */
    val marginTopPt: Float get() = marginTopTwips / 20f
    /** Bottom margin in points. */
    val marginBottomPt: Float get() = marginBottomTwips / 20f
    /** Printable content width in points. */
    val contentWidthPt: Float get() = pageWidthPt - marginLeftPt - marginRightPt
}

data class DocxDocument(
    val body: MutableList<DocxElement>,
    val styles: Map<String, DocxStyle> = emptyMap(),
    val numbering: Map<String, NumberingDefinition> = emptyMap(),
    val images: Map<String, ByteArray> = emptyMap(),
    /** Raw ZIP entries for round-trip save (preserves unmodified content). */
    val rawEntries: Map<String, ByteArray> = emptyMap(),
    /** Page size and margins from section properties. */
    val pageSetup: PageSetup = PageSetup(),
)

// --- Elements ---

sealed interface DocxElement

data class DocxParagraph(
    val runs: List<DocxRun>,
    val properties: ParagraphProperties = ParagraphProperties(),
    val listInfo: ListInfo? = null,
) : DocxElement {
    /** Plain text of all runs concatenated. */
    val text: String get() = runs.joinToString("") { it.text }
}

data class DocxTable(
    val rows: List<DocxTableRow>,
    val properties: TableProperties = TableProperties(),
    /** Column widths from tblGrid in twips — authoritative source for consistent column layout. */
    val gridColWidths: List<Int> = emptyList(),
) : DocxElement

data class DocxTableRow(
    val cells: List<DocxTableCell>,
)

data class DocxTableCell(
    val paragraphs: List<DocxParagraph>,
    val widthTwips: Int? = null,
    val gridSpan: Int = 1,
    val vMerge: VMergeType = VMergeType.NONE,
    val shading: Int? = null,
    val borders: CellBorders? = null,
)

/** Per-side border definition for a table cell. null = inherit from table default. */
data class CellBorders(
    val top: CellBorder? = null,
    val bottom: CellBorder? = null,
    val left: CellBorder? = null,
    val right: CellBorder? = null,
)

data class CellBorder(
    val style: BorderStyle = BorderStyle.SINGLE,
    val widthEighthPt: Int = 4, // border width in 1/8 pt (OOXML sz attribute)
    val color: Int? = null, // null = auto/black
)

enum class VMergeType { NONE, RESTART, CONTINUE }

data class DocxImage(
    val relationId: String,
    val widthEmu: Long = 0,
    val heightEmu: Long = 0,
    val altText: String = "",
) : DocxElement

// --- Run (inline text with formatting) ---

data class DocxRun(
    val text: String,
    val properties: RunProperties = RunProperties(),
) {
    /** Whether this run represents a line break. */
    val isBreak: Boolean get() = text == "\n"
    /** Whether this run represents a page break. */
    val isPageBreak: Boolean get() = text == "\u000C"
}

// --- Properties ---

data class RunProperties(
    val bold: Boolean = false,
    val italic: Boolean = false,
    val underline: Boolean = false,
    val strikethrough: Boolean = false,
    val fontName: String? = null,
    val fontSizeHalfPt: Int? = null,
    val color: Int? = null,
    val highlight: Int? = null,
    val superscript: Boolean = false,
    val subscript: Boolean = false,
) {
    /** Font size in sp. Half-points / 2 = points ≈ sp. */
    val fontSizeSp: Float? get() = fontSizeHalfPt?.let { it / 2f }
}

data class ParagraphProperties(
    val alignment: ParagraphAlignment = ParagraphAlignment.LEFT,
    val spacingBeforePt: Float = 0f,
    val spacingAfterPt: Float = 8f,
    val lineSpacingMultiplier: Float = 1.15f,
    /** Exact line height in points when [lineSpacingExact] is true. */
    val lineSpacingPt: Float = 0f,
    /** True when line spacing is exact/atLeast (use [lineSpacingPt]), false for auto (use [lineSpacingMultiplier]). */
    val lineSpacingExact: Boolean = false,
    val indentLeftTwips: Int = 0,
    val indentRightTwips: Int = 0,
    val indentFirstLineTwips: Int = 0,
    val headingLevel: Int = 0,
    val styleId: String? = null,
    val keepNext: Boolean = false,
    val keepLines: Boolean = false,
    val pageBreakBefore: Boolean = false,
    /** When true, suppress spacingBefore/After between adjacent paragraphs with the same style. */
    val contextualSpacing: Boolean = false,
    /** True when spacingAfterPt was explicitly set in XML (not just default). */
    val spacingAfterExplicit: Boolean = false,
    /** True when spacingBeforePt was explicitly set in XML. */
    val spacingBeforeExplicit: Boolean = false,
    /** True when line spacing was explicitly set in XML. */
    val lineSpacingExplicitlySet: Boolean = false,
) {
    /** Effective line height multiplier — accounts for exact mode. */
    fun effectiveLineHeight(fontSize: Float): Float {
        return if (lineSpacingExact && lineSpacingPt > 0f) {
            lineSpacingPt
        } else {
            fontSize * lineSpacingMultiplier
        }
    }
}

enum class ParagraphAlignment {
    LEFT, CENTER, RIGHT, JUSTIFY
}

data class TableProperties(
    val widthTwips: Int? = null,
    val alignment: ParagraphAlignment = ParagraphAlignment.LEFT,
    val borders: TableBorders = TableBorders(),
)

/** Table-level default borders (from tblBorders). Applied to all cells unless overridden by tcBorders. */
data class TableBorders(
    val top: CellBorder = CellBorder(),
    val bottom: CellBorder = CellBorder(),
    val left: CellBorder = CellBorder(),
    val right: CellBorder = CellBorder(),
    val insideH: CellBorder = CellBorder(),
    val insideV: CellBorder = CellBorder(),
)

enum class BorderStyle { NONE, SINGLE, DOUBLE, DASHED, DOTTED, DASH_SMALL_GAP }

// --- Styles ---

data class DocxStyle(
    val styleId: String,
    val name: String = "",
    val basedOn: String? = null,
    val runProperties: RunProperties = RunProperties(),
    val paragraphProperties: ParagraphProperties = ParagraphProperties(),
    val isDefault: Boolean = false,
)

// --- Numbering (lists) ---

data class ListInfo(
    val numId: String,
    val level: Int = 0,
)

data class NumberingDefinition(
    val abstractNumId: String,
    val levels: Map<Int, NumberingLevel> = emptyMap(),
)

data class NumberingLevel(
    val level: Int,
    val format: NumberFormat = NumberFormat.BULLET,
    val text: String = "•",
    val indentTwips: Int = 720,
)

enum class NumberFormat {
    BULLET, DECIMAL, LOWER_ALPHA, UPPER_ALPHA, LOWER_ROMAN, UPPER_ROMAN;

    fun formatNumber(number: Int): String = when (this) {
        BULLET -> "•"
        DECIMAL -> "$number."
        LOWER_ALPHA -> "${('a' + (number - 1) % 26)}."
        UPPER_ALPHA -> "${('A' + (number - 1) % 26)}."
        LOWER_ROMAN -> toRoman(number).lowercase() + "."
        UPPER_ROMAN -> toRoman(number) + "."
    }

    companion object {
        private fun toRoman(num: Int): String {
            val values = intArrayOf(1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1)
            val symbols = arrayOf("M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I")
            val sb = StringBuilder()
            var n = num
            for (i in values.indices) {
                while (n >= values[i]) {
                    sb.append(symbols[i])
                    n -= values[i]
                }
            }
            return sb.toString()
        }
    }
}

// --- Color helpers ---

object DocxColors {
    /** Parse OOXML color string (e.g., "FF0000", "auto") to Android color int. */
    fun parseColor(colorStr: String?): Int? {
        if (colorStr.isNullOrBlank() || colorStr == "auto") return null
        return try {
            val hex = if (colorStr.length == 6) "FF$colorStr" else colorStr
            AndroidColor.parseColor("#$hex")
        } catch (_: Exception) {
            null
        }
    }

    /** Parse a highlight color name to Android color int. */
    fun parseHighlight(name: String?): Int? = when (name?.lowercase()) {
        "yellow" -> AndroidColor.parseColor("#FFFF00")
        "green" -> AndroidColor.parseColor("#00FF00")
        "cyan" -> AndroidColor.parseColor("#00FFFF")
        "magenta" -> AndroidColor.parseColor("#FF00FF")
        "blue" -> AndroidColor.parseColor("#0000FF")
        "red" -> AndroidColor.parseColor("#FF0000")
        "darkblue", "darkBlue" -> AndroidColor.parseColor("#000080")
        "darkred", "darkRed" -> AndroidColor.parseColor("#800000")
        "darkgreen", "darkGreen" -> AndroidColor.parseColor("#008000")
        "darkyellow", "darkYellow" -> AndroidColor.parseColor("#808000")
        "darkmagenta", "darkMagenta" -> AndroidColor.parseColor("#800080")
        "darkcyan", "darkCyan" -> AndroidColor.parseColor("#008080")
        "lightgray", "lightGray" -> AndroidColor.parseColor("#C0C0C0")
        "darkgray", "darkGray" -> AndroidColor.parseColor("#808080")
        "black" -> AndroidColor.parseColor("#000000")
        "white" -> AndroidColor.parseColor("#FFFFFF")
        else -> null
    }
}

// --- Unit conversion helpers ---

object DocxUnits {
    /** Twips to dp (1 twip = 1/1440 inch, 1dp ≈ 1/160 inch). */
    fun twipsToDp(twips: Int): Float = twips / 1440f * 160f

    /** EMU to dp (1 EMU = 1/914400 inch). */
    fun emuToDp(emu: Long): Float = emu / 914400f * 160f

    /** Half-points to sp (1 half-point = 0.5 pt ≈ 0.5 sp). */
    fun halfPtToSp(halfPt: Int): Float = halfPt / 2f
}
