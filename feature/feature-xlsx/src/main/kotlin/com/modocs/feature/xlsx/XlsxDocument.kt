package com.modocs.feature.xlsx

import android.graphics.Color as AndroidColor

/**
 * In-memory representation of a parsed XLSX file.
 */
data class XlsxDocument(
    val sheets: List<XlsxSheet>,
    val styles: List<XlsxCellStyle> = emptyList(),
    val rawEntries: Map<String, ByteArray> = emptyMap(),
    /** Maps sheet index to its ZIP entry path (e.g., 0 -> "xl/worksheets/sheet1.xml"). */
    val sheetPaths: Map<Int, String> = emptyMap(),
    /** Tracks which sheet indices have been modified by the editor. */
    val modifiedSheets: MutableSet<Int> = mutableSetOf(),
)

data class XlsxSheet(
    val name: String,
    val rows: MutableList<XlsxRow>,
    /** Column widths in character units (default ~8.43). */
    val columnWidths: Map<Int, Float> = emptyMap(),
    /** Default row height in points. */
    val defaultRowHeight: Float = 15f,
    /** Merged cell ranges. */
    val mergedCells: List<CellRange> = emptyList(),
    /** Freeze pane: number of frozen columns. */
    val frozenColumns: Int = 0,
    /** Freeze pane: number of frozen rows. */
    val frozenRows: Int = 0,
) {
    /** Maximum column index across all rows. */
    val columnCount: Int
        get() = rows.maxOfOrNull { row -> row.cells.maxOfOrNull { it.columnIndex + 1 } ?: 0 } ?: 0

    /** Get a cell at the given row/column, or null. */
    fun cellAt(rowIndex: Int, colIndex: Int): XlsxCell? {
        val row = rows.find { it.rowIndex == rowIndex } ?: return null
        return row.cells.find { it.columnIndex == colIndex }
    }
}

data class XlsxRow(
    val rowIndex: Int,
    val cells: MutableList<XlsxCell>,
    /** Custom row height in points, or null for default. */
    val heightPt: Float? = null,
)

data class XlsxCell(
    val columnIndex: Int,
    val value: String,
    val type: CellType = CellType.STRING,
    val styleIndex: Int = 0,
    /** Raw formula text, if any. */
    val formula: String? = null,
    /** Raw OOXML cached value. For numbers/dates this differs from the formatted display value. */
    val rawValue: String? = null,
)

enum class CellType {
    STRING, NUMBER, BOOLEAN, DATE, ERROR, INLINE_STRING
}

data class XlsxCellStyle(
    val fontBold: Boolean = false,
    val fontItalic: Boolean = false,
    val fontUnderline: Boolean = false,
    val fontSize: Float = 11f,
    val fontColor: Int? = null,
    val fillColor: Int? = null,
    val horizontalAlignment: CellAlignment = CellAlignment.GENERAL,
    val verticalAlignment: CellVerticalAlignment = CellVerticalAlignment.BOTTOM,
    val numberFormatId: Int = 0,
    val numberFormatCode: String? = null,
    val wrapText: Boolean = false,
)

enum class CellAlignment {
    GENERAL, LEFT, CENTER, RIGHT, FILL, JUSTIFY
}

enum class CellVerticalAlignment {
    TOP, CENTER, BOTTOM
}

data class CellRange(
    val startRow: Int,
    val startCol: Int,
    val endRow: Int,
    val endCol: Int,
)

object XlsxColors {
    fun parseColor(colorStr: String?): Int? {
        if (colorStr.isNullOrBlank()) return null
        return try {
            // OOXML colors are ARGB hex (e.g., "FF000000")
            val hex = when (colorStr.length) {
                6 -> "FF$colorStr"
                8 -> colorStr
                else -> return null
            }
            AndroidColor.parseColor("#$hex")
        } catch (_: Exception) {
            null
        }
    }

    /** Map OOXML indexed colors to ARGB ints. */
    fun indexedColor(index: Int): Int? = when (index) {
        0 -> AndroidColor.parseColor("#FF000000")  // Black
        1 -> AndroidColor.parseColor("#FFFFFFFF")  // White
        2 -> AndroidColor.parseColor("#FFFF0000")  // Red
        3 -> AndroidColor.parseColor("#FF00FF00")  // Green
        4 -> AndroidColor.parseColor("#FF0000FF")  // Blue
        5 -> AndroidColor.parseColor("#FFFFFF00")  // Yellow
        6 -> AndroidColor.parseColor("#FFFF00FF")  // Magenta
        7 -> AndroidColor.parseColor("#FF00FFFF")  // Cyan
        8 -> AndroidColor.parseColor("#FF000000")  // Black
        9 -> AndroidColor.parseColor("#FFFFFFFF")  // White
        else -> null
    }

    /** Map OOXML theme colors (basic set). */
    fun themeColor(index: Int): Int? = when (index) {
        0 -> AndroidColor.parseColor("#FFFFFFFF")  // lt1 (usually white)
        1 -> AndroidColor.parseColor("#FF000000")  // dk1 (usually black)
        2 -> AndroidColor.parseColor("#FFE7E6E6")  // lt2
        3 -> AndroidColor.parseColor("#FF44546A")  // dk2
        4 -> AndroidColor.parseColor("#FF4472C4")  // accent1
        5 -> AndroidColor.parseColor("#FFED7D31")  // accent2
        6 -> AndroidColor.parseColor("#FFA5A5A5")  // accent3
        7 -> AndroidColor.parseColor("#FFFFC000")  // accent4
        8 -> AndroidColor.parseColor("#FF5B9BD5")  // accent5
        9 -> AndroidColor.parseColor("#FF70AD47")  // accent6
        else -> null
    }
}
