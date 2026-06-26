package com.modocs.feature.xlsx

import android.content.Context
import android.net.Uri
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.withContext
import java.io.OutputStream
import java.util.zip.ZipEntry
import java.util.zip.ZipOutputStream

/**
 * Writes a [XlsxDocument] back to an XLSX ZIP file.
 * Uses round-trip approach: copies all original ZIP entries unchanged,
 * except modified sheet XMLs which are regenerated.
 */
object XlsxWriter {

    suspend fun save(context: Context, document: XlsxDocument, outputUri: Uri) {
        withContext(Dispatchers.IO) {
            context.contentResolver.openOutputStream(outputUri)?.use { stream ->
                write(document, stream)
            } ?: throw IllegalStateException("Cannot open output URI for writing")
        }
    }

    fun write(document: XlsxDocument, outputStream: OutputStream) {
        // Determine which ZIP entries need to be replaced
        val modifiedPaths = mutableSetOf<String>()
        for (sheetIdx in document.modifiedSheets) {
            val path = document.sheetPaths[sheetIdx]
            if (path != null) modifiedPaths.add(path)
        }

        ZipOutputStream(outputStream).use { zip ->
            // Write all original entries except those we're replacing
            for ((name, bytes) in document.rawEntries) {
                if (name in modifiedPaths) continue
                zip.putNextEntry(ZipEntry(name))
                zip.write(bytes)
                zip.closeEntry()
            }

            // Write modified sheet XMLs
            for (sheetIdx in document.modifiedSheets) {
                val path = document.sheetPaths[sheetIdx] ?: continue
                val sheet = document.sheets.getOrNull(sheetIdx) ?: continue
                val sheetXml = generateSheetXml(sheet)
                zip.putNextEntry(ZipEntry(path))
                zip.write(sheetXml.toByteArray(Charsets.UTF_8))
                zip.closeEntry()
            }
        }
    }

    /**
     * Generate sheet XML for a modified sheet.
     * Preserves structure from the original XML where possible by reading the
     * original bytes and replacing just the sheetData section.
     */
    private fun generateSheetXml(
        sheet: XlsxSheet,
    ): String {
        // For simplicity and reliability, we rebuild the essential sheet XML.
        // We preserve column widths, merged cells, and freeze pane settings.
        val sb = StringBuilder()
        sb.append("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>""")
        sb.append("\n")
        sb.append("""<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"""")
        sb.append(""" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">""")
        sb.append("\n")

        // Sheet format properties
        sb.append("""<sheetFormatPr defaultRowHeight="${sheet.defaultRowHeight}"/>""")
        sb.append("\n")

        // Column widths
        if (sheet.columnWidths.isNotEmpty()) {
            sb.append("<cols>\n")
            for ((col, width) in sheet.columnWidths.toSortedMap()) {
                val colNum = col + 1 // 1-indexed
                sb.append("""<col min="$colNum" max="$colNum" width="$width" customWidth="1"/>""")
                sb.append("\n")
            }
            sb.append("</cols>\n")
        }

        // Sheet data
        sb.append("<sheetData>\n")
        for (row in sheet.rows) {
            sb.append("""<row r="${row.rowIndex + 1}"""")
            if (row.heightPt != null) {
                sb.append(""" ht="${row.heightPt}" customHeight="1"""")
            }
            sb.append(">\n")

            for (cell in row.cells) {
                val cellRef = columnLetter(cell.columnIndex) + (row.rowIndex + 1)
                sb.append("""<c r="$cellRef"""")

                when (cell.type) {
                    CellType.STRING, CellType.INLINE_STRING -> {
                        if (cell.styleIndex != 0) {
                            sb.append(""" s="${cell.styleIndex}"""")
                        }
                        if (cell.formula != null) {
                            sb.append(""" t="str">""")
                            sb.append("<f>")
                            sb.append(escXml(cell.formula))
                            sb.append("</f>")
                            sb.append("<v>")
                            sb.append(escXml(cell.rawValue ?: cell.value))
                            sb.append("</v>")
                        } else {
                            sb.append(""" t="inlineStr">""")
                            sb.append("<is><t")
                            appendSpacePreserveIfNeeded(sb, cell.value)
                            sb.append(">")
                            sb.append(escXml(cell.value))
                            sb.append("</t></is>")
                        }
                    }
                    CellType.NUMBER, CellType.DATE -> {
                        if (cell.styleIndex != 0) {
                            sb.append(""" s="${cell.styleIndex}"""")
                        }
                        sb.append(">")
                        if (cell.formula != null) {
                            sb.append("<f>")
                            sb.append(escXml(cell.formula))
                            sb.append("</f>")
                        }
                        sb.append("<v>")
                        sb.append(escXml(cell.rawValue ?: cell.value))
                        sb.append("</v>")
                    }
                    CellType.BOOLEAN -> {
                        sb.append(""" t="b"""")
                        if (cell.styleIndex != 0) {
                            sb.append(""" s="${cell.styleIndex}"""")
                        }
                        sb.append(">")
                        sb.append("<v>")
                        sb.append(if (cell.value.equals("TRUE", ignoreCase = true)) "1" else "0")
                        sb.append("</v>")
                    }
                    CellType.ERROR -> {
                        sb.append(""" t="e"""")
                        if (cell.styleIndex != 0) {
                            sb.append(""" s="${cell.styleIndex}"""")
                        }
                        sb.append(">")
                        sb.append("<v>")
                        sb.append(escXml(cell.value))
                        sb.append("</v>")
                    }
                }

                sb.append("</c>\n")
            }

            sb.append("</row>\n")
        }
        sb.append("</sheetData>\n")

        // Merged cells
        if (sheet.mergedCells.isNotEmpty()) {
            sb.append("""<mergeCells count="${sheet.mergedCells.size}">""")
            sb.append("\n")
            for (range in sheet.mergedCells) {
                val startRef = columnLetter(range.startCol) + (range.startRow + 1)
                val endRef = columnLetter(range.endCol) + (range.endRow + 1)
                sb.append("""<mergeCell ref="$startRef:$endRef"/>""")
                sb.append("\n")
            }
            sb.append("</mergeCells>\n")
        }

        // Freeze panes
        if (sheet.frozenRows > 0 || sheet.frozenColumns > 0) {
            sb.append("<sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\">")
            val topLeftCell = columnLetter(sheet.frozenColumns) + (sheet.frozenRows + 1)
            sb.append("""<pane""")
            if (sheet.frozenColumns > 0) sb.append(""" xSplit="${sheet.frozenColumns}"""")
            if (sheet.frozenRows > 0) sb.append(""" ySplit="${sheet.frozenRows}"""")
            sb.append(""" topLeftCell="$topLeftCell" state="frozen"/>""")
            sb.append("</sheetView></sheetViews>\n")
        }

        sb.append("</worksheet>")
        return sb.toString()
    }

    private fun appendSpacePreserveIfNeeded(sb: StringBuilder, text: String) {
        if (text.isNotEmpty() && (text.first() == ' ' || text.last() == ' ')) {
            sb.append(""" xml:space="preserve"""")
        }
    }

    /** Convert 0-based column index to Excel-style letter (A, B, ..., Z, AA, AB, ...). */
    private fun columnLetter(index: Int): String {
        val sb = StringBuilder()
        var n = index
        do {
            sb.insert(0, ('A' + n % 26))
            n = n / 26 - 1
        } while (n >= 0)
        return sb.toString()
    }

    private fun escXml(text: String): String {
        return text.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace("\"", "&quot;")
            .replace("'", "&apos;")
    }
}
