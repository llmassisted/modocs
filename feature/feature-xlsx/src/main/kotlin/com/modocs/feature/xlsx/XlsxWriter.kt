package com.modocs.feature.xlsx

import android.content.Context
import android.net.Uri
import com.modocs.core.common.*
import java.io.OutputStream
import java.util.zip.ZipEntry
import java.util.zip.ZipOutputStream

object XlsxWriter {
    private const val NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

    suspend fun save(context: Context, document: XlsxDocument, outputUri: Uri) =
        saveDocumentCopy(context, outputUri) { write(document, it) }

    fun write(document: XlsxDocument, outputStream: OutputStream) {
        val replacements = mutableMapOf<String, ByteArray>()
        for (sheetIndex in document.modifiedSheets) {
            val path = requireNotNull(document.sheetPaths[sheetIndex]) { "Missing original worksheet" }
            val xml = parseXmlDocument(requireNotNull(document.rawEntries[path]))
            val root = xml.documentElement
            val data = root.childrenNamed(NS, "sheetData").single()
            val sheet = document.sheets[sheetIndex]
            val formulas = data.getElementsByTagNameNS(NS, "f")
            val protectedRanges = (0 until formulas.length).mapNotNull { index ->
                val formula = formulas.item(index) as org.w3c.dom.Element
                if (formula.getAttribute("t") !in listOf("array", "shared", "dataTable")) return@mapNotNull null
                val bounds = formula.getAttribute("ref").split(":")
                val first = parseCellReference(bounds.first()) ?: return@mapNotNull null
                val last = parseCellReference(bounds.last()) ?: first
                first to last
            }
            require(document.editedCells[sheetIndex].orEmpty().none { (r, c) ->
                protectedRanges.any { (first, last) -> r in first.first..last.first && c in first.second..last.second }
            }) { "An edited cell belongs to a shared or array formula. Undo that edit before saving." }
            for ((rowIndex, colIndex) in document.editedCells[sheetIndex].orEmpty()) {
                val cell = sheet.cellAt(rowIndex, colIndex) ?: continue
                val ref = cellReference(rowIndex, colIndex)
                val row = data.childrenNamed(NS, "row").find { it.getAttribute("r") == "${rowIndex + 1}" }
                    ?: xml.createElementNS(NS, "row").also { newRow ->
                        newRow.setAttribute("r", "${rowIndex + 1}")
                        data.insertBefore(newRow, data.childrenNamed(NS, "row").find {
                            (it.getAttribute("r").toIntOrNull() ?: 0) > rowIndex + 1
                        })
                    }
                val node = row.childrenNamed(NS, "c").find { it.getAttribute("r") == ref }
                    ?: xml.createElementNS(NS, "c").also { newCell ->
                        newCell.setAttribute("r", ref)
                        row.insertBefore(newCell, row.childrenNamed(NS, "c").find {
                            (parseCellReference(it.getAttribute("r"))?.second ?: -1) > colIndex
                        })
                    }
                require(node.childrenNamed(NS, "f").none { it.getAttribute("t") in listOf("array", "shared", "dataTable") }) {
                    "This cell belongs to a shared or array formula; edit it in a spreadsheet app"
                }
                node.elements().filter { it.namespaceURI == NS && it.localName in listOf("f", "v", "is") }
                    .forEach { node.removeChild(it) }
                if (cell.type == CellType.NUMBER) {
                    node.removeAttribute("t")
                    node.appendChild(xml.createElementNS(NS, "v").also { it.textContent = cell.value })
                } else {
                    node.setAttribute("t", "inlineStr")
                    node.appendChild(xml.createElementNS(NS, "is").also { inline ->
                        inline.appendChild(xml.createElementNS(NS, "t").also {
                            it.setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:space", "preserve")
                            it.textContent = cell.value
                        })
                    })
                }
            }
            // A stale dimension must not conceal newly edited cells from other readers.
            root.childrenNamed(NS, "dimension").forEach { root.removeChild(it) }
            replacements[path] = xml.xmlBytes()
        }
        if (document.modifiedSheets.isNotEmpty()) {
            document.rawEntries["xl/workbook.xml"]?.let { bytes ->
                val xml = parseXmlDocument(bytes)
                val root = xml.documentElement
                val calc = root.childrenNamed(NS, "calcPr").firstOrNull()
                    ?: xml.createElementNS(NS, "calcPr").also { node ->
                        val afterCalc = setOf("oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr", "smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst")
                        root.insertBefore(node, root.elements().firstOrNull { it.localName in afterCalc })
                    }
                calc.setAttribute("fullCalcOnLoad", "1")
                calc.setAttribute("forceFullCalc", "1")
                calc.setAttribute("calcMode", "auto")
                replacements["xl/workbook.xml"] = xml.xmlBytes()
            }
        }
        ZipOutputStream(outputStream).use { zip ->
            for ((name, bytes) in document.rawEntries) {
                zip.putNextEntry(ZipEntry(name))
                zip.write(replacements[name] ?: bytes)
                zip.closeEntry()
            }
        }
    }
}
