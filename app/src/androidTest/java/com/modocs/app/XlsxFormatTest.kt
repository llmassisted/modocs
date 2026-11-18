package com.modocs.app

import android.net.Uri
import androidx.test.platform.app.InstrumentationRegistry
import kotlinx.coroutines.flow.first
import kotlinx.coroutines.runBlocking
import kotlinx.coroutines.withTimeout
import java.io.File
import androidx.test.ext.junit.runners.AndroidJUnit4
import com.modocs.feature.xlsx.*
import org.junit.After
import org.junit.Assert.*
import org.junit.Before
import org.junit.Test
import org.junit.runner.RunWith
import java.io.ByteArrayOutputStream
import java.util.Locale
import java.util.zip.ZipEntry
import java.util.zip.ZipOutputStream

/** Cells show their number format, while edits still start from the stored number. */
@RunWith(AndroidJUnit4::class)
class XlsxFormatTest {
    private val original = Locale.getDefault()
    @Before fun pinLocale() = Locale.setDefault(Locale.US)
    @After fun restoreLocale() = Locale.setDefault(original)

    private fun xlsx(date1904: Boolean): ByteArray = ByteArrayOutputStream().also { out ->
        ZipOutputStream(out).use { zip ->
            fun put(name: String, content: String) { zip.putNextEntry(ZipEntry(name)); zip.write(content.toByteArray()); zip.closeEntry() }
            val ns = """xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""""
            put("xl/workbook.xml", """<workbook $ns><workbookPr date1904="${if (date1904) 1 else 0}"/><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>""")
            put("xl/_rels/workbook.xml.rels", """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="t" Target="worksheets/sheet1.xml"/></Relationships>""")
            put("xl/styles.xml", """<styleSheet $ns><numFmts count="1"><numFmt numFmtId="164" formatCode="&quot;$&quot;#,##0.00"/></numFmts>
                <fonts><font/></fonts><fills><fill/></fills><cellXfs><xf/><xf numFmtId="9"/><xf numFmtId="164"/><xf numFmtId="14"/><xf numFmtId="18"/></cellXfs></styleSheet>""")
            put("xl/worksheets/sheet1.xml", """<worksheet $ns><sheetData><row r="1">
                <c r="A1" s="1"><v>0.25</v></c><c r="B1" s="2"><v>1234.5</v></c><c r="C1" s="3"><v>45231</v></c><c r="D1" s="4"><v>0.5</v></c><c r="E1"><v>12.5</v></c>
                </row></sheetData></worksheet>""")
        }
    }.toByteArray()

    @Test fun cellsShowTheirNumberFormat() {
        val sheet = XlsxParser.parse(xlsx(date1904 = false).inputStream()).sheets.single()
        assertEquals(listOf("25%", "$1,234.50", "11/1/2023", "12:00 PM", "12.5"), (0..4).map { sheet.cellAt(0, it)?.value })
        assertEquals(listOf(CellType.NUMBER, CellType.NUMBER, CellType.DATE, CellType.DATE, CellType.NUMBER), (0..4).map { sheet.cellAt(0, it)?.type })
        assertEquals("Editing a formatted number starts from its stored value", "0.25", sheet.cellAt(0, 0)?.editableText)
        assertEquals("1234.5", sheet.cellAt(0, 1)?.editableText)
    }

    @Test fun editedPercentAndMacDateKeepTheirFormatAcrossSave() = runBlocking {
        val instrumentation = InstrumentationRegistry.getInstrumentation()
        val context = instrumentation.targetContext
        val source = File.createTempFile("formatted", ".xlsx", context.cacheDir)
        val output = File.createTempFile("edited", ".xlsx", context.cacheDir)
        try {
            source.writeBytes(xlsx(date1904 = true))
            val vm = XlsxViewerViewModel(context)
            instrumentation.runOnMainSync { vm.loadXlsx(Uri.fromFile(source), source.name) }
            withTimeout(15000) { vm.state.first { it.document != null } }
            instrumentation.runOnMainSync {
                vm.updateCellValue(0, 0, "0.5")
                vm.updateCellValue(0, 2, "1")
            }
            fun checkCells(sheet: XlsxSheet) {
                assertEquals("50%", sheet.cellAt(0, 0)?.value)
                assertEquals("0.5", sheet.cellAt(0, 0)?.editableText)
                assertEquals("1/2/1904", sheet.cellAt(0, 2)?.value)
                assertEquals("1", sheet.cellAt(0, 2)?.editableText)
            }
            checkCells(vm.state.value.document!!.sheets.single())
            instrumentation.runOnMainSync { vm.saveDocumentAs(Uri.fromFile(output)) }
            withTimeout(15000) { vm.state.first { it.saveRevision == 1 } }
            checkCells(output.inputStream().use { XlsxParser.parse(it) }.sheets.single())
        } finally {
            source.delete()
            output.delete()
        }
    }

    @Test fun macWorkbooksCountDatesFrom1904() {
        assertEquals("11/2/2027", XlsxParser.parse(xlsx(date1904 = true).inputStream()).sheets.single().cellAt(0, 2)?.value)
    }
}
