package com.modocs.app

import android.net.Uri
import androidx.test.ext.junit.runners.AndroidJUnit4
import androidx.test.platform.app.InstrumentationRegistry
import com.modocs.core.common.*
import com.modocs.feature.docx.*
import com.modocs.feature.xlsx.*
import com.modocs.feature.pdf.*
import com.tom_roush.pdfbox.android.PDFBoxResourceLoader
import com.tom_roush.pdfbox.pdmodel.*
import com.tom_roush.pdfbox.pdmodel.common.PDRectangle
import com.tom_roush.pdfbox.pdmodel.font.PDType1Font
import com.tom_roush.pdfbox.pdmodel.interactive.form.*
import com.tom_roush.pdfbox.text.PDFTextStripper
import kotlinx.coroutines.runBlocking
import kotlinx.coroutines.withTimeout
import kotlinx.coroutines.flow.first
import org.junit.Assert.*
import org.junit.Test
import org.junit.runner.RunWith
import java.io.ByteArrayOutputStream
import java.io.File

@RunWith(AndroidJUnit4::class)
class DocumentRegressionTest {
    private val instrumentation = InstrumentationRegistry.getInstrumentation()
    private val context = instrumentation.targetContext
    private fun asset(name: String) = instrumentation.context.assets.open(name)

    @Test fun docxEditsPreserveSectionsHeadersHyperlinksAndRunFormatting() {
        val original = asset("preservation.docx").use { DocxParser.parse(it) }
        val paragraph = original.body[0] as DocxParagraph
        assertEquals("Preserved header", original.headerParagraphs.single().text)
        assertTrue(paragraph.safelyEditable)
        assertFalse((original.body[1] as DocxParagraph).safelyEditable)
        val body = original.body.toMutableList()
        body[0] = paragraph.copy(runs = replaceParagraphText(paragraph.runs, "Hello brave world"))
        val bytes = ByteArrayOutputStream().also { DocxWriter.write(original.copy(body = body), it) }.toByteArray()
        val saved = DocxParser.parse(bytes.inputStream())
        assertEquals("Hello brave world", (saved.body[0] as DocxParagraph).text)
        assertArrayEquals(original.rawEntries["word/header1.xml"], saved.rawEntries["word/header1.xml"])
        assertArrayEquals(original.rawEntries["word/_rels/document.xml.rels"], saved.rawEntries["word/_rels/document.xml.rels"])
        val xml = parseXmlDocument(saved.rawEntries.getValue("word/document.xml"))
        val w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        assertEquals(1, xml.getElementsByTagNameNS(w, "sectPr").length)
        assertEquals(1, xml.getElementsByTagNameNS(w, "headerReference").length)
        assertEquals(1, xml.getElementsByTagNameNS(w, "fldSimple").length)
        assertEquals(1, xml.getElementsByTagNameNS(w, "hyperlink").length)
        assertEquals(original.pageSetup, saved.pageSetup)
        assertTrue((saved.body[0] as DocxParagraph).runs.last().properties.italic)
        val unchanged = ByteArrayOutputStream().also { DocxWriter.write(original, it) }.toByteArray()
        assertArrayEquals(original.rawEntries.getValue("word/document.xml"), readZipEntriesCapped(unchanged.inputStream())["word/document.xml"])
    }

    @Test fun xlsxEditsKeepValidationFormattingViewsAndFormulaCaches() {
        val original = asset("preservation.xlsx").use { XlsxParser.parse(it) }
        val rows = original.sheets[0].rows.map { row -> row.copy(cells = row.cells.map { cell ->
            if (row.rowIndex == 1 && cell.columnIndex == 0) cell.copy(value = "15", rawValue = null) else cell
        }.toMutableList()) }.toMutableList()
        val changed = original.copy(sheets = listOf(original.sheets[0].copy(rows = rows)),
            modifiedSheets = mutableSetOf(0), editedCells = mapOf(0 to setOf(1 to 0)))
        val bytes = ByteArrayOutputStream().also { XlsxWriter.write(changed, it) }.toByteArray()
        val saved = XlsxParser.parse(bytes.inputStream())
        assertEquals("15", saved.sheets[0].cellAt(1, 0)?.value)
        assertEquals("A2*2", saved.sheets[0].cellAt(1, 1)?.formula)
        assertEquals("20", saved.sheets[0].cellAt(1, 1)?.value)
        assertEquals(original.sheets[0].mergedCells, saved.sheets[0].mergedCells)
        assertEquals(1, saved.sheets[0].frozenRows)
        val ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        val xml = parseXmlDocument(saved.rawEntries.getValue("xl/worksheets/sheet1.xml"))
        for (tag in listOf("dataValidations", "conditionalFormatting", "pageMargins", "sheetViews"))
            assertEquals(tag, 1, xml.getElementsByTagNameNS(ns, tag).length)
        val workbook = parseXmlDocument(saved.rawEntries.getValue("xl/workbook.xml"))
        assertEquals("1", (workbook.getElementsByTagNameNS(ns, "calcPr").item(0) as org.w3c.dom.Element).getAttribute("fullCalcOnLoad"))
    }

    @Test fun pdfCopyPreservesSearchablePagesRotationAndFormValues() = runBlocking {
        PDFBoxResourceLoader.init(context)
        val source = File(context.cacheDir, "regression-source.pdf")
        val output = File(context.cacheDir, "regression-output.pdf")
        try {
            PDDocument().use { doc ->
                for (rotation in listOf(0, 90, 180, 270)) {
                    val page = PDPage(PDRectangle.LETTER)
                    page.rotation = rotation
                    page.cropBox = PDRectangle(10f, 20f, 580f, 740f)
                    doc.addPage(page)
                    PDPageContentStream(doc, page).use { stream ->
                        stream.beginText(); stream.setFont(PDType1Font.HELVETICA, 12f)
                        stream.newLineAtOffset(40f, 500f); stream.showText("Searchable original $rotation"); stream.endText()
                    }
                }
                val form = PDAcroForm(doc)
                doc.documentCatalog.acroForm = form
                val resources = PDResources()
                resources.put(com.tom_roush.pdfbox.cos.COSName.getPDFName("Helv"), PDType1Font.HELVETICA)
                form.defaultResources = resources
                form.defaultAppearance = "/Helv 12 Tf 0 g"
                val field = PDTextField(form)
                field.partialName = "name"
                field.widgets[0].rectangle = PDRectangle(40f, 600f, 200f, 24f)
                field.widgets[0].page = doc.getPage(0)
                doc.getPage(0).annotations.add(field.widgets[0])
                form.fields.add(field)
                doc.save(source)
            }
            assertEquals("name", PdfForms.read(context, Uri.fromFile(source)).single().name)
            val annotations = (0..3).map { TextAnnotation("$it", it, .15f, .2f, "Added mark") }
            PdfDocumentWriter.save(context, Uri.fromFile(source), Uri.fromFile(output), 4, annotations, mapOf("name" to "MoDocs Test"))
            PDDocument.load(output).use { saved ->
                assertEquals(4, saved.numberOfPages)
                assertTrue(PDFTextStripper().getText(saved).contains("Searchable original"))
                assertEquals("MoDocs Test", saved.documentCatalog.acroForm.getField("name").valueAsString)
                for (i in 0..3) {
                    assertEquals(612f, saved.getPage(i).mediaBox.width)
                    assertEquals(792f, saved.getPage(i).mediaBox.height)
                    assertEquals(i * 90, saved.getPage(i).rotation)
                    assertEquals(10f, saved.getPage(i).cropBox.lowerLeftX)
                    assertEquals(580f, saved.getPage(i).cropBox.width)
                }
            }
            PDDocument.load(source).use { assertEquals("", it.documentCatalog.acroForm.getField("name").valueAsString) }
        } finally { source.delete(); output.delete() }
    }

    @Test fun saveAsCommitsDraftAndTracksCopyWithoutChangingOriginal() = runBlocking {
        val source = File(context.cacheDir, "workflow-source.xlsx")
        val copy = File(context.cacheDir, "workflow-copy.xlsx")
        try {
            asset("preservation.xlsx").use { input -> source.outputStream().use { input.copyTo(it) } }
            lateinit var vm: XlsxViewerViewModel
            instrumentation.runOnMainSync {
                vm = XlsxViewerViewModel(context)
                vm.loadXlsx(Uri.fromFile(source), source.name)
            }
            withTimeout(15000) { vm.state.first { !it.isLoading } }
            assertNotNull(vm.state.value.document)
            instrumentation.runOnMainSync {
                vm.toggleEditMode(); vm.startEditingCell(1, 0); vm.updateDraft("42")
                vm.saveDocumentAs(Uri.fromFile(copy))
            }
            withTimeout(15000) { vm.state.first { it.saveRevision == 1 || it.errorMessage != null } }
            assertEquals(Uri.fromFile(copy), vm.state.value.activeUri)
            assertFalse(vm.state.value.isDirty)
            assertFalse(vm.state.value.draftDirty)
            assertEquals("42", copy.inputStream().use { XlsxParser.parse(it) }.sheets[0].cellAt(1, 0)?.value)
            assertEquals("10", source.inputStream().use { XlsxParser.parse(it) }.sheets[0].cellAt(1, 0)?.value)
            instrumentation.runOnMainSync { vm.undo() }
            assertEquals("10", vm.state.value.document!!.sheets[0].cellAt(1, 0)?.value)
            assertTrue(vm.state.value.isDirty)
        } finally { source.delete(); copy.delete() }
    }

    @Test fun partialPresentationReportsSkippedSlides() {
        val presentation = asset("partial.pptx").use { com.modocs.feature.pptx.PptxParser.parse(it) }
        assertEquals(1, presentation.slides.size)
        assertTrue(presentation.warnings.first().contains("1 of 2"))
        assertTrue(presentation.warnings.any { it.contains("slide 1") })
    }

    @Test fun preferencesPersistAndPositionOptOutClearsHistory() {
        val store = AppPreferences.get(context)
        val original = store.state.value
        try {
            store.update(original.copy(theme = "Dark", rememberPosition = true, spreadsheetZoom = 1.25f))
            store.savePosition("synthetic-regression-document", 8)
            assertEquals(8, store.position("synthetic-regression-document"))
            assertEquals("Dark", context.getSharedPreferences("reader_settings", 0).getString("theme", ""))
            store.update(store.state.value.copy(rememberPosition = false))
            assertEquals(0, store.position("synthetic-regression-document"))
            assertTrue(context.getSharedPreferences("reading_positions", 0).all.isEmpty())
        } finally { store.update(original) }
    }

    @Test fun malformedOptionalPartIsReportedAndXmlEntitiesAreRejected() {
        val doc = asset("preservation.xlsx").use { XlsxParser.parse(it) }
        val archive = ByteArrayOutputStream()
        java.util.zip.ZipOutputStream(archive).use { zip ->
            for ((name, bytes) in doc.rawEntries + ("xl/styles.xml" to "<broken>".toByteArray())) {
                zip.putNextEntry(java.util.zip.ZipEntry(name)); zip.write(bytes); zip.closeEntry()
            }
        }
        val partial = XlsxParser.parse(archive.toByteArray().inputStream())
        assertFalse(partial.warnings.isEmpty())
        assertEquals(1, partial.sheets.size)
        try {
            parseXmlDocument("<!DOCTYPE x [<!ENTITY test SYSTEM 'file:///etc/passwd'>]><x>&test;</x>".toByteArray())
            fail("External entities must be rejected")
        } catch (_: org.xml.sax.SAXException) { }
    }
}
