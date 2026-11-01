package com.modocs.app

import androidx.test.ext.junit.runners.AndroidJUnit4
import com.modocs.core.common.readZipEntriesCapped
import com.modocs.feature.docx.*
import org.junit.Assert.*
import org.junit.Test
import org.junit.runner.RunWith
import java.io.ByteArrayOutputStream
import java.util.zip.ZipEntry
import java.util.zip.ZipOutputStream

/** Pictures in the shapes Word actually writes them; documents are built here, so nothing private is checked in. */
@RunWith(AndroidJUnit4::class)
class DocxImageTest {
    private val png = byteArrayOf(0x89.toByte(), 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A)
    private val emf = byteArrayOf(1, 0, 0, 0, 0x6C, 0, 0, 0)
    private val namespaces = """xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
        xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
        xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office""""

    private fun rels(vararg targets: Pair<String, String>) =
        """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">""" +
            targets.joinToString("") { (id, target) -> """<Relationship Id="$id" Type="t" Target="$target"/>""" } + "</Relationships>"

    private fun docx(parts: Map<String, Any>): ByteArray = ByteArrayOutputStream().also { out ->
        ZipOutputStream(out).use { zip ->
            for ((name, content) in parts) {
                zip.putNextEntry(ZipEntry(name))
                zip.write(if (content is ByteArray) content else content.toString().toByteArray())
                zip.closeEntry()
            }
        }
    }.toByteArray()

    private fun document(body: String) = """<?xml version="1.0" encoding="UTF-8"?><w:document $namespaces><w:body>$body</w:body></w:document>"""

    private fun inline(relId: String, cx: Long = 1905000, cy: Long = 952500, descr: String = "") =
        """<w:drawing><wp:inline><wp:extent cx="$cx" cy="$cy"/><wp:docPr id="1" name="Picture" descr="$descr"/>
        <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>
        <pic:blipFill><a:blip r:embed="$relId"/></pic:blipFill><pic:spPr><a:xfrm><a:ext cx="$cx" cy="$cy"/></a:xfrm></pic:spPr>
        </pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>"""

    private fun parse(bytes: ByteArray) = DocxParser.parse(bytes.inputStream())

    @Test fun inlinePicturesInRunsAreShownWithSizeAltTextAndAlignment() {
        val doc = parse(docx(mapOf(
            "word/document.xml" to document("""<w:p><w:r><w:t>Before</w:t></w:r></w:p>
                <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r>${inline("rId4", descr = "Logo")}</w:r></w:p>
                <w:p><w:hyperlink r:id="rId9"><w:r>${inline("rId5")}</w:r></w:hyperlink></w:p>"""),
            "word/_rels/document.xml.rels" to rels("rId4" to "media/image1.png", "rId5" to "../word/media/image2.png", "rId9" to "media/missing.png"),
            "word/media/image1.png" to png, "word/media/image2.png" to png,
        )))
        val images = doc.body.filterIsInstance<DocxImage>()
        assertEquals(listOf("word/media/image1.png", "word/media/image2.png"), images.map { it.partName })
        assertEquals(DocxImage("word/media/image1.png", 1905000, 952500, "Logo", ParagraphAlignment.CENTER), images[0])
        assertEquals(ParagraphAlignment.LEFT, images[1].alignment)
        assertEquals(setOf("word/media/image1.png", "word/media/image2.png"), doc.images.keys)
        assertTrue(doc.warnings.isEmpty())
    }

    @Test fun alternateContentAndTextBoxesAppearOnce() {
        val doc = parse(docx(mapOf(
            "word/document.xml" to document("""<w:p><w:r><mc:AlternateContent>
                <mc:Choice Requires="wps">${inline("rId4")}</mc:Choice>
                <mc:Fallback><w:pict><v:shape style="width:150pt;height:75pt"><v:imagedata r:id="rId4"/></v:shape></w:pict></mc:Fallback>
                </mc:AlternateContent></w:r></w:p>
                <w:p><mc:AlternateContent><mc:Choice Requires="w99"><w99:future xmlns:w99="urn:future"/></mc:Choice>
                <mc:Fallback><w:r><w:t>Old style</w:t></w:r></mc:Fallback></mc:AlternateContent></w:p>
                <w:p><w:r><w:t>Note: </w:t></w:r><w:r><mc:AlternateContent>
                <mc:Choice Requires="wps"><w:drawing><wp:anchor><wp:extent cx="10" cy="10"/><a:graphic><a:graphicData uri="wps">
                <wps:wsp><wps:txbx><w:txbxContent><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Boxed</w:t></w:r></w:p><w:p><w:r><w:t>text</w:t></w:r></w:p></w:txbxContent></wps:txbx></wps:wsp>
                </a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>
                <mc:Fallback><w:pict><v:shape><v:textbox><w:txbxContent><w:p><w:r><w:t>Boxed</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback>
                </mc:AlternateContent></w:r><w:r><w:t> end</w:t></w:r></w:p>"""),
            "word/_rels/document.xml.rels" to rels("rId4" to "media/image1.png"),
            "word/media/image1.png" to png,
        )))
        assertEquals(1, doc.body.count { it is DocxImage })
        val unreadable = doc.body.filterIsInstance<DocxParagraph>()[1]
        assertEquals("Runs only in a Fallback are still read", "Old style", unreadable.text)
        val note = doc.body.filterIsInstance<DocxParagraph>().last()
        assertEquals("Note: Boxed\ntext end", note.text)
        assertEquals("Boxed\ntext", note.runs[1].text)
        assertFalse("Text-box formatting must not leak into the anchoring run", note.runs[1].properties.bold)
    }

    @Test fun legacyVmlPicturesTablesHeadersAndPlaceholders() {
        val doc = parse(docx(mapOf(
            "word/document.xml" to document("""
                <w:p><w:r><w:pict><v:shape style="width:2in;height:72pt;mso-position-horizontal:center"><v:imagedata r:id="rId4" o:title="Old"/></v:shape></w:pict></w:r></w:p>
                <w:tbl><w:tr><w:tc><w:p><w:r><w:t>Photo</w:t></w:r></w:p><w:p><w:r>${inline("rId5")}</w:r></w:p></w:tc></w:tr></w:tbl>
                <w:p><w:r><w:drawing><wp:inline><wp:extent cx="100" cy="50"/><wp:docPr id="2" name="Chart 2"/>
                <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"/></a:graphic></wp:inline></w:drawing></w:r></w:p>
                <w:p><w:r><w:drawing><wp:inline><wp:extent cx="100" cy="50"/><a:graphic><a:graphicData uri="p"><pic:pic><pic:blipFill><a:blip r:link="rId8"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>
                <w:sectPr><w:headerReference w:type="default" r:id="rId1"/></w:sectPr>"""),
            "word/_rels/document.xml.rels" to rels("rId1" to "header1.xml", "rId4" to "media/old.png", "rId5" to "media/cell.emf"),
            "word/header1.xml" to """<w:hdr $namespaces><w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r>${inline("rId1")}</w:r><w:r><w:t>ACME</w:t></w:r></w:p></w:hdr>""",
            "word/_rels/header1.xml.rels" to rels("rId1" to "media/logo.png"),
            "word/media/old.png" to png, "word/media/cell.emf" to emf, "word/media/logo.png" to png,
        )))
        val body = doc.body.filterIsInstance<DocxImage>()
        assertEquals(DocxImage("word/media/old.png", 144L * 12700, 72L * 12700, "Old", ParagraphAlignment.CENTER), body[0])
        assertEquals(listOf("Chart", ""), body.drop(1).map { it.altText.ifEmpty { it.partName } })
        assertEquals(listOf("word/media/cell.emf"), doc.body.filterIsInstance<DocxTable>().single().rows[0].cells[0].images.map { it.partName })
        // The header's rId1 is its own logo, not the body's rId1 (the header part itself).
        assertEquals(listOf(DocxImage("word/media/logo.png", 1905000, 952500, "", ParagraphAlignment.RIGHT)), doc.headerImages)
        assertEquals("ACME", doc.headerParagraphs.single().text)
        assertEquals(2, doc.warnings.size)
        assertTrue(doc.warnings.any { "placeholders (2)" in it })
        assertTrue(doc.warnings.any { "such as EMF or WMF (1)" in it })
    }

    @Test fun editingTextNextToPicturesPatchesTheRightParagraph() {
        val original = parse(docx(mapOf(
            "word/document.xml" to document("""<w:p><w:r><w:t>Intro</w:t></w:r></w:p><w:p><w:r>${inline("rId4")}</w:r></w:p><w:p><w:r><w:t>After</w:t></w:r></w:p>"""),
            "word/_rels/document.xml.rels" to rels("rId4" to "media/image1.png"),
            "word/media/image1.png" to png,
        )))
        val pictureParagraph = original.body[1] as DocxParagraph
        assertFalse("A paragraph holding a picture must stay read-only", pictureParagraph.safelyEditable)
        val index = original.body.indexOfLast { it is DocxParagraph }
        val after = original.body[index] as DocxParagraph
        assertTrue(after.safelyEditable)
        val body = original.body.toMutableList()
        body[index] = after.copy(runs = replaceParagraphText(after.runs, "After edit"))
        val saved = ByteArrayOutputStream().also { DocxWriter.write(original.copy(body = body), it) }.toByteArray()
        val reopened = parse(saved)
        assertEquals(listOf("Intro", "", "After edit"), reopened.body.filterIsInstance<DocxParagraph>().map { it.text })
        assertEquals(1, reopened.body.count { it is DocxImage })
        assertArrayEquals(png, readZipEntriesCapped(saved.inputStream())["word/media/image1.png"])
    }
}
