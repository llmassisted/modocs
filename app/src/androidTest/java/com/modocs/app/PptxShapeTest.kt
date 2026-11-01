package com.modocs.app

import androidx.test.ext.junit.runners.AndroidJUnit4
import com.modocs.feature.pptx.*
import org.junit.Assert.*
import org.junit.Test
import org.junit.runner.RunWith
import java.io.ByteArrayOutputStream
import java.util.zip.ZipEntry
import java.util.zip.ZipOutputStream

/** Shapes as PowerPoint writes them, including the extension lists it adds after a shape's size. */
@RunWith(AndroidJUnit4::class)
class PptxShapeTest {
    private val ns = """xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""""

    private fun pptx(slide: String): ByteArray = ByteArrayOutputStream().also { out ->
        ZipOutputStream(out).use { zip ->
            fun put(name: String, content: Any) {
                zip.putNextEntry(ZipEntry(name))
                zip.write(if (content is ByteArray) content else content.toString().toByteArray())
                zip.closeEntry()
            }
            put("ppt/presentation.xml", """<p:presentation $ns><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>""")
            put("ppt/_rels/presentation.xml.rels", """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="t" Target="slides/slide1.xml"/></Relationships>""")
            put("ppt/slides/slide1.xml", """<p:sld $ns><p:cSld><p:spTree>$slide</p:spTree></p:cSld></p:sld>""")
            put("ppt/slides/_rels/slide1.xml.rels", """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="t" Target="../media/image1.png"/></Relationships>""")
            put("ppt/media/image1.png", byteArrayOf(0x89.toByte(), 0x50, 0x4E, 0x47))
        }
    }.toByteArray()

    private val hiddenFill = """<a:extLst><a:ext uri="{909E8E84-426E-40DD-AFC4-6F175D3DCCD1}"><a14:hiddenFill xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a14:hiddenFill></a:ext></a:extLst>"""

    @Test fun extensionListsAfterTheSizeDoNotZeroIt() {
        val slide = PptxParser.parse(pptx("""
            <p:pic><p:nvPicPr><p:cNvPr id="4" name="Picture 3"><a:extLst><a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"/></a:extLst></p:cNvPr><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
            <p:blipFill><a:blip r:embed="rId2"><a:extLst><a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}"/></a:extLst></a:blip></p:blipFill>
            <p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="3000" cy="2000"/></a:xfrm><a:prstGeom prst="rect"/><a:noFill/>$hiddenFill</p:spPr></p:pic>
            <p:sp><p:nvSpPr><p:cNvPr id="5" name="Box"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
            <p:spPr><a:xfrm><a:off x="10" y="20"/><a:ext cx="500" cy="400"/></a:xfrm><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>$hiddenFill</p:spPr></p:sp>
        """).inputStream()).slides.single()
        val picture = slide.shapes.filterIsInstance<PptxImageShape>().single()
        assertEquals(listOf(100L, 200L, 3000L, 2000L), listOf(picture.x, picture.y, picture.width, picture.height))
        val box = slide.shapes.filterIsInstance<PptxRectangle>().single()
        assertEquals(listOf(10L, 20L, 500L, 400L), listOf(box.x, box.y, box.width, box.height))
    }
}
