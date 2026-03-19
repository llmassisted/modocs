package com.modocs.feature.docx

import android.content.Context
import android.net.Uri
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.withContext
import java.io.ByteArrayOutputStream
import java.io.OutputStream
import java.util.zip.ZipEntry
import java.util.zip.ZipOutputStream

/**
 * Writes a [DocxDocument] back to a DOCX ZIP file.
 * Uses round-trip approach: copies all original ZIP entries unchanged,
 * except word/document.xml which is regenerated from the in-memory model.
 */
object DocxWriter {

    private const val NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    private const val NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    private const val NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"

    suspend fun save(context: Context, document: DocxDocument, outputUri: Uri) {
        withContext(Dispatchers.IO) {
            context.contentResolver.openOutputStream(outputUri)?.use { stream ->
                write(document, stream)
            } ?: throw IllegalStateException("Cannot open output URI for writing")
        }
    }

    fun write(document: DocxDocument, outputStream: OutputStream) {
        val newDocumentXml = generateDocumentXml(document)

        ZipOutputStream(outputStream).use { zip ->
            // Write all original entries except document.xml
            for ((name, bytes) in document.rawEntries) {
                if (name == "word/document.xml") continue
                zip.putNextEntry(ZipEntry(name))
                zip.write(bytes)
                zip.closeEntry()
            }

            // Write the updated document.xml
            zip.putNextEntry(ZipEntry("word/document.xml"))
            zip.write(newDocumentXml.toByteArray(Charsets.UTF_8))
            zip.closeEntry()
        }
    }

    private fun generateDocumentXml(document: DocxDocument): String {
        val sb = StringBuilder()
        sb.append("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>""")
        sb.append("\n")
        sb.append("""<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" """)
        sb.append("""xmlns:mo="http://schemas.microsoft.com/office/mac/office/2008/main" """)
        sb.append("""xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" """)
        sb.append("""xmlns:mv="urn:schemas-microsoft-com:mac:vml" """)
        sb.append("""xmlns:o="urn:schemas-microsoft-com:office:office" """)
        sb.append("""xmlns:r="$NS_R" """)
        sb.append("""xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" """)
        sb.append("""xmlns:v="urn:schemas-microsoft-com:vml" """)
        sb.append("""xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" """)
        sb.append("""xmlns:wp="$NS_WP" """)
        sb.append("""xmlns:w10="urn:schemas-microsoft-com:office:word" """)
        sb.append("""xmlns:w="$NS_W" """)
        sb.append("""xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" """)
        sb.append("""xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" """)
        sb.append("""xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" """)
        sb.append("""xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" """)
        sb.append("""xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">""")
        sb.append("\n<w:body>\n")

        for (element in document.body) {
            when (element) {
                is DocxParagraph -> writeParagraph(sb, element)
                is DocxTable -> writeTable(sb, element)
                is DocxImage -> writeImage(sb, element)
            }
        }

        sb.append("</w:body>\n</w:document>")
        return sb.toString()
    }

    private fun writeParagraph(sb: StringBuilder, paragraph: DocxParagraph) {
        sb.append("<w:p>")

        // Paragraph properties
        val props = paragraph.properties
        if (hasParaProps(props) || paragraph.listInfo != null) {
            sb.append("<w:pPr>")

            if (props.styleId != null) {
                sb.append("""<w:pStyle w:val="${escXml(props.styleId!!)}"/>""")
            }

            if (props.alignment != ParagraphAlignment.LEFT) {
                val jcVal = when (props.alignment) {
                    ParagraphAlignment.CENTER -> "center"
                    ParagraphAlignment.RIGHT -> "right"
                    ParagraphAlignment.JUSTIFY -> "both"
                    ParagraphAlignment.LEFT -> "left"
                }
                sb.append("""<w:jc w:val="$jcVal"/>""")
            }

            if (props.spacingBeforePt > 0f || props.spacingAfterPt != 8f || props.lineSpacingMultiplier != 1.15f) {
                sb.append("<w:spacing")
                if (props.spacingBeforePt > 0f) {
                    sb.append(""" w:before="${(props.spacingBeforePt * 20).toInt()}"""")
                }
                sb.append(""" w:after="${(props.spacingAfterPt * 20).toInt()}"""")
                if (props.lineSpacingMultiplier != 1.15f) {
                    sb.append(""" w:line="${(props.lineSpacingMultiplier * 240).toInt()}"""")
                }
                sb.append("/>")
            }

            if (props.indentLeftTwips > 0 || props.indentRightTwips > 0 || props.indentFirstLineTwips != 0) {
                sb.append("<w:ind")
                if (props.indentLeftTwips > 0) sb.append(""" w:left="${props.indentLeftTwips}"""")
                if (props.indentRightTwips > 0) sb.append(""" w:right="${props.indentRightTwips}"""")
                if (props.indentFirstLineTwips > 0) sb.append(""" w:firstLine="${props.indentFirstLineTwips}"""")
                if (props.indentFirstLineTwips < 0) sb.append(""" w:hanging="${-props.indentFirstLineTwips}"""")
                sb.append("/>")
            }

            if (props.headingLevel > 0 && props.styleId == null) {
                sb.append("""<w:outlineLvl w:val="${props.headingLevel - 1}"/>""")
            }

            if (props.keepNext) sb.append("<w:keepNext/>")
            if (props.keepLines) sb.append("<w:keepLines/>")
            if (props.pageBreakBefore) sb.append("<w:pageBreakBefore/>")

            if (paragraph.listInfo != null) {
                sb.append("<w:numPr>")
                sb.append("""<w:ilvl w:val="${paragraph.listInfo!!.level}"/>""")
                sb.append("""<w:numId w:val="${escXml(paragraph.listInfo!!.numId)}"/>""")
                sb.append("</w:numPr>")
            }

            sb.append("</w:pPr>")
        }

        // Runs
        for (run in paragraph.runs) {
            writeRun(sb, run)
        }

        sb.append("</w:p>\n")
    }

    private fun writeRun(sb: StringBuilder, run: DocxRun) {
        if (run.isPageBreak) {
            sb.append("""<w:r><w:br w:type="page"/></w:r>""")
            return
        }
        if (run.isBreak) {
            sb.append("<w:r><w:br/></w:r>")
            return
        }
        if (run.text == "\t") {
            sb.append("<w:r>")
            writeRunProperties(sb, run.properties)
            sb.append("<w:tab/></w:r>")
            return
        }

        sb.append("<w:r>")
        writeRunProperties(sb, run.properties)
        // Preserve whitespace with xml:space
        sb.append("""<w:t xml:space="preserve">${escXml(run.text)}</w:t>""")
        sb.append("</w:r>")
    }

    private fun writeRunProperties(sb: StringBuilder, rp: RunProperties) {
        if (!hasRunProps(rp)) return

        sb.append("<w:rPr>")
        if (rp.bold) sb.append("<w:b/>")
        if (rp.italic) sb.append("<w:i/>")
        if (rp.underline) sb.append("""<w:u w:val="single"/>""")
        if (rp.strikethrough) sb.append("<w:strike/>")
        if (rp.fontName != null) {
            sb.append("""<w:rFonts w:ascii="${escXml(rp.fontName!!)}" w:hAnsi="${escXml(rp.fontName!!)}"/>""")
        }
        if (rp.fontSizeHalfPt != null) {
            sb.append("""<w:sz w:val="${rp.fontSizeHalfPt}"/>""")
        }
        if (rp.color != null) {
            val hex = String.format("%06X", rp.color!! and 0xFFFFFF)
            sb.append("""<w:color w:val="$hex"/>""")
        }
        if (rp.highlight != null) {
            sb.append("""<w:highlight w:val="yellow"/>""")
        }
        if (rp.superscript) sb.append("""<w:vertAlign w:val="superscript"/>""")
        if (rp.subscript) sb.append("""<w:vertAlign w:val="subscript"/>""")
        sb.append("</w:rPr>")
    }

    private fun writeTable(sb: StringBuilder, table: DocxTable) {
        sb.append("<w:tbl>")

        // Table properties
        sb.append("<w:tblPr>")
        if (table.properties.widthTwips != null) {
            sb.append("""<w:tblW w:w="${table.properties.widthTwips}" w:type="dxa"/>""")
        }
        writeBorderElement(sb, "tblBorders", table.properties.borders)
        sb.append("</w:tblPr>")

        // Grid columns
        if (table.gridColWidths.isNotEmpty()) {
            sb.append("<w:tblGrid>")
            for (w in table.gridColWidths) {
                sb.append("""<w:gridCol w:w="$w"/>""")
            }
            sb.append("</w:tblGrid>")
        }

        for (row in table.rows) {
            sb.append("<w:tr>")
            for (cell in row.cells) {
                sb.append("<w:tc>")
                sb.append("<w:tcPr>")
                if (cell.widthTwips != null) {
                    sb.append("""<w:tcW w:w="${cell.widthTwips}" w:type="dxa"/>""")
                }
                if (cell.gridSpan > 1) {
                    sb.append("""<w:gridSpan w:val="${cell.gridSpan}"/>""")
                }
                when (cell.vMerge) {
                    VMergeType.RESTART -> sb.append("""<w:vMerge w:val="restart"/>""")
                    VMergeType.CONTINUE -> sb.append("<w:vMerge/>")
                    VMergeType.NONE -> {}
                }
                if (cell.shading != null) {
                    val hex = String.format("%06X", cell.shading!! and 0xFFFFFF)
                    sb.append("""<w:shd w:val="clear" w:color="auto" w:fill="$hex"/>""")
                }
                if (cell.borders != null) {
                    writeCellBorders(sb, cell.borders!!)
                }
                sb.append("</w:tcPr>")

                for (para in cell.paragraphs) {
                    writeParagraph(sb, para)
                }
                if (cell.paragraphs.isEmpty()) {
                    sb.append("<w:p/>")
                }
                sb.append("</w:tc>")
            }
            sb.append("</w:tr>")
        }

        sb.append("</w:tbl>\n")
    }

    private fun writeBorderElement(sb: StringBuilder, tag: String, borders: TableBorders) {
        sb.append("<w:$tag>")
        writeSingleBorder(sb, "top", borders.top)
        writeSingleBorder(sb, "left", borders.left)
        writeSingleBorder(sb, "bottom", borders.bottom)
        writeSingleBorder(sb, "right", borders.right)
        writeSingleBorder(sb, "insideH", borders.insideH)
        writeSingleBorder(sb, "insideV", borders.insideV)
        sb.append("</w:$tag>")
    }

    private fun writeCellBorders(sb: StringBuilder, borders: CellBorders) {
        sb.append("<w:tcBorders>")
        borders.top?.let { writeSingleBorder(sb, "top", it) }
        borders.left?.let { writeSingleBorder(sb, "left", it) }
        borders.bottom?.let { writeSingleBorder(sb, "bottom", it) }
        borders.right?.let { writeSingleBorder(sb, "right", it) }
        sb.append("</w:tcBorders>")
    }

    private fun writeSingleBorder(sb: StringBuilder, name: String, border: CellBorder) {
        val valStr = when (border.style) {
            BorderStyle.NONE -> "none"
            BorderStyle.SINGLE -> "single"
            BorderStyle.DOUBLE -> "double"
            BorderStyle.DASHED -> "dashed"
            BorderStyle.DOTTED -> "dotted"
            BorderStyle.DASH_SMALL_GAP -> "dashSmallGap"
        }
        val colorStr = if (border.color != null) {
            String.format("%06X", border.color and 0xFFFFFF)
        } else "auto"
        sb.append("""<w:$name w:val="$valStr" w:sz="${border.widthEighthPt}" w:space="0" w:color="$colorStr"/>""")
    }

    private fun writeImage(sb: StringBuilder, image: DocxImage) {
        // Write a paragraph containing the inline drawing
        sb.append("<w:p><w:r><w:drawing>")
        sb.append("""<wp:inline distT="0" distB="0" distL="0" distR="0">""")
        sb.append("""<wp:extent cx="${image.widthEmu}" cy="${image.heightEmu}"/>""")
        sb.append("""<wp:docPr id="1" name="Image" descr="${escXml(image.altText)}"/>""")
        sb.append("<a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">")
        sb.append("<a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">")
        sb.append("<pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">")
        sb.append("<pic:nvPicPr><pic:cNvPr id=\"0\" name=\"Image\"/><pic:cNvPicPr/></pic:nvPicPr>")
        sb.append("""<pic:blipFill><a:blip r:embed="${escXml(image.relationId)}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>""")
        sb.append("<pic:spPr><a:xfrm>")
        sb.append("""<a:off x="0" y="0"/>""")
        sb.append("""<a:ext cx="${image.widthEmu}" cy="${image.heightEmu}"/>""")
        sb.append("</a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></pic:spPr>")
        sb.append("</pic:pic></a:graphicData></a:graphic>")
        sb.append("</wp:inline></w:drawing></w:r></w:p>\n")
    }

    private fun hasParaProps(props: ParagraphProperties): Boolean {
        return props.styleId != null ||
            props.alignment != ParagraphAlignment.LEFT ||
            props.spacingBeforePt > 0f ||
            props.spacingAfterPt != 8f ||
            props.lineSpacingMultiplier != 1.15f ||
            props.indentLeftTwips > 0 ||
            props.indentRightTwips > 0 ||
            props.indentFirstLineTwips != 0 ||
            props.headingLevel > 0 ||
            props.keepNext ||
            props.keepLines ||
            props.pageBreakBefore
    }

    private fun hasRunProps(rp: RunProperties): Boolean {
        return rp.bold || rp.italic || rp.underline || rp.strikethrough ||
            rp.fontName != null || rp.fontSizeHalfPt != null ||
            rp.color != null || rp.highlight != null ||
            rp.superscript || rp.subscript
    }

    private fun escXml(text: String): String {
        return text.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace("\"", "&quot;")
            .replace("'", "&apos;")
    }
}
