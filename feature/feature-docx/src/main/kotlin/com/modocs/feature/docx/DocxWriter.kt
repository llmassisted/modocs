package com.modocs.feature.docx

import android.content.Context
import android.net.Uri
import com.modocs.core.common.*
import org.w3c.dom.Element
import java.io.OutputStream
import java.util.zip.ZipEntry
import java.util.zip.ZipOutputStream

/** Patch edited runs in the original XML; retain sections, relationships and unknown content. */
object DocxWriter {
    private const val W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    suspend fun save(context: Context, document: DocxDocument, outputUri: Uri) =
        saveDocumentCopy(context, outputUri) { write(document, it) }

    private fun paragraphs(root: Element): List<Element> {
        val nodes = root.getElementsByTagNameNS(W, "p")
        return (0 until nodes.length).map { nodes.item(it) as Element }.filter { p ->
            var parent = p.parentNode
            var nested = false
            while (parent is Element) {
                if (parent.namespaceURI == W && parent.localName in listOf("tbl", "p")) nested = true
                parent = parent.parentNode
            }
            !nested
        }
    }

    private fun runs(p: Element): List<Element> {
        val nodes = p.getElementsByTagNameNS(W, "r")
        return (0 until nodes.length).map { nodes.item(it) as Element }
    }

    private fun runText(run: Element): String = run.elements().joinToString("") {
        when (it.localName) {
            "t" -> it.textContent
            "tab" -> "\t"
            "br" -> if (it.getAttributeNS(W, "type") == "page") "\u000C" else "\n"
            "cr" -> "\n"
            else -> ""
        }
    }

    fun attachSources(body: List<DocxElement>, bytes: ByteArray): List<DocxElement> {
        val xmlParagraphs = paragraphs(parseXmlDocument(bytes).documentElement)
        var ordinal = 0
        return body.map { element ->
            if (element !is DocxParagraph) return@map element
            val index = ordinal++
            val p = xmlParagraphs.getOrNull(index) ?: return@map element
            val xmlRuns = runs(p)
            val simple = p.elements().all { it.namespaceURI == W && it.localName in setOf("pPr", "r", "hyperlink", "bookmarkStart", "bookmarkEnd", "commentRangeStart", "commentRangeEnd", "proofErr") } &&
                xmlRuns.all { r -> r.elements().all { it.namespaceURI == W && it.localName in setOf("rPr", "t", "tab", "br", "cr") } }
            if (!simple || xmlRuns.joinToString("") { runText(it) } != element.text) return@map element
            var offset = 0
            val ends = xmlRuns.map { runText(it).length }.runningFold(0, Int::plus).drop(1)
            element.copy(sourceParagraph = index, safelyEditable = true, runs = element.runs.map { run ->
                val source = ends.indexOfFirst { it > offset }.let { if (it < 0) (xmlRuns.size - 1) else it }
                offset += run.text.length
                run.copy(sourceRun = source)
            })
        }
    }

    fun write(document: DocxDocument, outputStream: OutputStream) {
        require(document.body.size == document.originalBody.size) { "Structural editing is not supported" }
        val changes = document.body.mapIndexedNotNull { index, element ->
            if (element == document.originalBody[index]) null else {
                require(element is DocxParagraph && element.safelyEditable) { "This paragraph contains content that cannot be safely edited" }
                element
            }
        }
        var replacement: ByteArray? = null
        if (changes.isNotEmpty()) {
            val xml = parseXmlDocument(requireNotNull(document.rawEntries["word/document.xml"]))
            val xmlParagraphs = paragraphs(xml.documentElement)
            for (paragraph in changes) {
                val node = xmlParagraphs[paragraph.sourceParagraph]
                val originals = runs(node)
                val baseline = document.originalBody.filterIsInstance<DocxParagraph>().first { it.sourceParagraph == paragraph.sourceParagraph }
                if (originals.isEmpty()) {
                    val r = xml.createElementNS(W, "w:r")
                    node.appendChild(r)
                    originals.plus(r).forEachIndexed { i, source -> replaceRun(source, paragraph.runs.filter { it.sourceRun == i || it.sourceRun == -1 }, RunProperties()) }
                } else {
                    originals.forEachIndexed { i, source -> replaceRun(source, paragraph.runs.filter { it.sourceRun == i || (it.sourceRun == -1 && i == 0) }, baseline.runs.firstOrNull { it.sourceRun == i }?.properties ?: RunProperties()) }
                }
            }
            replacement = xml.xmlBytes()
        }
        ZipOutputStream(outputStream).use { zip ->
            for ((name, bytes) in document.rawEntries) {
                zip.putNextEntry(ZipEntry(name))
                zip.write(if (name == "word/document.xml") replacement ?: bytes else bytes)
                zip.closeEntry()
            }
        }
    }

    private fun replaceRun(source: Element, values: List<DocxRun>, originalProps: RunProperties) {
        if (values.isEmpty() && runText(source).isEmpty()) return
        val xml = source.ownerDocument
        for (value in values) {
            val run = source.cloneNode(true) as Element
            run.elements().filter { it.localName != "rPr" }.forEach { run.removeChild(it) }
            val props = run.childrenNamed(W, "rPr").firstOrNull()
                ?: xml.createElementNS(W, "w:rPr").also { run.insertBefore(it, run.firstChild) }
            for ((name, enabled) in listOf("b" to value.properties.bold, "i" to value.properties.italic, "u" to value.properties.underline)) {
                val original = when (name) { "b" -> originalProps.bold; "i" -> originalProps.italic; else -> originalProps.underline }
                if (enabled == original) continue
                props.childrenNamed(W, name).forEach { props.removeChild(it) }
                props.appendChild(xml.createElementNS(W, "w:$name").also {
                    it.setAttributeNS(W, "w:val", if (name == "u") { if (enabled) "single" else "none" } else if (enabled) "1" else "0")
                })
            }
            val parts = Regex("[\\t\\n\\u000C]|[^\\t\\n\\u000C]+").findAll(value.text)
            for (part in parts) {
                val tag = when (part.value) { "\t" -> "tab"; "\n", "\u000C" -> "br"; else -> "t" }
                run.appendChild(xml.createElementNS(W, "w:$tag").also {
                    if (tag == "t") {
                        it.setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:space", "preserve")
                        it.textContent = part.value
                    } else if (part.value == "\u000C") it.setAttributeNS(W, "w:type", "page")
                })
            }
            source.parentNode.insertBefore(run, source)
        }
        source.parentNode.removeChild(source)
    }
}
