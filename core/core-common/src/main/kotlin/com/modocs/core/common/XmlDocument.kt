package com.modocs.core.common

import org.w3c.dom.Document
import org.w3c.dom.Element
import org.w3c.dom.Node
import java.io.ByteArrayInputStream
import java.io.ByteArrayOutputStream
import javax.xml.parsers.DocumentBuilderFactory
import javax.xml.transform.TransformerFactory
import javax.xml.transform.dom.DOMSource
import javax.xml.transform.stream.StreamResult

/** Namespace-aware edits preserve elements the lightweight viewers do not model. */
fun parseXmlDocument(bytes: ByteArray): Document {
    val factory = DocumentBuilderFactory.newInstance()
    factory.isNamespaceAware = true
    factory.isExpandEntityReferences = false
    // Android's DOM factory does not implement the Xerces security feature flags.
    // Validate with the pull reader first; it rejects DTDs before any DOM parse.
    val validator = createXmlParser(bytes)
    while (validator.next() != org.xmlpull.v1.XmlPullParser.END_DOCUMENT) { }
    val builder = factory.newDocumentBuilder()
    builder.setEntityResolver { _, _ -> throw org.xml.sax.SAXException("External XML entities are not supported") }
    return builder.parse(ByteArrayInputStream(bytes))
}

fun Document.xmlBytes(): ByteArray = ByteArrayOutputStream().use { output ->
    TransformerFactory.newInstance().newTransformer().transform(DOMSource(this), StreamResult(output))
    output.toByteArray()
}

fun Node.elements(): List<Element> = (0 until childNodes.length)
    .mapNotNull { childNodes.item(it) as? Element }

fun Element.childrenNamed(namespace: String, name: String): List<Element> =
    elements().filter { it.namespaceURI == namespace && it.localName == name }
