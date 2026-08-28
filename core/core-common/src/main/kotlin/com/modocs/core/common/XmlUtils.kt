package com.modocs.core.common

import android.util.Log
import org.xmlpull.v1.XmlPullParser
import org.xmlpull.v1.XmlPullParserFactory
import java.io.ByteArrayInputStream

/**
 * Run an optional OOXML parse step, degrading to [fallback] if it throws.
 *
 * Real-world OOXML comes from dozens of producers and a surprising amount of it
 * is subtly out of spec. Losing one part — styles, numbering, a single slide —
 * is nearly always better than refusing to open the document at all. Reserve
 * this for parts the document can survive without; let genuinely required parts
 * propagate so the user gets a real error instead of a blank screen.
 */
inline fun <T> tolerateParse(tag: String, what: String, fallback: T, block: () -> T): T = try {
    block()
} catch (e: Exception) {
    Log.w(tag, "Skipping unreadable $what: ${e.message}")
    fallback
}

/**
 * Create a namespace-aware pull parser over an OOXML part.
 *
 * The encoding is deliberately left for the parser to detect instead of being
 * pinned to UTF-8. Plenty of producers — ClosedXML, the .NET OpenXML SDK,
 * python-docx/openpyxl, Excel itself for some parts — write parts with a UTF-8
 * byte-order mark. Passing an explicit encoding makes the BOM part of the
 * character stream, so the leading `<?xml ...?>` is read as a processing
 * instruction preceded by content and the whole parse dies with
 * "PI must not start with xml". Auto-detection consumes the BOM and also copes
 * with UTF-16 parts, whose BOM is left in place below for exactly that reason.
 */
fun createXmlParser(bytes: ByteArray): XmlPullParser {
    val factory = XmlPullParserFactory.newInstance()
    factory.isNamespaceAware = true
    val parser = factory.newPullParser()
    parser.setInput(ByteArrayInputStream(stripUtf8Bom(bytes)), null)
    return parser
}

/**
 * Drop a leading UTF-8 BOM, if present. Auto-detection already handles this on
 * the parsers we ship against; stripping it keeps us correct on any
 * XmlPullParser implementation that does not. UTF-16 BOMs are left alone —
 * they are how the parser detects UTF-16 in the first place.
 */
private fun stripUtf8Bom(bytes: ByteArray): ByteArray =
    if (bytes.size >= 3 &&
        bytes[0] == 0xEF.toByte() &&
        bytes[1] == 0xBB.toByte() &&
        bytes[2] == 0xBF.toByte()
    ) {
        bytes.copyOfRange(3, bytes.size)
    } else {
        bytes
    }
