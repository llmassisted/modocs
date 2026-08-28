package com.modocs.feature.pdf

import android.content.Context
import android.net.Uri
import android.util.Log
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.withContext
import java.io.BufferedInputStream
import java.io.ByteArrayOutputStream

/**
 * Extracts text content and positions from PDF files for search + highlight.
 *
 * Uses a lightweight parser that reads PDF stream operators to extract text
 * and approximate glyph positions. Works on API 26+ without external libraries.
 *
 * Supports:
 * - Literal strings `(text)`
 * - Hex strings `<hex>`
 * - ToUnicode CMap decoding for embedded/CID fonts
 * - Indirect object content stream resolution
 * - Zlib-compressed streams
 */
object PdfTextExtractor {

    private const val TAG = "PdfTextExtractor"

    /**
     * Largest PDF we will pull into memory for text extraction. See
     * [readAtMost] for why this is far below the OOXML container cap.
     */
    private const val MAX_SOURCE_BYTES = 32L * 1024 * 1024 // 32 MB

    /** Output cap for a single Flate-compressed content stream. */
    private const val MAX_INFLATED_STREAM_BYTES = 64L * 1024 * 1024 // 64 MB

    /**
     * A positioned piece of text on a page.
     * Coordinates are in normalized form (0..1 relative to page size).
     */
    data class TextSegment(
        val text: String,
        val x: Float,      // normalized x (0 = left, 1 = right)
        val y: Float,      // normalized y (0 = top, 1 = bottom)
        val width: Float,  // normalized width
        val height: Float, // normalized height
    )

    data class PageText(
        val pageIndex: Int,
        val text: String,
        val segments: List<TextSegment>,
    )

    /**
     * Extract text from all pages of a PDF.
     */
    suspend fun extractText(context: Context, uri: Uri): List<PageText> =
        withContext(Dispatchers.IO) {
            try {
                val stream = context.contentResolver.openInputStream(uri)
                    ?: return@withContext emptyList()
                val bytes = stream.use { BufferedInputStream(it).readAtMost(MAX_SOURCE_BYTES) }

                if (bytes == null) {
                    // Too large to parse safely. Text search is unavailable for
                    // this document, but viewing and rendering are unaffected —
                    // far better than an OOM on the way in.
                    Log.i(TAG, "Skipping text extraction: PDF exceeds ${MAX_SOURCE_BYTES / (1024 * 1024)} MB")
                    return@withContext emptyList()
                }

                extractTextFromBytes(bytes)
            } catch (_: Exception) {
                emptyList()
            }
        }

    /**
     * Read the whole stream, or return null if it exceeds [limit].
     *
     * The extractor holds the file as a ByteArray *and* as an ISO-8859-1
     * String. PDF bodies are full of bytes above 0x7F, so that String is two
     * bytes per char and peak usage is roughly three times the file size —
     * which is why the limit here is well below the OOXML one.
     */
    private fun java.io.InputStream.readAtMost(limit: Long): ByteArray? {
        val out = ByteArrayOutputStream()
        val buffer = ByteArray(64 * 1024)
        var total = 0L
        while (true) {
            val read = read(buffer)
            if (read == -1) break
            total += read
            if (total > limit) return null
            out.write(buffer, 0, read)
        }
        return out.toByteArray()
    }

    private fun extractTextFromBytes(bytes: ByteArray): List<PageText> {
        val content = String(bytes, Charsets.ISO_8859_1)
        val pages = mutableListOf<PageText>()

        // Build object table for indirect reference resolution
        val objectTable = buildObjectTable(content)

        val pagePattern = Regex("""/Type\s*/Page\b""")
        val pageMatches = pagePattern.findAll(content).toList()

        val streamPattern = Regex("""stream\r?\n([\s\S]*?)endstream""")
        val allStreams = streamPattern.findAll(content).toList()

        val mediaBoxPattern = Regex("""/MediaBox\s*\[\s*([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)\s*]""")

        if (pageMatches.isEmpty()) {
            val allSegments = mutableListOf<TextSegment>()
            val allText = StringBuilder()
            for (match in allStreams) {
                val streamContent = match.groupValues[1]
                val (text, segments) = extractTextFromStream(streamContent, 612f, 792f, emptyMap())
                if (text.isNotBlank()) {
                    allText.append(text).append(" ")
                    allSegments.addAll(segments)
                }
            }
            if (allText.isNotBlank()) {
                pages.add(PageText(0, allText.toString().trim(), allSegments))
            }
            return pages
        }

        for (i in pageMatches.indices) {
            val pageStart = pageMatches[i].range.first
            val pageEnd = if (i + 1 < pageMatches.size) pageMatches[i + 1].range.first else content.length
            val pageSection = content.substring(pageStart, pageEnd)

            val mediaBoxMatch = mediaBoxPattern.find(pageSection)
            val pageWidth = mediaBoxMatch?.groupValues?.get(3)?.toFloatOrNull() ?: 612f
            val pageHeight = mediaBoxMatch?.groupValues?.get(4)?.toFloatOrNull() ?: 792f

            // Parse ToUnicode CMaps for fonts on this page
            val toUnicodeMaps = parseToUnicodeMaps(pageSection, content, objectTable)

            val pageText = StringBuilder()
            val pageSegments = mutableListOf<TextSegment>()

            // Find content streams — both inline and via indirect references
            val contentStreams = findContentStreams(pageSection, content, objectTable, allStreams, pageStart, pageEnd)

            for (streamContent in contentStreams) {
                val (text, segments) = extractTextFromStream(streamContent, pageWidth, pageHeight, toUnicodeMaps)
                if (text.isNotBlank()) {
                    pageText.append(text).append(" ")
                    pageSegments.addAll(segments)
                }
            }

            pages.add(PageText(i, pageText.toString().trim(), pageSegments))
        }

        return pages
    }

    // ---------------------------------------------------------------
    // Object table for indirect reference resolution
    // ---------------------------------------------------------------

    /**
     * Build a map from object number to (startOffset, endOffset) in the content.
     */
    private fun buildObjectTable(content: String): Map<Int, IntRange> {
        val table = mutableMapOf<Int, IntRange>()
        val objPattern = Regex("""(\d+)\s+\d+\s+obj\b""")

        for (match in objPattern.findAll(content)) {
            val objNum = match.groupValues[1].toIntOrNull() ?: continue
            val objStart = match.range.first
            val endObjIdx = content.indexOf("endobj", objStart)
            if (endObjIdx > objStart) {
                table[objNum] = objStart..endObjIdx + 5
            }
        }
        return table
    }

    /**
     * Get the content of an indirect object by number.
     */
    private fun getObjectContent(objNum: Int, content: String, objectTable: Map<Int, IntRange>): String? {
        val range = objectTable[objNum] ?: return null
        return content.substring(range.first, range.last.coerceAtMost(content.length))
    }

    // ---------------------------------------------------------------
    // Content stream discovery
    // ---------------------------------------------------------------

    /**
     * Find all content streams for a page. Handles:
     * - Inline streams (physically within the page object)
     * - /Contents references to indirect objects
     * - /Contents arrays of indirect references
     */
    private fun findContentStreams(
        pageSection: String,
        fullContent: String,
        objectTable: Map<Int, IntRange>,
        allStreams: List<MatchResult>,
        pageStart: Int,
        pageEnd: Int,
    ): List<String> {
        val streams = mutableListOf<String>()

        // Try to find /Contents reference
        val contentsRef = Regex("""/Contents\s+(\d+)\s+\d+\s+R""").find(pageSection)
        val contentsArray = Regex("""/Contents\s*\[([\s\d+R]+)]""").find(pageSection)

        if (contentsRef != null) {
            // Single indirect reference
            val objNum = contentsRef.groupValues[1].toIntOrNull()
            if (objNum != null) {
                val objContent = getObjectContent(objNum, fullContent, objectTable)
                if (objContent != null) {
                    val stream = extractStreamFromObject(objContent)
                    if (stream != null) streams.add(stream)
                }
            }
        } else if (contentsArray != null) {
            // Array of indirect references
            val refs = Regex("""(\d+)\s+\d+\s+R""").findAll(contentsArray.groupValues[1])
            for (ref in refs) {
                val objNum = ref.groupValues[1].toIntOrNull() ?: continue
                val objContent = getObjectContent(objNum, fullContent, objectTable)
                if (objContent != null) {
                    val stream = extractStreamFromObject(objContent)
                    if (stream != null) streams.add(stream)
                }
            }
        }

        // Also check inline streams (fallback)
        if (streams.isEmpty()) {
            for (match in allStreams) {
                if (match.range.first in pageStart until pageEnd) {
                    val streamContent = match.groupValues[1]
                    val decoded = tryDecompress(streamContent) ?: streamContent
                    streams.add(decoded)
                }
            }
        }

        return streams
    }

    /**
     * Extract and decode a stream from an object body.
     */
    private fun extractStreamFromObject(objContent: String): String? {
        val streamMatch = Regex("""stream\r?\n([\s\S]*?)endstream""").find(objContent) ?: return null
        val raw = streamMatch.groupValues[1]
        return tryDecompress(raw) ?: raw
    }

    // ---------------------------------------------------------------
    // ToUnicode CMap parsing
    // ---------------------------------------------------------------

    /**
     * Parse ToUnicode CMaps for all fonts referenced in the page section.
     * Returns a map from font name (e.g. "F1") to glyph code → Unicode string mapping.
     */
    private fun parseToUnicodeMaps(
        pageSection: String,
        fullContent: String,
        objectTable: Map<Int, IntRange>,
    ): Map<String, Map<Int, String>> {
        val maps = mutableMapOf<String, Map<Int, String>>()

        // Find font references in page resources
        // Pattern: /F1 <objNum> 0 R   or within /Font << /F1 <objNum> 0 R ... >>
        val fontRefs = Regex("""/(F\d+)\s+(\d+)\s+\d+\s+R""").findAll(pageSection)

        for (fontRef in fontRefs) {
            val fontName = fontRef.groupValues[1]
            val fontObjNum = fontRef.groupValues[2].toIntOrNull() ?: continue

            val fontObj = getObjectContent(fontObjNum, fullContent, objectTable) ?: continue

            // Look for /ToUnicode reference
            val toUnicodeRef = Regex("""/ToUnicode\s+(\d+)\s+\d+\s+R""").find(fontObj) ?: continue
            val cmapObjNum = toUnicodeRef.groupValues[1].toIntOrNull() ?: continue

            val cmapObj = getObjectContent(cmapObjNum, fullContent, objectTable) ?: continue
            val cmapStream = extractStreamFromObject(cmapObj) ?: continue

            val mapping = parseCMap(cmapStream)
            if (mapping.isNotEmpty()) {
                maps[fontName] = mapping
            }
        }

        return maps
    }

    /**
     * Parse a CMap stream to build glyph code → Unicode string mapping.
     * Handles beginbfchar/endbfchar and beginbfrange/endbfrange.
     */
    private fun parseCMap(cmap: String): Map<Int, String> {
        val mapping = mutableMapOf<Int, String>()

        // Parse bfchar entries: <srcCode> <dstUnicode>
        val bfcharSections = Regex("""beginbfchar([\s\S]*?)endbfchar""").findAll(cmap)
        for (section in bfcharSections) {
            val entries = Regex("""<([0-9a-fA-F]+)>\s*<([0-9a-fA-F]+)>""").findAll(section.groupValues[1])
            for (entry in entries) {
                val srcCode = entry.groupValues[1].toIntOrNull(16) ?: continue
                val dstUnicode = hexToUnicodeString(entry.groupValues[2])
                mapping[srcCode] = dstUnicode
            }
        }

        // Parse bfrange entries: <srcCodeLo> <srcCodeHi> <dstUnicodeStart>
        val bfrangeSections = Regex("""beginbfrange([\s\S]*?)endbfrange""").findAll(cmap)
        for (section in bfrangeSections) {
            val entries = Regex("""<([0-9a-fA-F]+)>\s*<([0-9a-fA-F]+)>\s*<([0-9a-fA-F]+)>""")
                .findAll(section.groupValues[1])
            for (entry in entries) {
                val lo = entry.groupValues[1].toIntOrNull(16) ?: continue
                val hi = entry.groupValues[2].toIntOrNull(16) ?: continue
                val dstStart = entry.groupValues[3].toIntOrNull(16) ?: continue

                for (code in lo..hi) {
                    val unicodePoint = dstStart + (code - lo)
                    mapping[code] = String(Character.toChars(unicodePoint))
                }
            }
        }

        return mapping
    }

    /**
     * Convert a hex string like "0041" to a Unicode string "A".
     */
    private fun hexToUnicodeString(hex: String): String {
        val sb = StringBuilder()
        var i = 0
        while (i + 3 < hex.length) {
            val codePoint = hex.substring(i, i + 4).toIntOrNull(16) ?: break
            sb.appendCodePoint(codePoint)
            i += 4
        }
        // Handle 2-digit codes
        if (i < hex.length && sb.isEmpty()) {
            val codePoint = hex.toIntOrNull(16) ?: return ""
            sb.appendCodePoint(codePoint)
        }
        return sb.toString()
    }

    // ---------------------------------------------------------------
    // Stream decompression
    // ---------------------------------------------------------------

    private fun tryDecompress(streamContent: String): String? {
        return try {
            val bytes = streamContent.toByteArray(Charsets.ISO_8859_1)
            if (bytes.size < 2 || (bytes[0].toInt() and 0xFF) != 0x78) return null

            val inflater = java.util.zip.Inflater()
            inflater.setInput(bytes)
            val output = ByteArrayOutputStream()
            val buffer = ByteArray(4096)
            var total = 0L
            try {
                while (!inflater.finished()) {
                    val count = inflater.inflate(buffer)
                    if (count == 0 && inflater.needsInput()) break
                    total += count
                    // A content stream that inflates past this is a decompression
                    // bomb, not a page of text. Give up on this stream rather
                    // than growing the buffer until the process dies.
                    if (total > MAX_INFLATED_STREAM_BYTES) return null
                    output.write(buffer, 0, count)
                }
            } finally {
                inflater.end()
            }
            String(output.toByteArray(), Charsets.ISO_8859_1)
        } catch (_: Exception) {
            null
        }
    }

    // ---------------------------------------------------------------
    // Text extraction from content streams
    // ---------------------------------------------------------------

    private fun extractTextFromStream(
        streamContent: String,
        pageWidth: Float,
        pageHeight: Float,
        toUnicodeMaps: Map<String, Map<Int, String>>,
    ): Pair<String, List<TextSegment>> {
        val decodedContent = tryDecompress(streamContent) ?: streamContent
        return extractTextOperators(decodedContent, pageWidth, pageHeight, toUnicodeMaps)
    }

    private fun extractTextOperators(
        content: String,
        pageWidth: Float,
        pageHeight: Float,
        toUnicodeMaps: Map<String, Map<Int, String>>,
    ): Pair<String, List<TextSegment>> {
        val fullText = StringBuilder()
        val segments = mutableListOf<TextSegment>()

        var pos = 0
        while (pos < content.length) {
            val btIndex = content.indexOf("BT", pos)
            if (btIndex == -1) break

            val etIndex = content.indexOf("ET", btIndex)
            if (etIndex == -1) break

            val textBlock = content.substring(btIndex, etIndex)
            extractFromTextBlock(textBlock, fullText, segments, pageWidth, pageHeight, toUnicodeMaps)

            pos = etIndex + 2
        }

        val text = fullText.toString().replace(Regex("\\s+"), " ").trim()
        return text to segments
    }

    private fun extractFromTextBlock(
        block: String,
        fullText: StringBuilder,
        segments: MutableList<TextSegment>,
        pageWidth: Float,
        pageHeight: Float,
        toUnicodeMaps: Map<String, Map<Int, String>>,
    ) {
        var textX = 0f
        var textY = 0f
        var fontSize = 12f
        var leading = 14f
        var currentFontName = ""
        var currentCMap: Map<Int, String>? = null

        val lines = block.split('\n')
        for (line in lines) {
            val trimmed = line.trim()
            when {
                // Tf: set font — "/<fontName> <size> Tf"
                trimmed.endsWith(" Tf") || trimmed.endsWith(" tf") -> {
                    val parts = trimmed.split(Regex("\\s+"))
                    if (parts.size >= 3) {
                        fontSize = parts[parts.size - 2].toFloatOrNull() ?: fontSize
                        val fontRef = parts[parts.size - 3].removePrefix("/")
                        if (fontRef != currentFontName) {
                            currentFontName = fontRef
                            currentCMap = toUnicodeMaps[fontRef]
                        }
                    }
                }
                // TL: set leading
                trimmed.endsWith(" TL") -> {
                    val parts = trimmed.split(Regex("\\s+"))
                    if (parts.size >= 2) {
                        leading = parts[0].toFloatOrNull()?.let { kotlin.math.abs(it) } ?: leading
                    }
                }
                // Tm: set text matrix
                trimmed.endsWith(" Tm") || trimmed.endsWith(" tm") -> {
                    val parts = trimmed.split(Regex("\\s+"))
                    if (parts.size >= 7) {
                        textX = parts[parts.size - 3].toFloatOrNull() ?: textX
                        textY = parts[parts.size - 2].toFloatOrNull() ?: textY
                        val matrixFontSize = parts[parts.size - 5].toFloatOrNull()
                        if (matrixFontSize != null && matrixFontSize > 0) {
                            fontSize = matrixFontSize
                        }
                    }
                }
                // Td/TD: move text position
                trimmed.endsWith(" Td") || trimmed.endsWith(" TD") -> {
                    val parts = trimmed.split(Regex("\\s+"))
                    if (parts.size >= 3) {
                        val tx = parts[parts.size - 3].toFloatOrNull() ?: 0f
                        val ty = parts[parts.size - 2].toFloatOrNull() ?: 0f
                        textX += tx
                        textY += ty
                        if (trimmed.endsWith(" TD") && ty != 0f) {
                            leading = kotlin.math.abs(ty)
                        }
                    }
                }
                // T*: move to next line
                trimmed == "T*" -> {
                    textY -= leading
                }
                // ' operator
                trimmed.endsWith(" '") -> {
                    textY -= leading
                    extractStringsFromLine(trimmed, fullText, segments, textX, textY, fontSize, pageWidth, pageHeight, currentCMap)
                }
                // Tj/TJ operators
                trimmed.endsWith(" Tj") || trimmed.endsWith(" TJ") -> {
                    extractStringsFromLine(trimmed, fullText, segments, textX, textY, fontSize, pageWidth, pageHeight, currentCMap)
                }
                // Lines containing string literals with text operators
                (trimmed.contains("(") || trimmed.contains("<")) &&
                    (trimmed.contains("Tj") || trimmed.contains("TJ") || trimmed.contains("'")) -> {
                    extractStringsFromLine(trimmed, fullText, segments, textX, textY, fontSize, pageWidth, pageHeight, currentCMap)
                }
            }
        }
    }

    private fun extractStringsFromLine(
        line: String,
        fullText: StringBuilder,
        segments: MutableList<TextSegment>,
        textX: Float,
        textY: Float,
        fontSize: Float,
        pageWidth: Float,
        pageHeight: Float,
        cmap: Map<Int, String>?,
    ) {
        var i = 0
        val lineText = StringBuilder()

        while (i < line.length) {
            when {
                line[i] == '(' -> {
                    // Literal string
                    val result = readPdfLiteralString(line, i, cmap)
                    if (result != null) {
                        lineText.append(result.first)
                        i = result.second
                    } else {
                        i++
                    }
                }
                line[i] == '<' && i + 1 < line.length && line[i + 1] != '<' -> {
                    // Hex string
                    val result = readPdfHexString(line, i, cmap)
                    if (result != null) {
                        lineText.append(result.first)
                        i = result.second
                    } else {
                        i++
                    }
                }
                line[i] == '[' -> {
                    // TJ array — contains mix of strings and numbers
                    i++
                    while (i < line.length && line[i] != ']') {
                        when {
                            line[i] == '(' -> {
                                val result = readPdfLiteralString(line, i, cmap)
                                if (result != null) {
                                    lineText.append(result.first)
                                    i = result.second
                                } else {
                                    i++
                                }
                            }
                            line[i] == '<' && i + 1 < line.length && line[i + 1] != '<' -> {
                                val result = readPdfHexString(line, i, cmap)
                                if (result != null) {
                                    lineText.append(result.first)
                                    i = result.second
                                } else {
                                    i++
                                }
                            }
                            line[i].isDigit() || line[i] == '-' -> {
                                val numStart = i
                                while (i < line.length && (line[i].isDigit() || line[i] == '-' || line[i] == '.')) i++
                                val num = line.substring(numStart, i).toDoubleOrNull() ?: 0.0
                                if (num < -100) lineText.append(' ')
                            }
                            else -> i++
                        }
                    }
                    if (i < line.length) i++ // skip ']'
                }
                else -> i++
            }
        }

        val text = lineText.toString()
        if (text.isNotBlank()) {
            fullText.append(text)

            val approxWidth = text.length * fontSize * 0.5f
            val normX = (textX / pageWidth).coerceIn(0f, 1f)
            val normY = (1f - (textY + fontSize) / pageHeight).coerceIn(0f, 1f)
            val normW = (approxWidth / pageWidth).coerceIn(0f, 1f - normX)
            val normH = (fontSize * 1.3f / pageHeight).coerceIn(0f, 0.1f)

            segments.add(TextSegment(text = text, x = normX, y = normY, width = normW, height = normH))
        }
    }

    // ---------------------------------------------------------------
    // String reading
    // ---------------------------------------------------------------

    /**
     * Read a PDF literal string starting at '('.
     * Applies CMap if available.
     */
    private fun readPdfLiteralString(text: String, startPos: Int, cmap: Map<Int, String>?): Pair<String, Int>? {
        if (startPos >= text.length || text[startPos] != '(') return null

        val result = StringBuilder()
        var i = startPos + 1
        var depth = 1

        while (i < text.length && depth > 0) {
            when {
                text[i] == '\\' && i + 1 < text.length -> {
                    i++
                    when (text[i]) {
                        'n' -> result.append('\n')
                        'r' -> result.append('\r')
                        't' -> result.append('\t')
                        '(' -> result.append('(')
                        ')' -> result.append(')')
                        '\\' -> result.append('\\')
                        in '0'..'7' -> {
                            var octal = ""
                            while (i < text.length && text[i] in '0'..'7' && octal.length < 3) {
                                octal += text[i]
                                i++
                            }
                            val charCode = octal.toIntOrNull(8) ?: 0
                            if (cmap != null) {
                                val mapped = cmap[charCode]
                                if (mapped != null) {
                                    result.append(mapped)
                                } else if (charCode in 32..126) {
                                    result.append(charCode.toChar())
                                }
                            } else if (charCode in 32..126) {
                                result.append(charCode.toChar())
                            }
                            continue
                        }
                        else -> result.append(text[i])
                    }
                }
                text[i] == '(' -> {
                    depth++
                    result.append('(')
                }
                text[i] == ')' -> {
                    depth--
                    if (depth > 0) result.append(')')
                }
                else -> {
                    val c = text[i]
                    if (cmap != null) {
                        val mapped = cmap[c.code]
                        if (mapped != null) {
                            result.append(mapped)
                        } else if (c.code in 32..126 || c == '\n' || c == '\r' || c == '\t') {
                            result.append(c)
                        }
                    } else if (c.code in 32..126 || c == '\n' || c == '\r' || c == '\t') {
                        result.append(c)
                    }
                }
            }
            i++
        }

        return result.toString() to i
    }

    /**
     * Read a PDF hex string starting at '<'.
     * Decodes hex pairs and applies CMap if available.
     */
    private fun readPdfHexString(text: String, startPos: Int, cmap: Map<Int, String>?): Pair<String, Int>? {
        if (startPos >= text.length || text[startPos] != '<') return null

        var i = startPos + 1
        val hexChars = StringBuilder()

        while (i < text.length && text[i] != '>') {
            val c = text[i]
            if (c in '0'..'9' || c in 'a'..'f' || c in 'A'..'F') {
                hexChars.append(c)
            }
            i++
        }
        if (i < text.length) i++ // skip '>'

        val hex = hexChars.toString()
        if (hex.isEmpty()) return "" to i

        val result = StringBuilder()

        if (cmap != null && cmap.isNotEmpty()) {
            // Try 2-byte (4 hex char) codes first, then 1-byte (2 hex char)
            var h = 0
            while (h + 3 < hex.length) {
                val code2 = hex.substring(h, h + 4).toIntOrNull(16)
                if (code2 != null && cmap.containsKey(code2)) {
                    result.append(cmap[code2])
                    h += 4
                    continue
                }
                // Try 1-byte
                val code1 = hex.substring(h, h + 2).toIntOrNull(16)
                if (code1 != null) {
                    val mapped = cmap[code1]
                    if (mapped != null) {
                        result.append(mapped)
                    } else if (code1 in 32..126) {
                        result.append(code1.toChar())
                    }
                }
                h += 2
            }
            // Handle remaining 2 hex chars
            if (h + 1 < hex.length) {
                val code = hex.substring(h, h + 2).toIntOrNull(16)
                if (code != null) {
                    val mapped = cmap[code]
                    if (mapped != null) result.append(mapped)
                    else if (code in 32..126) result.append(code.toChar())
                }
            }
        } else {
            // No CMap — decode as simple byte values
            var h = 0
            while (h + 1 < hex.length) {
                val code = hex.substring(h, h + 2).toIntOrNull(16)
                if (code != null && code in 32..126) {
                    result.append(code.toChar())
                }
                h += 2
            }
        }

        return result.toString() to i
    }
}
