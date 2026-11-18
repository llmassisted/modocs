package com.modocs.feature.xlsx

import android.content.Context
import android.net.Uri
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.withContext
import java.io.ByteArrayInputStream
import java.io.InputStream
import java.io.InputStreamReader
import java.io.Reader
import java.io.SequenceInputStream
import java.io.StringReader
import java.nio.ByteBuffer
import java.nio.CharBuffer
import java.nio.charset.Charset
import java.nio.charset.CodingErrorAction

/**
 * Parses CSV text into a one-sheet [XlsxDocument] so it opens in the spreadsheet viewer.
 *
 * Quoting follows RFC 4180 (quoted fields may hold delimiters, line breaks and `""`
 * escapes) but is read leniently, as spreadsheet apps do: a quote inside an unquoted
 * field is kept as text, and an unterminated quote runs to the end of the file rather
 * than failing. The delimiter is sniffed because "CSV" exported in many locales is
 * semicolon-separated, and Excel's `sep=` hint line is honoured.
 *
 * The file is streamed instead of read whole, and rows past [MAX_CELLS] are dropped
 * with a warning: every cell is an object, so a 100 MB file would otherwise exhaust
 * the heap while the grid is built.
 */
object CsvParser {

    /** Keeps the grid comfortably inside a small device's heap. */
    const val MAX_CELLS = 500_000

    // Excel's sheet limits; the grid clamps to the same bounds.
    private const val MAX_ROWS = 1_048_576
    private const val MAX_COLUMNS = 16_384
    private const val MAX_CELL_CHARS = 32_767

    private const val SAMPLE_BYTES = 64 * 1024
    private const val SNIFF_RECORDS = 20
    private val DELIMITERS = charArrayOf(',', ';', '\t', '|')
    private val SeparatorHint = Regex("""^sep=([^"\r\n])(\r?\n|$)""", RegexOption.IGNORE_CASE)

    suspend fun parse(context: Context, uri: Uri): XlsxDocument = withContext(Dispatchers.IO) {
        val inputStream = context.contentResolver.openInputStream(uri)
            ?: throw IllegalStateException("Cannot open file")
        inputStream.use { parse(it) }
    }

    fun parse(inputStream: InputStream): XlsxDocument {
        // The opening bytes decide the encoding and delimiter; they are then
        // replayed ahead of the rest of the stream.
        val sample = ByteArray(SAMPLE_BYTES)
        var sampleSize = 0
        while (sampleSize < SAMPLE_BYTES) {
            val read = inputStream.read(sample, sampleSize, SAMPLE_BYTES - sampleSize)
            if (read < 0) break
            sampleSize += read
        }
        val complete = sampleSize < SAMPLE_BYTES
        val (charset, bomSize) = detectEncoding(sample, sampleSize, complete)
        val text = String(sample, bomSize, sampleSize - bomSize, charset)

        val hint = SeparatorHint.find(text)
        val delimiter = hint?.groupValues?.get(1)?.single() ?: detectDelimiter(text, complete)
        val body = SequenceInputStream(ByteArrayInputStream(sample, bomSize, sampleSize - bomSize), inputStream)
        val records = RecordReader(InputStreamReader(body, charset), delimiter)
        // Excel hides its hint line too.
        if (hint != null) records.next()

        val rows = mutableListOf<XlsxRow>()
        val longest = IntArray(MAX_COLUMNS)
        var columnCount = 0
        var cellCount = 0
        var rowIndex = 0
        var truncated = false
        while (true) {
            val fields = records.next() ?: break
            val filled = fields.indices.filter { fields[it].isNotEmpty() }
            if (rowIndex == MAX_ROWS || cellCount + filled.size > MAX_CELLS) {
                truncated = true
                break
            }
            // Blank lines keep their place in the grid, as in a spreadsheet app.
            if (filled.isNotEmpty()) {
                rows.add(XlsxRow(rowIndex, filled.mapTo(ArrayList(filled.size)) { column ->
                    val value = fields[column]
                    longest[column] = maxOf(longest[column], value.length)
                    XlsxCell(columnIndex = column, value = value, type = inferCellType(value))
                }))
                cellCount += filled.size
                columnCount = maxOf(columnCount, filled.last() + 1)
            }
            rowIndex++
        }

        val warnings = buildList {
            if (truncated) add("Showing the first ${"%,d".format(rowIndex)} rows; the rest of this file is too large to display")
            if (records.clipped) add("Some cells or rows were cut short to fit within spreadsheet limits")
        }
        val sheet = XlsxSheet(
            name = "Sheet1",
            rows = rows,
            columnWidths = (0 until columnCount).associateWith { longest[it].coerceIn(6, 30).toFloat() },
            frozenRows = if (hasHeaderRow(rows, columnCount)) 1 else 0,
        )
        // A default style lets numbers right-align as they do in spreadsheet apps.
        return XlsxDocument(sheets = listOf(sheet), styles = listOf(XlsxCellStyle()), warnings = warnings)
    }

    private fun detectEncoding(sample: ByteArray, size: Int, complete: Boolean): Pair<Charset, Int> {
        fun startsWith(vararg bom: Int) = size >= bom.size && bom.indices.all { sample[it].toInt() and 0xFF == bom[it] }
        if (startsWith(0xEF, 0xBB, 0xBF)) return Charsets.UTF_8 to 3
        if (startsWith(0xFF, 0xFE)) return Charsets.UTF_16LE to 2
        if (startsWith(0xFE, 0xFF)) return Charsets.UTF_16BE to 2
        // Without a UTF-16 byte order mark, NUL bytes mean this is not text at all,
        // such as a workbook renamed to .csv; showing it as cells would only be noise.
        require((0 until size).none { sample[it] == 0.toByte() }) { "This file does not contain CSV text" }
        // Excel's legacy "CSV (Comma delimited)" export writes the Windows code page.
        return (if (isUtf8(sample, size, complete)) Charsets.UTF_8 else Charset.forName("windows-1252")) to 0
    }

    private fun isUtf8(sample: ByteArray, size: Int, complete: Boolean): Boolean {
        val decoder = Charsets.UTF_8.newDecoder()
            .onMalformedInput(CodingErrorAction.REPORT)
            .onUnmappableCharacter(CodingErrorAction.REPORT)
        // A partial sample may end mid-character; only a whole file must end cleanly.
        return !decoder.decode(ByteBuffer.wrap(sample, 0, size), CharBuffer.allocate(size), complete).isError
    }

    /**
     * Pick the delimiter that splits the opening records into the same number of
     * fields most often, preferring more fields. The header usually decides it:
     * `Item;Price` is one comma field but two semicolon fields, even when the data
     * below uses decimal commas.
     */
    private fun detectDelimiter(sample: String, complete: Boolean): Char {
        var best = ','
        var bestAgreement = 0
        var bestFields = 1
        for (candidate in DELIMITERS) {
            val reader = RecordReader(StringReader(sample), candidate)
            val counts = generateSequence { reader.next() }
                .filter { it.size > 1 || it[0].isNotEmpty() }
                .take(SNIFF_RECORDS + 1)
                .map { it.size }
                .toList()
                // A cut-off sample may also cut off its last record.
                .let { if (complete) it else it.dropLast(1) }
            val fields = counts.firstOrNull() ?: continue
            if (fields < 2) continue
            val agreement = counts.count { it == fields }
            if (agreement > bestAgreement || (agreement == bestAgreement && fields > bestFields)) {
                best = candidate
                bestAgreement = agreement
                bestFields = fields
            }
        }
        return best
    }

    /** A first row that names every column, with no numbers in it, is almost always a header. */
    private fun hasHeaderRow(rows: List<XlsxRow>, columnCount: Int): Boolean {
        val first = rows.firstOrNull() ?: return false
        return rows.size > 1 && first.rowIndex == 0 && first.cells.size == columnCount &&
            first.cells.all { it.type == CellType.STRING && it.value.isNotBlank() }
    }

    /** Splits CSV text into records, one call to [next] per record. */
    private class RecordReader(private val reader: Reader, private val delimiter: Char) {
        private val buffer = CharArray(8192)
        private var position = 0
        private var limit = 0

        /** Set once a field was cut at [MAX_CELL_CHARS] or a record at [MAX_COLUMNS]. */
        var clipped = false
            private set

        private fun read(): Int {
            if (position == limit) {
                limit = reader.read(buffer).coerceAtLeast(0)
                position = 0
                if (limit == 0) return -1
            }
            return buffer[position++].code
        }

        /** The next record's fields, or null once the input is exhausted. */
        fun next(): List<String>? {
            var c = read()
            if (c < 0) return null
            val fields = ArrayList<String>()
            val field = StringBuilder()
            var quoted = false
            fun append(char: Int) {
                if (field.length < MAX_CELL_CHARS) field.append(char.toChar()) else clipped = true
            }
            fun endField() {
                if (fields.size < MAX_COLUMNS) fields.add(field.toString()) else clipped = true
                field.setLength(0)
            }
            while (true) {
                when {
                    c < 0 -> {
                        endField()
                        return fields
                    }
                    quoted -> if (c != '"'.code) {
                        append(c)
                    } else {
                        c = read()
                        if (c == '"'.code) {
                            append(c)
                        } else {
                            quoted = false
                            continue
                        }
                    }
                    c == '"'.code && field.isEmpty() -> quoted = true
                    c == delimiter.code -> endField()
                    c == '\n'.code || c == '\r'.code -> {
                        // Treat CRLF as one break, and push back whatever followed a lone CR.
                        if (c == '\r'.code && read() != '\n'.code && limit > 0) position--
                        endField()
                        return fields
                    }
                    else -> append(c)
                }
                c = read()
            }
        }
    }
}
