package com.modocs.app

import com.modocs.feature.xlsx.*
import org.junit.Assert.*
import org.junit.Test
import java.nio.charset.Charset

class CsvParsingTest {
    private fun parse(bytes: ByteArray) = CsvParser.parse(bytes.inputStream())
    private fun parse(text: String) = parse(text.toByteArray())
    private fun bytes(vararg values: Int) = ByteArray(values.size) { values[it].toByte() }
    private fun XlsxDocument.grid(): List<List<String>> {
        val sheet = sheets.single()
        val rowCount = sheet.rows.maxOfOrNull { it.rowIndex + 1 } ?: 0
        return List(rowCount) { row -> List(sheet.columnCount) { column -> sheet.cellAt(row, column)?.value.orEmpty() } }
    }

    @Test fun quotedFieldsKeepDelimitersQuotesAndLineBreaks() {
        val document = parse("name,note,qty\r\n\"Smith, J\",\"said \"\"hi\"\"\r\nthen left\",3\r\nplain,,4\r\n")
        assertEquals(listOf(
            listOf("name", "note", "qty"),
            listOf("Smith, J", "said \"hi\"\r\nthen left", "3"),
            listOf("plain", "", "4"),
        ), document.grid())
        assertEquals(listOf(listOf("a", "b\nc")), parse("a,\"b\nc").grid())
        assertEquals(listOf(listOf("5\" pipe", "x")), parse("5\" pipe,x").grid())
        assertEquals(listOf(listOf("a"), listOf("b")), parse("a\rb").grid())
        assertTrue(document.warnings.isEmpty())
    }

    @Test fun byteOrderMarksAndLegacyEncodingsDecode() {
        assertEquals(listOf(listOf("Größe", "€"), listOf("1", "2")),
            parse(bytes(0xEF, 0xBB, 0xBF) + "Größe,€\n1,2".toByteArray()).grid())
        assertEquals(listOf(listOf("a", "b"), listOf("1", "2")),
            parse(bytes(0xFF, 0xFE) + "a\tb\r\n1\t2".toByteArray(Charsets.UTF_16LE)).grid())
        assertEquals(listOf(listOf("café", "crème"), listOf("1", "2")),
            parse("café;crème\n1;2".toByteArray(Charset.forName("windows-1252"))).grid())
        assertThrows(IllegalArgumentException::class.java) { parse(bytes(0x50, 0x4B, 0x03, 0x04, 0x00, 0x00)) }
    }

    @Test fun delimiterIsSniffedAndSepHintHonoured() {
        assertEquals(listOf(listOf("Item", "Price"), listOf("Apple", "1,5"), listOf("Pear", "2,25")),
            parse("Item;Price\nApple;1,5\nPear;2,25\n").grid())
        assertEquals(listOf(listOf("a", "b", "c"), listOf("1", "2", "3")), parse("a|b|c\n1|2|3").grid())
        assertEquals(listOf(listOf("a", "b"), listOf("x;y", "z")), parse("a,b\n\"x;y\",z").grid())
        assertEquals(listOf(listOf("a,b", "c"), listOf("1", "2")), parse("sep=;\na,b;c\n1;2").grid())
    }

    @Test fun typesHeaderRowAndBlankLines() {
        val sheet = parse("id,zip,amount\n1,02134,-3.5e2\n\n2,10001,7\n\n").sheets.single()
        assertEquals(4, sheet.rows.maxOf { it.rowIndex + 1 })
        assertNull(sheet.rows.find { it.rowIndex == 2 })
        assertEquals(CellType.NUMBER, sheet.cellAt(1, 0)?.type)
        assertEquals(CellType.STRING, sheet.cellAt(1, 1)?.type)
        assertEquals(CellType.NUMBER, sheet.cellAt(1, 2)?.type)
        assertEquals(1, sheet.frozenRows)
        assertEquals(0, parse("1,2\n3,4").sheets.single().frozenRows)
        assertEquals(0, parse("a,b").sheets.single().frozenRows)
        assertEquals(0, parse("a,\n1,2").sheets.single().frozenRows)
        assertTrue(parse("").sheets.single().rows.isEmpty())
    }

    @Test fun columnWidthsFollowContent() {
        val widths = parse("id,description\n1,${"x".repeat(20)}\n").sheets.single().columnWidths
        assertEquals(6f, widths[0])
        assertEquals(20f, widths[1])
        assertEquals(30f, parse("x".repeat(500)).sheets.single().columnWidths[0])
    }

    @Test fun oversizedFilesAreCappedWithAWarning() {
        val capped = parse("1\n".repeat(CsvParser.MAX_CELLS + 1))
        assertEquals(CsvParser.MAX_CELLS, capped.sheets.single().rows.size)
        assertEquals(1, capped.warnings.size)
        val longCell = parse("\"${"x".repeat(40_000)}\",y")
        assertEquals(32_767, longCell.sheets.single().cellAt(0, 0)?.value?.length)
        assertEquals("y", longCell.sheets.single().cellAt(0, 1)?.value)
        assertEquals(1, longCell.warnings.size)
    }
}
