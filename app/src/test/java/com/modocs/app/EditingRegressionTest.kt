package com.modocs.app

import com.modocs.feature.docx.*
import com.modocs.feature.xlsx.*
import org.junit.Assert.*
import org.junit.Test

class EditingRegressionTest {
    @Test fun textReplacementKeepsSourceIdentityAndUnchangedFormatting() {
        val runs = listOf(DocxRun("Hello ", RunProperties(bold = true), sourceRun = 0),
            DocxRun("world", RunProperties(italic = true), sourceRun = 1))
        val result = replaceParagraphText(runs, "Hello brave world")
        assertEquals("Hello brave world", result.joinToString("") { it.text })
        assertEquals(runs.last(), result.last())
        assertEquals(0, result[1].sourceRun)
        assertTrue(result[1].properties.bold)
        assertEquals(emptyList<DocxRun>(), replaceParagraphText(runs, ""))
    }
    @Test fun numberingContinuesAcrossPlainParagraphsAndUsesStartOverrides() {
        val numbering = mapOf("a" to NumberingDefinition("0", mapOf(0 to NumberingLevel(0, NumberFormat.DECIMAL, "%1)", start = 3))))
        val body = listOf(DocxParagraph(listOf(DocxRun("one")), listInfo = ListInfo("a")),
            DocxParagraph(listOf(DocxRun("plain"))), DocxParagraph(listOf(DocxRun("two")), listInfo = ListInfo("a")))
        val resolved = resolveListLabels(body, numbering)
        assertEquals("3)  ", (resolved[0] as DocxParagraph).listLabel)
        assertEquals("4)  ", (resolved[2] as DocxParagraph).listLabel)
        assertEquals(720, (resolved[2] as DocxParagraph).properties.indentLeftTwips)
    }
    @Test fun addressBoundsAndRoundTrip() {
        assertEquals(1048575 to 16383, parseCellReference("XFD1048576"))
        assertEquals("XFD1048576", cellReference(1048575, 16383))
        assertNull(parseCellReference("XFE1"))
        assertNull(parseCellReference("A0"))
        assertNull(parseCellReference("A1048577"))
        assertEquals(29 to 9, parseCellReference(" j30 "))
    }
}
