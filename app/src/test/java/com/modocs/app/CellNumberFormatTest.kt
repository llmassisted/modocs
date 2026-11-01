package com.modocs.app

import com.modocs.feature.xlsx.CellNumberFormat
import org.junit.After
import org.junit.Assert.*
import org.junit.Before
import org.junit.Test
import java.util.Locale

class CellNumberFormatTest {
    private val original = Locale.getDefault()
    @Before fun pinLocale() = Locale.setDefault(Locale.US)
    @After fun restoreLocale() = Locale.setDefault(original)

    private fun f(raw: String, code: String?, id: Int = 164, date1904: Boolean = false) = CellNumberFormat.format(raw, id, code, date1904)

    @Test fun builtInAndCustomNumberFormats() {
        assertEquals("25%", f("0.25", null, 9))
        assertEquals("12.50%", f("0.125", null, 10))
        assertEquals("1,234,568", f("1234567.89", null, 3))
        assertEquals("1,234.50", f("1234.5", null, 4))
        assertEquals("$1,234.50 ", f("1234.5", null, 7)) // _) pads like Excel, to line up with (negatives)
        assertEquals("($1,234.50)", f("-1234.5", null, 7))
        assertEquals("1.23E+04", f("12345", null, 11))
        assertEquals("2.68", f("2.675", "0.00"))
        assertEquals("-3", f("-3", "0"))
        assertEquals("0", f("-0.001", "0"))
        assertEquals("€ 1,234.50", f("1234.5", "[\$€-407] #,##0.00"))
        assertEquals("1,235K", f("1234567", "#,##0,\"K\""))
        assertEquals("02134-0042", f("21340042", "00000-0000"))
        assertEquals("(555) 123-4567", f("5551234567", "[<=9999999]###-####;(###) ###-####"))
        assertEquals("zero", f("0", "0.00;(0.00);\"zero\""))
        assertEquals("(4.20)", f("-4.2", "0.00;[Red](0.00)"))
        assertEquals(".5", f("0.5", "#.##"))
        assertEquals("12.5", f("12.5", "General", 0))
        assertEquals("1.5E+20", f("1.5e20", null, 0))
        assertEquals("hello", f("hello", null, 4))
    }

    @Test fun separatorsAndMonthNamesFollowTheDeviceLocale() {
        Locale.setDefault(Locale.GERMANY)
        assertEquals("1.234,50", f("1234.5", null, 4))
        assertEquals("12,5", f("12.5", null, 0))
        assertEquals("01.11.2023", f("45231", null, 14))
        assertEquals("1. November 2023", f("45231", "d. mmmm yyyy"))
    }

    @Test fun datesAndTimes() {
        assertTrue(CellNumberFormat.isDateFormat(14, null))
        assertTrue(CellNumberFormat.isDateFormat(164, "dd/mm/yyyy"))
        assertTrue(CellNumberFormat.isDateFormat(164, "h:mm AM/PM"))
        assertTrue(CellNumberFormat.isDateFormat(164, "[h]:mm"))
        assertFalse(CellNumberFormat.isDateFormat(164, "#,##0.00 \"days\""))
        assertFalse(CellNumberFormat.isDateFormat(164, "[Red]0.00"))
        assertFalse(CellNumberFormat.isDateFormat(0, null))

        assertEquals("11/1/2023", f("45231", null, 14))
        assertEquals("11/1/2023 13:30", f("45231.5625", null, 22))
        assertEquals("1-Nov-23", f("45231", null, 15))
        assertEquals("12:00 PM", f("0.5", null, 18))
        assertEquals("6:30:15 AM", f("0.271006944", null, 19))
        assertEquals("27:00:00", f("1.125", null, 46))
        assertEquals("Wednesday, November 01, 2023", f("45231", "dddd, mmmm dd, yyyy"))
        assertEquals("2023-11-01 09:05", f("45231.37847", "yyyy-mm-dd hh:mm"))
        assertEquals("05:07.3", f("0.003557", "mm:ss.0"))
        assertEquals("2/28/1900", f("59", "m/d/yyyy"))
        assertEquals("3/1/1900", f("61", "m/d/yyyy"))
        assertEquals("1/1/1904", f("0", "m/d/yyyy", date1904 = true))
        assertEquals("11/2/2027", f("45231", "m/d/yyyy", date1904 = true))
    }
}
