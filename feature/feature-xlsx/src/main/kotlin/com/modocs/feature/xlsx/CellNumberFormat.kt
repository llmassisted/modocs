package com.modocs.feature.xlsx

import java.math.BigDecimal
import java.math.MathContext
import java.math.RoundingMode
import java.text.DecimalFormatSymbols
import java.time.DayOfWeek
import java.time.LocalDate
import java.time.Month
import java.time.chrono.IsoChronology
import java.time.format.DateTimeFormatter
import java.time.format.DateTimeFormatterBuilder
import java.time.format.FormatStyle
import java.time.format.TextStyle
import java.util.Locale
import kotlin.math.abs
import kotlin.math.floor
import kotlin.math.log10
import kotlin.math.pow

/**
 * Shows a cell's stored number the way spreadsheet apps do for its number format: percentages,
 * currency, thousands separators, fixed decimals, scientific notation, digit masks, dates and
 * times. Separators and month names follow the device locale, as Excel's do. Fractions fall
 * back to General.
 */
object CellNumberFormat {

    /** Formats files reference by id alone (ECMA-376 §18.8.30); ids 5-8 are Excel's en-US currency. */
    private val BUILT_IN = mapOf(
        1 to "0", 2 to "0.00", 3 to "#,##0", 4 to "#,##0.00",
        5 to "\"$\"#,##0_);\\(\"$\"#,##0\\)", 6 to "\"$\"#,##0_);[Red]\\(\"$\"#,##0\\)",
        7 to "\"$\"#,##0.00_);\\(\"$\"#,##0.00\\)", 8 to "\"$\"#,##0.00_);[Red]\\(\"$\"#,##0.00\\)",
        9 to "0%", 10 to "0.00%", 11 to "0.00E+00",
        15 to "d-mmm-yy", 16 to "d-mmm", 17 to "mmm-yy",
        18 to "h:mm AM/PM", 19 to "h:mm:ss AM/PM", 20 to "h:mm", 21 to "h:mm:ss",
        37 to "#,##0 ;(#,##0)", 38 to "#,##0 ;[Red](#,##0)", 39 to "#,##0.00;(#,##0.00)", 40 to "#,##0.00;[Red](#,##0.00)",
        45 to "mm:ss", 46 to "[h]:mm:ss", 47 to "mm:ss.0", 48 to "##0.0E+0",
    )

    /** Ids that mean "the system's short date", which Excel renders in the reader's locale. */
    private fun isShortDate(id: Int) = id == 14 || id in 27..36 || id in 50..58

    fun isDateFormat(formatId: Int, formatCode: String?): Boolean {
        if (formatCode == null) return isShortDate(formatId) || formatId in 15..22 || formatId in 45..47
        val first = sections(formatCode).first()
        // Drop quoted text, escapes and bracketed tags (keeping elapsed [h]/[m]/[s]) before looking for date letters.
        val bare = first.replace(Regex("\"[^\"]*\"|\\\\.|_.|\\*.")) { "" }
            .replace(Regex("\\[([^\\]]*)]")) { if (it.groupValues[1].matches(Regex("(?i)h+|m+|s+"))) "h" else "" }
        return bare.any { it.lowercaseChar() in "ymdhs" } && !bare.equals("General", ignoreCase = true)
    }

    /** [raw] as displayed under format [formatId] / [formatCode]; non-numbers are returned unchanged. */
    fun format(raw: String, formatId: Int, formatCode: String?, date1904: Boolean = false): String {
        val value = raw.toDoubleOrNull()?.takeIf { it.isFinite() } ?: return raw
        if (formatCode == null && (isShortDate(formatId) || formatId == 22)) return shortDate(value, formatId == 22, date1904) ?: general(value)
        val code = formatCode ?: BUILT_IN[formatId] ?: return general(value)
        val parts = sections(code).let { if (it.size > 3) it.take(3) else it } // a 4th section is for text
        val (section, signed) = pick(value, parts)
        val bare = section.trim()
        return when {
            bare.equals("General", ignoreCase = true) || bare == "@" -> if (signed) general(value) else general(abs(value))
            isDateFormat(0, section) -> formatDate(value, section, date1904) ?: general(value)
            else -> formatNumber(value, section, signed)
        }
    }

    /** Excel's General: whole numbers as-is, others to 10 significant digits, very large or small ones in E notation. */
    fun general(value: Double): String {
        val magnitude = abs(value)
        val text = when {
            magnitude == 0.0 -> "0"
            magnitude >= 1e11 || magnitude < 1e-9 -> {
                val exponent = floor(log10(magnitude)).toInt()
                val mantissa = BigDecimal(value / 10.0.pow(exponent)).round(MathContext(6)).stripTrailingZeros().toPlainString()
                mantissa + "E" + (if (exponent < 0) "-" else "+") + abs(exponent).toString().padStart(2, '0')
            }
            else -> BigDecimal.valueOf(value).round(MathContext(10)).stripTrailingZeros().toPlainString()
        }
        return text.replace('.', DecimalFormatSymbols.getInstance().decimalSeparator)
    }

    // --- Sections ---

    private fun sections(code: String): List<String> {
        val result = mutableListOf<String>()
        val current = StringBuilder()
        var quoted = false
        var bracket = false
        var i = 0
        while (i < code.length) {
            val c = code[i]
            if (c == '\\' && !quoted && i + 1 < code.length) {
                current.append(c).append(code[i + 1])
                i += 2
                continue
            }
            when {
                c == '"' -> quoted = !quoted
                c == '[' && !quoted -> bracket = true
                c == ']' && !quoted -> bracket = false
                c == ';' && !quoted && !bracket -> {
                    result.add(current.toString())
                    current.clear()
                    i++
                    continue
                }
            }
            current.append(c)
            i++
        }
        result.add(current.toString())
        return result
    }

    private val Condition = Regex("""\[(<=|>=|<>|<|>|=)\s*(-?[0-9.]+)]""")

    /** The section for [value], and whether a minus sign is still needed (a negative section draws its own). */
    private fun pick(value: Double, parts: List<String>): Pair<String, Boolean> {
        if (parts.any { Condition.containsMatchIn(it) }) {
            for (part in parts) {
                val match = Condition.find(part) ?: return part to true
                val limit = match.groupValues[2].toDouble()
                val matches = when (match.groupValues[1]) {
                    "<" -> value < limit; "<=" -> value <= limit; ">" -> value > limit
                    ">=" -> value >= limit; "=" -> value == limit; else -> value != limit
                }
                if (matches) return part to true
            }
            return parts.last() to true
        }
        return when {
            parts.size >= 3 && value == 0.0 -> parts[2] to false
            parts.size >= 2 && value < 0 -> parts[1] to false
            else -> parts[0] to true
        }
    }

    /** Text a bracketed tag contributes: the symbol of a currency tag like [$€-407]; colors and conditions show nothing. */
    private fun bracketText(content: String): String? =
        if (content.startsWith("$")) content.substring(1).substringBefore('-').ifEmpty { null } else null

    // --- Numbers ---

    private sealed interface Token
    private data class Literal(val text: String) : Token
    private data class Digit(val kind: Char) : Token
    private data object Point : Token
    private data object Comma : Token
    private data object Percent : Token
    private data class Exponent(val plus: Boolean) : Token

    private fun tokenize(section: String): List<Token> {
        val tokens = mutableListOf<Token>()
        var seenPoint = false
        var seenExponent = false
        var i = 0
        while (i < section.length) {
            val c = section[i]
            val next = section.getOrNull(i + 1)
            when {
                c == '"' -> {
                    val end = section.indexOf('"', i + 1).let { if (it < 0) section.length else it }
                    tokens += Literal(section.substring(i + 1, end))
                    i = end + 1
                    continue
                }
                c == '\\' && next != null -> { tokens += Literal(next.toString()); i += 2; continue }
                c == '_' && next != null -> { tokens += Literal(" "); i += 2; continue }
                c == '*' && next != null -> { i += 2; continue }
                c == '[' -> {
                    val end = section.indexOf(']', i).let { if (it < 0) section.length else it }
                    bracketText(section.substring(i + 1, end))?.let { tokens += Literal(it) }
                    i = end + 1
                    continue
                }
                c == '0' || c == '#' || c == '?' -> tokens += Digit(c)
                c == '.' && !seenPoint && !seenExponent -> { tokens += Point; seenPoint = true }
                c == ',' -> tokens += Comma
                c == '%' -> tokens += Percent
                (c == 'E' || c == 'e') && (next == '+' || next == '-') && !seenExponent -> {
                    tokens += Exponent(next == '+')
                    seenExponent = true
                    i += 2
                    continue
                }
                c == '@' -> Unit
                else -> tokens += Literal(c.toString())
            }
            i++
        }
        return tokens
    }

    private fun formatNumber(value: Double, section: String, signed: Boolean): String {
        val tokens = tokenize(section)
        if (tokens.none { it is Digit }) {
            // Only text, like "Yes" or "-": shown as written.
            return tokens.joinToString("") { (it as? Literal)?.text ?: if (it == Percent) "%" else "" }
        }
        // Fractions ("# ?/?") are not drawn; General keeps the value readable.
        val slash = tokens.indexOfFirst { it == Literal("/") }
        if (slash > 0 && tokens.take(slash).any { it is Digit } && tokens.drop(slash).any { it is Digit }) {
            return general(if (signed) value else abs(value))
        }

        val exponentAt = tokens.indexOfFirst { it is Exponent }
        val mantissaEnd = if (exponentAt < 0) tokens.size else exponentAt
        val pointAt = tokens.subList(0, mantissaEnd).indexOf(Point).let { if (it < 0) mantissaEnd else it }
        val intSlots = (0 until pointAt).filter { tokens[it] is Digit }
        val fracSlots = (pointAt until mantissaEnd).filter { tokens[it] is Digit }
        val expSlots = (mantissaEnd until tokens.size).filter { tokens[it] is Digit }

        var number = abs(value) * 100.0.pow(tokens.count { it == Percent })
        // Commas straight after the last digit placeholder divide by a thousand each ("#,##0," shows thousands).
        var k = (0 until mantissaEnd).last { tokens[it] is Digit } + 1
        while (k < mantissaEnd && (tokens[k] == Comma || tokens[k] == Point)) {
            if (tokens[k] == Comma) number /= 1000
            k++
        }
        val grouping = (0 until pointAt).any { i ->
            tokens[i] == Comma && intSlots.any { it < i } && intSlots.any { it > i }
        }

        var exponent = 0
        val engineering = exponentAt >= 0 && intSlots.size > 1 && intSlots.any { (tokens[it] as Digit).kind == '#' }
        if (exponentAt >= 0 && number != 0.0) {
            val magnitude = floor(log10(number)).toInt()
            val width = intSlots.size.coerceAtLeast(1)
            exponent = if (engineering) Math.floorDiv(magnitude, width) * width else magnitude - (width - 1)
            number /= 10.0.pow(exponent)
        }
        var rounded = BigDecimal.valueOf(number).setScale(fracSlots.size, RoundingMode.HALF_UP)
        if (exponentAt >= 0 && !engineering && rounded >= BigDecimal.TEN.pow(intSlots.size.coerceAtLeast(1))) {
            exponent++
            rounded = BigDecimal.valueOf(number / 10).setScale(fracSlots.size, RoundingMode.HALF_UP)
        }
        val plain = rounded.toPlainString()
        var intText = plain.substringBefore('.')
        val fracText = plain.substringAfter('.', "")
        if (intText == "0" && intSlots.none { (tokens[it] as Digit).kind == '0' }) intText = ""

        val symbols = DecimalFormatSymbols.getInstance()
        val pieces = HashMap<Int, String>()
        if (grouping) {
            val padded = intText.padStart(intSlots.count { (tokens[it] as Digit).kind == '0' }, '0')
            pieces[intSlots.first()] = padded.reversed().chunked(3).joinToString(symbols.groupingSeparator.toString()).reversed()
        } else {
            // Fill digits right to left so masks like 000-00-0000 keep their literals; the leftmost slot takes any overflow.
            var remaining = intText
            for ((n, slot) in intSlots.reversed().withIndex()) {
                val kind = (tokens[slot] as Digit).kind
                pieces[slot] = when {
                    n == intSlots.lastIndex && remaining.isNotEmpty() -> remaining.also { remaining = "" }
                    remaining.isNotEmpty() -> remaining.last().toString().also { remaining = remaining.dropLast(1) }
                    kind == '0' -> "0"
                    kind == '?' -> " "
                    else -> ""
                }
            }
        }
        for ((n, slot) in fracSlots.withIndex()) {
            val kind = (tokens[slot] as Digit).kind
            val trailingZero = fracText.substring(n).all { it == '0' }
            pieces[slot] = when {
                !trailingZero || kind == '0' -> fracText[n].toString()
                kind == '?' -> " "
                else -> ""
            }
        }
        expSlots.firstOrNull()?.let { first ->
            pieces[first] = abs(exponent).toString().padStart(expSlots.count { (tokens[it] as Digit).kind == '0' }, '0')
        }

        val out = StringBuilder()
        for ((index, token) in tokens.withIndex()) {
            out.append(when (token) {
                is Literal -> token.text
                is Digit -> pieces[index].orEmpty()
                Point -> (if (intSlots.isEmpty()) intText else "") + symbols.decimalSeparator
                Comma -> ""
                Percent -> "%"
                is Exponent -> "E" + if (exponent < 0) "-" else if (token.plus) "+" else ""
            })
        }
        val text = out.toString()
        return if (signed && value < 0 && text.any { it in '1'..'9' }) "-$text" else text
    }

    // --- Dates and times ---

    private fun shortDate(value: Double, withTime: Boolean, date1904: Boolean): String? {
        val parts = dateParts(value, 0, date1904) ?: return null
        val locale = Locale.getDefault()
        val pattern = DateTimeFormatterBuilder.getLocalizedDateTimePattern(FormatStyle.SHORT, null, IsoChronology.INSTANCE, locale)
            .replace(Regex("y+"), "yyyy")
        val date = LocalDate.of(parts.year, parts.month, parts.day.coerceIn(1, 28)).let {
            if (parts.day > 28) it.withDayOfMonth(parts.day.coerceAtMost(it.lengthOfMonth())) else it
        }
        val text = DateTimeFormatter.ofPattern(pattern, locale).format(date)
        return if (withTime) "$text ${parts.hour}:${parts.minute.toString().padStart(2, '0')}" else text
    }

    private class DateParts(
        val year: Int, val month: Int, val day: Int, val weekday: DayOfWeek?,
        val days: Long, val hour: Int, val minute: Int, val second: Int, val fraction: Int,
    )

    /** Split an Excel serial date into calendar and clock parts, rounded to [fractionDigits] of a second. */
    private fun dateParts(value: Double, fractionDigits: Int, date1904: Boolean): DateParts? {
        if (value < 0 || value > 2_958_465.999) return null // Excel shows ##### outside 1900-9999
        val perSecond = 10.0.pow(fractionDigits).toLong()
        val perDay = 86_400L * perSecond
        val units = Math.round(value * perDay)
        val days = Math.floorDiv(units, perDay)
        val ofDay = Math.floorMod(units, perDay)
        val seconds = (ofDay / perSecond).toInt()
        val (year, month, day, weekday) = when {
            date1904 -> LocalDate.of(1904, 1, 1).plusDays(days).let { listOf(it.year, it.monthValue, it.dayOfMonth, it.dayOfWeek) }
            days == 0L -> listOf(1900, 1, 0, null) // Excel's day zero, "1/0/1900"
            days == 60L -> listOf(1900, 2, 29, null) // the leap day Excel keeps for Lotus compatibility
            days < 60L -> LocalDate.of(1899, 12, 31).plusDays(days).let { listOf(it.year, it.monthValue, it.dayOfMonth, it.dayOfWeek) }
            else -> LocalDate.of(1899, 12, 30).plusDays(days).let { listOf(it.year, it.monthValue, it.dayOfMonth, it.dayOfWeek) }
        }
        return DateParts(year as Int, month as Int, day as Int, weekday as DayOfWeek?, days,
            seconds / 3600, seconds / 60 % 60, seconds % 60, (ofDay % perSecond).toInt())
    }

    private fun formatDate(value: Double, section: String, date1904: Boolean): String? {
        // Tokens are (text, isLiteral).
        val tokens = mutableListOf<Pair<String, Boolean>>()
        var i = 0
        while (i < section.length) {
            val c = section[i]
            val next = section.getOrNull(i + 1)
            when {
                c == '"' -> {
                    val end = section.indexOf('"', i + 1).let { if (it < 0) section.length else it }
                    tokens += section.substring(i + 1, end) to true
                    i = end + 1
                }
                c == '\\' && next != null -> { tokens += next.toString() to true; i += 2 }
                c == '_' && next != null -> { tokens += " " to true; i += 2 }
                c == '*' && next != null -> i += 2
                c == '[' -> {
                    val end = section.indexOf(']', i).let { if (it < 0) section.length else it }
                    val content = section.substring(i + 1, end)
                    if (content.matches(Regex("(?i)h+|m+|s+"))) tokens += "[${content.lowercase()}]" to false
                    else bracketText(content)?.let { tokens += it to true }
                    i = end + 1
                }
                section.regionMatches(i, "AM/PM", 0, 5, ignoreCase = true) -> { tokens += "AM/PM" to false; i += 5 }
                section.regionMatches(i, "A/P", 0, 3, ignoreCase = true) -> { tokens += "A/P" to false; i += 3 }
                c.lowercaseChar() in "ymdhs" -> {
                    var end = i
                    while (end < section.length && section[end].lowercaseChar() == c.lowercaseChar()) end++
                    tokens += section.substring(i, end).lowercase() to false
                    i = end
                }
                c == '.' && next == '0' && tokens.lastOrNull { !it.second }?.first?.let { it.startsWith("s") || it.startsWith("[s") } == true -> {
                    var end = i + 1
                    while (end < section.length && section[end] == '0') end++
                    tokens += section.substring(i, end) to false
                    i = end
                }
                else -> { tokens += c.toString() to true; i++ }
            }
        }

        val fractionDigits = tokens.filter { !it.second && it.first.startsWith(".") }.maxOfOrNull { it.first.length - 1 }?.coerceAtMost(3) ?: 0
        val parts = dateParts(value, fractionDigits, date1904) ?: return null
        val twelveHour = tokens.any { !it.second && (it.first == "AM/PM" || it.first == "A/P") }
        val locale = Locale.getDefault()
        val fields = tokens.withIndex().filter { !it.value.second }.map { it.index }

        // "m" means minutes right after an hour or right before seconds; otherwise it is the month.
        fun isMinute(index: Int): Boolean {
            val position = fields.indexOf(index)
            val before = fields.getOrNull(position - 1)?.let { tokens[it].first }
            val after = fields.getOrNull(position + 1)?.let { tokens[it].first }
            return before?.let { it.startsWith("h") || it.startsWith("[h") } == true ||
                after?.let { it.startsWith("s") || it.startsWith("[s") } == true
        }

        val out = StringBuilder()
        for ((index, token) in tokens.withIndex()) {
            val (text, literal) = token
            if (literal) { out.append(text); continue }
            val hour = if (twelveHour) (parts.hour % 12).let { if (it == 0) 12 else it } else parts.hour
            out.append(when {
                text == "AM/PM" -> if (parts.hour < 12) "AM" else "PM"
                text == "A/P" -> if (parts.hour < 12) "A" else "P"
                text.startsWith("[h") -> (parts.days * 24 + parts.hour).toString().padStart(text.length - 2, '0')
                text.startsWith("[m") -> ((parts.days * 24 + parts.hour) * 60 + parts.minute).toString().padStart(text.length - 2, '0')
                text.startsWith("[s") -> (((parts.days * 24 + parts.hour) * 60 + parts.minute) * 60 + parts.second).toString().padStart(text.length - 2, '0')
                text.startsWith(".") -> "." + parts.fraction.toString().padStart(fractionDigits, '0').take(text.length - 1)
                text.startsWith("y") -> if (text.length <= 2) (parts.year % 100).toString().padStart(2, '0') else parts.year.toString()
                text.startsWith("d") -> when (text.length) {
                    1 -> parts.day.toString()
                    2 -> parts.day.toString().padStart(2, '0')
                    3 -> parts.weekday?.getDisplayName(TextStyle.SHORT, locale).orEmpty()
                    else -> parts.weekday?.getDisplayName(TextStyle.FULL, locale).orEmpty()
                }
                text.startsWith("h") -> if (text.length == 1) hour.toString() else hour.toString().padStart(2, '0')
                text.startsWith("s") -> if (text.length == 1) parts.second.toString() else parts.second.toString().padStart(2, '0')
                text.startsWith("m") && text.length <= 2 && isMinute(index) ->
                    if (text.length == 1) parts.minute.toString() else parts.minute.toString().padStart(2, '0')
                text.startsWith("m") -> when (text.length) {
                    1 -> parts.month.toString()
                    2 -> parts.month.toString().padStart(2, '0')
                    3 -> Month.of(parts.month).getDisplayName(TextStyle.SHORT, locale)
                    5 -> Month.of(parts.month).getDisplayName(TextStyle.NARROW, locale)
                    else -> Month.of(parts.month).getDisplayName(TextStyle.FULL, locale)
                }
                else -> text
            })
        }
        return out.toString()
    }
}
