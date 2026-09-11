package com.modocs.feature.docx

/** Keep the original runs on both sides of a text replacement, including XML source identities. */
fun replaceParagraphText(runs: List<DocxRun>, text: String): List<DocxRun> {
    val old = runs.joinToString("") { it.text }
    var prefix = 0
    while (prefix < minOf(old.length, text.length) && old[prefix] == text[prefix]) prefix++
    var suffix = 0
    while (suffix < minOf(old.length, text.length) - prefix && old[old.lastIndex - suffix] == text[text.lastIndex - suffix]) suffix++
    fun slice(start: Int, end: Int): List<DocxRun> {
        var offset = 0
        return runs.mapNotNull { run ->
            val from = (start - offset).coerceAtLeast(0)
            val to = (end - offset).coerceAtMost(run.text.length)
            offset += run.text.length
            if (from < to) run.copy(text = run.text.substring(from, to)) else null
        }
    }
    val before = slice(0, prefix)
    val after = slice(old.length - suffix, old.length)
    val inserted = text.substring(prefix, text.length - suffix)
    val template = before.lastOrNull() ?: after.firstOrNull() ?: runs.firstOrNull() ?: DocxRun("")
    return before + (if (inserted.isEmpty()) emptyList() else listOf(template.copy(text = inserted))) + after
}
