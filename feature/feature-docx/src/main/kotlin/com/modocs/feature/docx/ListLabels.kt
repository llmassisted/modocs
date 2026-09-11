package com.modocs.feature.docx

/** Resolve list counters once so page, reflow, and PDF renderers agree across page breaks. */
fun resolveListLabels(body: List<DocxElement>, numbering: Map<String, NumberingDefinition>): List<DocxElement> {
    val counters = mutableMapOf<String, MutableMap<Int, Int>>()
    fun paragraph(p: DocxParagraph): DocxParagraph {
        val info = p.listInfo ?: return p
        val levels = numbering[info.numId]?.levels ?: return p
        val level = levels[info.level] ?: return p
        val counts = counters.getOrPut(info.numId) { mutableMapOf() }
        counts.keys.filter { it > info.level }.forEach { counts.remove(it) }
        counts[info.level] = counts[info.level]?.plus(1) ?: level.start
        val label = if (level.format == NumberFormat.BULLET) null else Regex("%([1-9])").replace(level.text) { match ->
            val index = match.groupValues[1].toInt() - 1
            val definition = levels[index] ?: level
            definition.format.formatNumber(counts[index] ?: definition.start).removeSuffix(".")
        } + "  "
        return p.copy(listLabel = label, properties = if (p.properties.indentLeftTwips == 0)
            p.properties.copy(indentLeftTwips = level.indentTwips) else p.properties)
    }
    return body.map { element -> when (element) {
        is DocxParagraph -> paragraph(element)
        is DocxTable -> element.copy(rows = element.rows.map { row -> row.copy(cells = row.cells.map { cell -> cell.copy(paragraphs = cell.paragraphs.map(::paragraph)) }) })
        else -> element
    } }
}
