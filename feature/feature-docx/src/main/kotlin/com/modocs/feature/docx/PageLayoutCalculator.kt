package com.modocs.feature.docx

import android.content.Context
import android.graphics.Paint
import android.graphics.Typeface

/**
 * Calculates page layout positions using real font metrics.
 * Performs a full dry-run simulation of the PDF converter's rendering flow,
 * including word-by-word wrapping, mid-paragraph page splits, and row-by-row
 * table splitting. This ensures page breaks appear at identical positions
 * in both the in-app viewer and exported PDFs.
 */
class PageLayoutCalculator(private val context: Context) {

    companion object {
        /** Default A4 dimensions (used when no document page setup is available). */
        const val PAGE_WIDTH = 595f    // A4 width in points
        const val PAGE_HEIGHT = 842f   // A4 height in points
        const val MARGIN_LEFT = 72f
        const val MARGIN_RIGHT = 72f
        const val MARGIN_TOP = 72f
        const val MARGIN_BOTTOM = 72f
        val CONTENT_WIDTH = PAGE_WIDTH - MARGIN_LEFT - MARGIN_RIGHT
        val CONTENT_HEIGHT = PAGE_HEIGHT - MARGIN_TOP - MARGIN_BOTTOM
    }

    // Active layout values — default to companion constants, overridden via configure()
    var activePageWidth = PAGE_WIDTH; private set
    var activePageHeight = PAGE_HEIGHT; private set
    var activeMarginLeft = MARGIN_LEFT; private set
    var activeMarginRight = MARGIN_RIGHT; private set
    var activeMarginTop = MARGIN_TOP; private set
    var activeMarginBottom = MARGIN_BOTTOM; private set
    var activeContentWidth = CONTENT_WIDTH; private set

    /**
     * Apply document-specific page setup. Call before computeAutoPageBreaks
     * or any simulation methods.
     */
    fun configure(setup: PageSetup) {
        activePageWidth = setup.pageWidthPt
        activePageHeight = setup.pageHeightPt
        activeMarginLeft = setup.marginLeftPt
        activeMarginRight = setup.marginRightPt
        activeMarginTop = setup.marginTopPt
        activeMarginBottom = setup.marginBottomPt
        activeContentWidth = setup.contentWidthPt
    }

    private val typefaceCache = mutableMapOf<String, Typeface>()

    private val fontMapping = mapOf(
        "calibri" to "fonts/Carlito",
        "calibri light" to "fonts/Carlito",
        "cambria" to "fonts/Caladea",
        "arial" to "fonts/LiberationSans",
        "helvetica" to "fonts/LiberationSans",
        "times new roman" to "fonts/LiberationSerif",
        "times" to "fonts/LiberationSerif",
        "courier new" to "fonts/LiberationMono",
        "courier" to "fonts/LiberationMono",
    )

    /**
     * Tracks virtual page position during dry-run simulation.
     * Mirrors PdfState in DocxToPdfConverter exactly.
     */
    private inner class SimState {
        var pageNumber = 1
        var yPos = activeMarginTop
        var prevParaProps: ParagraphProperties? = null

        fun startNewPage() {
            pageNumber++
            yPos = activeMarginTop
        }

        fun needsNewPage(requiredHeight: Float): Boolean {
            return yPos + requiredHeight > activePageHeight - activeMarginBottom
        }
    }

    /**
     * Computes the set of element indices where an automatic page break
     * occurs BEFORE that element. Uses a full dry-run simulation that
     * mirrors the PDF converter's exact rendering flow.
     */
    fun computeAutoPageBreaks(document: DocxDocument): Set<Int> {
        configure(document.pageSetup)
        val breaks = mutableSetOf<Int>()
        val state = SimState()

        for ((index, element) in document.body.withIndex()) {
            val pageBefore = state.pageNumber

            when (element) {
                is DocxParagraph -> simParagraph(state, element, document)
                is DocxTable -> simTable(state, element)
                is DocxImage -> simImage(state, element, document)
            }

            // If this element caused a page change AND it wasn't due to
            // an explicit page break, record an auto page break.
            if (state.pageNumber > pageBefore) {
                val hasExplicitBreak = element is DocxParagraph &&
                    (element.properties.pageBreakBefore || element.runs.any { it.isPageBreak })

                if (!hasExplicitBreak) {
                    breaks.add(index)
                }
            }
        }

        return breaks
    }

    /**
     * Simulates paragraph rendering exactly as DocxToPdfConverter.drawParagraph
     * + flushTextParts do, tracking yPos and page breaks.
     */
    private fun simParagraph(state: SimState, paragraph: DocxParagraph, document: DocxDocument) {
        val props = paragraph.properties

        // Page break before (mirrors converter line 98)
        if (props.pageBreakBefore && state.yPos > activeMarginTop + 20) {
            state.startNewPage()
        }

        // Contextual spacing
        val prev = state.prevParaProps
        val suppressSpacing = prev != null &&
                (prev.contextualSpacing || props.contextualSpacing) &&
                prev.styleId != null && prev.styleId == props.styleId

        if (!suppressSpacing) state.yPos += props.spacingBeforePt

        val listPrefix = buildListPrefix(paragraph.listInfo, document.numbering)
        val indent = props.indentLeftTwips / 1440f * 72f
        val headingScale = when (props.headingLevel) {
            1 -> 1.8f; 2 -> 1.5f; 3 -> 1.3f; 4 -> 1.15f; else -> 1f
        }

        // Build text parts (same as converter)
        val parts = mutableListOf<Pair<String, RunProperties>>()
        if (listPrefix != null) {
            val rp = paragraph.runs.firstOrNull()?.properties ?: RunProperties()
            parts.add(listPrefix to rp)
        }
        for (run in paragraph.runs) {
            when {
                run.isPageBreak -> {
                    simFlushTextParts(state, parts, activeMarginLeft + indent, headingScale, props)
                    parts.clear()
                    state.startNewPage()
                }
                run.isBreak -> parts.add("\n" to run.properties)
                else -> parts.add(run.text to run.properties)
            }
        }

        simFlushTextParts(state, parts, activeMarginLeft + indent, headingScale, props)
        if (!suppressSpacing) state.yPos += props.spacingAfterPt
        state.prevParaProps = props
    }

    /**
     * Mirrors DocxToPdfConverter.flushTextParts exactly — word-by-word
     * wrapping with mid-paragraph page breaks.
     */
    private fun simFlushTextParts(
        state: SimState,
        parts: List<Pair<String, RunProperties>>,
        startX: Float,
        headingScale: Float,
        paraProps: ParagraphProperties,
    ) {
        if (parts.isEmpty()) return

        val indent = paraProps.indentLeftTwips / 1440f * 72f
        val rightIndent = paraProps.indentRightTwips / 1440f * 72f
        val firstLineIndent = paraProps.indentFirstLineTwips / 1440f * 72f
        val baseAvailWidth = activeContentWidth - indent - rightIndent

        var lineIndex = 0
        var lineWidth = 0f

        for ((text, rp) in parts) {
            if (text == "\n") {
                state.yPos += paraProps.effectiveLineHeight(14f)
                lineIndex++
                lineWidth = 0f
                continue
            }

            val paint = createPaint(rp, headingScale)

            if (text == "\t") {
                lineWidth += 36f
                continue
            }

            for (word in text.split(' ').filter { it.isNotEmpty() }) {
                val wordText = "$word "
                val wordWidth = paint.measureText(wordText)
                val availWidth = if (lineIndex == 0) baseAvailWidth + firstLineIndent else baseAvailWidth

                if (lineWidth + wordWidth > availWidth && lineWidth > 1f) {
                    state.yPos += paraProps.effectiveLineHeight(paint.textSize)
                    lineIndex++
                    lineWidth = 0f

                    if (state.needsNewPage(paint.textSize * 2)) {
                        state.startNewPage()
                    }
                }

                lineWidth += wordWidth
            }
        }

        val fontSize = parts.firstOrNull()?.let {
            createPaint(it.second, headingScale).textSize
        } ?: 11f
        state.yPos += paraProps.effectiveLineHeight(fontSize)
    }

    /**
     * Compute (x, width) positions for each cell in a row using grid column widths.
     */
    private fun computeCellPositions(cells: List<DocxTableCell>, gridColWidths: List<Int>): List<Pair<Float, Float>> {
        if (gridColWidths.isNotEmpty()) {
            val totalGridTwips = gridColWidths.sum()
            if (totalGridTwips > 0) {
                val result = mutableListOf<Pair<Float, Float>>()
                var x = activeMarginLeft
                var gridIdx = 0
                for (cell in cells) {
                    val span = cell.gridSpan.coerceAtLeast(1)
                    var cellTwips = 0
                    for (i in 0 until span) {
                        if (gridIdx + i < gridColWidths.size) cellTwips += gridColWidths[gridIdx + i]
                    }
                    gridIdx += span
                    val w = cellTwips.toFloat() / totalGridTwips * activeContentWidth
                    result.add(x to w)
                    x += w
                }
                return result
            }
        }
        val totalTwips = cells.sumOf { it.widthTwips ?: 0 }
        if (totalTwips > 0) {
            val result = mutableListOf<Pair<Float, Float>>()
            var x = activeMarginLeft
            for (cell in cells) {
                val w = if (cell.widthTwips != null && cell.widthTwips > 0) {
                    cell.widthTwips.toFloat() / totalTwips * activeContentWidth
                } else {
                    activeContentWidth / cells.size
                }
                result.add(x to w)
                x += w
            }
            return result
        }
        val w = activeContentWidth / cells.size.coerceAtLeast(1)
        return cells.mapIndexed { i, _ -> (activeMarginLeft + i * w) to w }
    }

    private fun simTable(state: SimState, table: DocxTable) {
        state.yPos += 8f

        for (row in table.rows) {
            val allContinue = row.cells.all { it.vMerge == VMergeType.CONTINUE }
            val positions = computeCellPositions(row.cells, table.gridColWidths)

            var maxH = 0f
            for ((ci, cell) in row.cells.withIndex()) {
                if (cell.vMerge == VMergeType.CONTINUE) continue
                val cellW = positions[ci].second
                val hasContent = cell.paragraphs.any { p -> p.runs.any { it.text.isNotBlank() } }
                val lines = cell.paragraphs.sumOf { it.text.lines().size.coerceAtLeast(1) }
                val cellH = if (hasContent) maxOf(lines * 14f + 8f, 16f) else 4f
                maxH = maxOf(maxH, cellH)
            }
            maxH = maxOf(maxH, if (allContinue) 0f else 4f)
            if (maxH < 1f) continue

            if (state.needsNewPage(maxH)) {
                state.startNewPage()
            }

            state.yPos += maxH
        }

        state.yPos += 8f
    }

    /**
     * Simulates image rendering. Mirrors DocxToPdfConverter.drawImage.
     */
    private fun simImage(state: SimState, image: DocxImage, document: DocxDocument) {
        val imageBytes = document.images[image.relationId] ?: return

        // We can't decode the bitmap here without overhead, so estimate from EMU
        val widthPt = if (image.widthEmu > 0) {
            (image.widthEmu / 914400f * 72f).coerceAtMost(CONTENT_WIDTH)
        } else {
            CONTENT_WIDTH
        }
        val heightPt = if (image.heightEmu > 0 && image.widthEmu > 0) {
            image.heightEmu.toFloat() / image.widthEmu.toFloat() * widthPt
        } else {
            200f
        }

        if (state.needsNewPage(heightPt + 8f)) {
            state.startNewPage()
        }

        state.yPos += 4f + heightPt + 4f
    }

    // --- Shared helpers (used by both calculator and converter) ---

    fun createPaint(props: RunProperties, headingScale: Float = 1f): Paint {
        val fontSize = (props.fontSizeSp ?: 11f) * headingScale
        return Paint().apply {
            textSize = fontSize
            typeface = resolveTypeface(props.fontName, props.bold, props.italic)
            color = props.color ?: android.graphics.Color.BLACK
            isAntiAlias = true
            isUnderlineText = props.underline
            isStrikeThruText = props.strikethrough
        }
    }

    fun resolveTypeface(fontName: String?, bold: Boolean, italic: Boolean): Typeface {
        val key = fontName?.lowercase()?.trim() ?: "calibri"
        val basePath = fontMapping[key] ?: fontMapping["calibri"]!!
        val suffix = when {
            bold && italic -> "-BoldItalic"
            bold -> "-Bold"
            italic -> "-Italic"
            else -> "-Regular"
        }
        val assetPath = "$basePath$suffix.ttf"

        return typefaceCache.getOrPut(assetPath) {
            try {
                Typeface.createFromAsset(context.assets, assetPath)
            } catch (_: Exception) {
                val style = when {
                    bold && italic -> Typeface.BOLD_ITALIC
                    bold -> Typeface.BOLD
                    italic -> Typeface.ITALIC
                    else -> Typeface.NORMAL
                }
                Typeface.create(Typeface.DEFAULT, style)
            }
        }
    }

    fun buildListPrefix(listInfo: ListInfo?, numbering: Map<String, NumberingDefinition>, counter: Int = 1): String? {
        if (listInfo == null) return null
        val numDef = numbering[listInfo.numId]
        val level = numDef?.levels?.get(listInfo.level)
        return if (level != null) {
            if (level.format == NumberFormat.BULLET) {
                val bullet = resolveBulletChar(level.text, listInfo.level)
                "$bullet  "
            } else {
                level.format.formatNumber(counter) + "  "
            }
        } else {
            val bullet = defaultBulletForLevel(listInfo.level)
            "$bullet  "
        }
    }

    /**
     * Resolve bullet character from numbering level text.
     * DOCX bullet text may use Symbol/Wingdings font chars that we can't render,
     * so we map common patterns to Unicode equivalents.
     */
    private fun resolveBulletChar(text: String, level: Int): String {
        return when {
            text == "•" || text == "·" -> "•"
            text == "o" -> "○"
            text == "–" || text == "-" -> "–"
            text == "■" || text == "▪" -> "■"
            text.isEmpty() -> defaultBulletForLevel(level)
            text.length == 1 && text[0].code > 0xF000 -> defaultBulletForLevel(level)
            else -> text
        }
    }

    private fun defaultBulletForLevel(level: Int): String = when (level % 3) {
        0 -> "•"
        1 -> "○"
        else -> "■"
    }
}
