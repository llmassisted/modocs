package com.modocs.feature.docx

import android.content.Context
import android.graphics.Bitmap
import android.graphics.BitmapFactory
import android.graphics.Canvas
import android.graphics.Paint
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.sync.Mutex
import kotlinx.coroutines.sync.withLock
import kotlinx.coroutines.withContext

/**
 * Renders individual pages of a [DocxDocument] as [Bitmap] objects.
 *
 * Uses the same Canvas/Paint drawing code as [DocxToPdfConverter] to guarantee
 * the in-app viewer looks identical to the exported PDF.
 *
 * Two-pass approach:
 * 1. [prepare] — layout simulation to compute per-page element ranges
 * 2. [renderPage] — render a single page to a Bitmap on demand
 */
class DocxPageRenderer(private val context: Context) {

    companion object {
        private const val MAX_CACHE_SIZE = 6
    }

    // Active layout values — default to A4/1" margins, updated from document in prepare()
    private var PAGE_WIDTH = PageLayoutCalculator.PAGE_WIDTH
    private var PAGE_HEIGHT = PageLayoutCalculator.PAGE_HEIGHT
    private var MARGIN_LEFT = PageLayoutCalculator.MARGIN_LEFT
    private var MARGIN_RIGHT = PageLayoutCalculator.MARGIN_RIGHT
    private var MARGIN_TOP = PageLayoutCalculator.MARGIN_TOP
    private var MARGIN_BOTTOM = PageLayoutCalculator.MARGIN_BOTTOM
    private var CONTENT_WIDTH = PageLayoutCalculator.CONTENT_WIDTH

    private val layout = PageLayoutCalculator(context)
    private val mutex = Mutex()

    // Prepared page data
    private var pages: List<PageInfo> = emptyList()
    private var document: DocxDocument? = null

    val pageCount: Int get() = pages.size

    // LRU bitmap cache
    private val pageCache = LinkedHashMap<CacheKey, Bitmap>(16, 0.75f, true)

    data class CacheKey(val pageIndex: Int, val width: Int)

    /**
     * Per-page layout data computed during prepare().
     */
    data class PageInfo(
        /** Indices into document.body for elements starting on this page. */
        val elementStartIndex: Int,
        /** yPos at which to begin drawing the first element on this page. */
        val startYPos: Float,
        /** Text segments for search highlighting. Populated during render. */
        var textSegments: List<TextSegment> = emptyList(),
        /** Concatenated plain text of this page. Populated during render. */
        var pageText: String = "",
    )

    /**
     * A positioned piece of text on a page (for search highlights).
     * Coordinates are normalized (0..1 relative to page dimensions).
     */
    data class TextSegment(
        val text: String,
        val x: Float,
        val y: Float,
        val width: Float,
        val height: Float,
    )

    // ---------------------------------------------------------------
    // Pass 1: Layout simulation
    // ---------------------------------------------------------------

    /**
     * Compute page boundaries by simulating the full layout.
     * Must be called before [renderPage].
     */
    fun prepare(doc: DocxDocument) {
        document = doc

        // Apply document's page setup (margins, page size)
        val setup = doc.pageSetup
        PAGE_WIDTH = setup.pageWidthPt
        PAGE_HEIGHT = setup.pageHeightPt
        MARGIN_LEFT = setup.marginLeftPt
        MARGIN_RIGHT = setup.marginRightPt
        MARGIN_TOP = setup.marginTopPt
        MARGIN_BOTTOM = setup.marginBottomPt
        CONTENT_WIDTH = setup.contentWidthPt

        // Also configure the shared layout calculator
        layout.configure(setup)

        val result = mutableListOf<PageInfo>()
        val state = SimState()

        // Page 1 starts at element 0
        result.add(PageInfo(elementStartIndex = 0, startYPos = MARGIN_TOP))

        for ((index, element) in doc.body.withIndex()) {
            val pageBefore = state.pageNumber

            when (element) {
                is DocxParagraph -> simParagraph(state, element, doc)
                is DocxTable -> simTable(state, element)
                is DocxImage -> simImage(state, element, doc)
            }

            // If a new page was started during this element, record it
            if (state.pageNumber > pageBefore) {
                // The element that caused the page break may have started on the old page
                // and continued onto the new one — but for rendering purposes we track
                // the first element index that appears on each new page.
                // If yPos == MARGIN_TOP, this element starts fresh on the new page.
                // Otherwise, it was a mid-element split.
                val newPagesAdded = state.pageNumber - pageBefore
                for (p in 0 until newPagesAdded) {
                    val pageStartElement = if (p == 0 && state.yPos > MARGIN_TOP + 1f) {
                        index // Mid-element split: same element continues
                    } else {
                        index // Element that triggered the new page
                    }
                    // Only add if we don't already have enough pages
                    if (result.size < state.pageNumber) {
                        result.add(PageInfo(
                            elementStartIndex = pageStartElement,
                            startYPos = MARGIN_TOP,
                        ))
                    }
                }
            }
        }

        pages = result
    }

    private inner class SimState {
        var pageNumber = 1
        var yPos = MARGIN_TOP
        var prevParaProps: ParagraphProperties? = null

        fun startNewPage() {
            pageNumber++
            yPos = MARGIN_TOP
        }

        fun needsNewPage(requiredHeight: Float): Boolean {
            return yPos + requiredHeight > PAGE_HEIGHT - MARGIN_BOTTOM
        }
    }

    private fun simParagraph(state: SimState, paragraph: DocxParagraph, document: DocxDocument) {
        val props = paragraph.properties
        if (props.pageBreakBefore && state.yPos > MARGIN_TOP + 20) {
            state.startNewPage()
        }
        val prev = state.prevParaProps
        val suppressSpacing = prev != null &&
                (prev.contextualSpacing || props.contextualSpacing) &&
                prev.styleId != null && prev.styleId == props.styleId
        if (!suppressSpacing) state.yPos += props.spacingBeforePt
        val listPrefix = layout.buildListPrefix(paragraph.listInfo, document.numbering)
        val indent = props.indentLeftTwips / 1440f * 72f
        val headingScale = when (props.headingLevel) {
            1 -> 1.8f; 2 -> 1.5f; 3 -> 1.3f; 4 -> 1.15f; else -> 1f
        }
        val parts = mutableListOf<Pair<String, RunProperties>>()
        if (listPrefix != null) {
            val rp = paragraph.runs.firstOrNull()?.properties ?: RunProperties()
            parts.add(listPrefix to rp)
        }
        for (run in paragraph.runs) {
            when {
                run.isPageBreak -> {
                    simFlush(state, parts, MARGIN_LEFT + indent, headingScale, props)
                    parts.clear()
                    state.startNewPage()
                }
                run.isBreak -> parts.add("\n" to run.properties)
                else -> parts.add(run.text to run.properties)
            }
        }
        simFlush(state, parts, MARGIN_LEFT + indent, headingScale, props)
        if (!suppressSpacing) state.yPos += props.spacingAfterPt
        state.prevParaProps = props
    }

    private fun simFlush(
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
        val baseAvailWidth = CONTENT_WIDTH - indent - rightIndent

        var lineIndex = 0
        var lineWidth = 0f

        for ((text, rp) in parts) {
            if (text == "\n") {
                state.yPos += paraProps.effectiveLineHeight(14f)
                lineIndex++
                lineWidth = 0f
                continue
            }

            val paint = layout.createPaint(rp, headingScale)

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
            layout.createPaint(it.second, headingScale).textSize
        } ?: 11f
        state.yPos += paraProps.effectiveLineHeight(fontSize)
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
                val cellH = measureCellHeight(cell, cellW)
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

    private fun simImage(state: SimState, image: DocxImage, document: DocxDocument) {
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

    // ---------------------------------------------------------------
    // Pass 2: Bitmap rendering
    // ---------------------------------------------------------------

    /**
     * Render a specific page to a Bitmap at the given pixel width.
     * The height is determined by A4 aspect ratio.
     */
    suspend fun renderPage(pageIndex: Int, renderWidth: Int): Bitmap? {
        if (pageIndex < 0 || pageIndex >= pages.size) return null
        val doc = document ?: return null

        val cacheKey = CacheKey(pageIndex, renderWidth)
        pageCache[cacheKey]?.let { return it }

        return withContext(Dispatchers.Default) {
            mutex.withLock {
                pageCache[cacheKey]?.let { return@withContext it }

                try {
                    val renderHeight = (renderWidth * PAGE_HEIGHT / PAGE_WIDTH).toInt()
                    val bitmap = Bitmap.createBitmap(renderWidth, renderHeight, Bitmap.Config.RGB_565)
                    bitmap.eraseColor(android.graphics.Color.WHITE)

                    val canvas = Canvas(bitmap)
                    val scaleFactor = renderWidth / PAGE_WIDTH
                    canvas.scale(scaleFactor, scaleFactor)

                    val textSegments = mutableListOf<TextSegment>()
                    renderPageContent(canvas, pageIndex, doc, textSegments)

                    // Store text data for search
                    pages[pageIndex].textSegments = textSegments
                    pages[pageIndex].pageText = textSegments.joinToString("") { it.text }

                    // Evict oldest if cache full
                    if (pageCache.size >= MAX_CACHE_SIZE) {
                        val oldestKey = pageCache.keys.first()
                        pageCache.remove(oldestKey)?.recycle()
                    }
                    pageCache[cacheKey] = bitmap
                    bitmap
                } catch (_: Exception) {
                    null
                }
            }
        }
    }

    fun invalidateCache() {
        pageCache.values.forEach { it.recycle() }
        pageCache.clear()
        // Also clear text segments so they get recomputed
        pages.forEach { it.textSegments = emptyList(); it.pageText = "" }
    }

    fun getPageText(pageIndex: Int): String {
        return pages.getOrNull(pageIndex)?.pageText ?: ""
    }

    fun getTextSegments(pageIndex: Int): List<TextSegment> {
        return pages.getOrNull(pageIndex)?.textSegments ?: emptyList()
    }

    // ---------------------------------------------------------------
    // Drawing code — mirrors DocxToPdfConverter exactly
    // ---------------------------------------------------------------

    /**
     * Render all content for a single page. Replays the same drawing logic
     * as DocxToPdfConverter but only for elements on this page.
     */
    private fun renderPageContent(
        canvas: Canvas,
        pageIndex: Int,
        doc: DocxDocument,
        textSegments: MutableList<TextSegment>,
    ) {
        val startElementIdx = pages[pageIndex].elementStartIndex
        val endElementIdx = if (pageIndex + 1 < pages.size) {
            pages[pageIndex + 1].elementStartIndex
        } else {
            doc.body.size
        }

        // We need to replay from the start to get correct yPos state.
        // But that's expensive. Instead, we re-simulate from scratch for this page
        // by running the full draw starting from element 0 but only enabling
        // actual drawing for the target page.
        val state = DrawState(canvas, textSegments)

        for ((index, element) in doc.body.withIndex()) {
            if (state.currentPage > pageIndex + 1) break // Past our page

            when (element) {
                is DocxParagraph -> drawParagraph(state, element, doc, pageIndex)
                is DocxTable -> drawTable(state, element, doc, pageIndex)
                is DocxImage -> drawImage(state, element, doc, pageIndex)
            }
        }
    }

    /**
     * Mutable drawing state that tracks page and position.
     * Only draws when on the target page.
     */
    private inner class DrawState(
        val canvas: Canvas,
        val textSegments: MutableList<TextSegment>,
    ) {
        var currentPage = 0
        var yPos = MARGIN_TOP
        /** numId → (level → counter) for tracking numbered list positions. */
        val listCounters = mutableMapOf<String, MutableMap<Int, Int>>()
        /** Previous paragraph's properties for contextual spacing. */
        var prevParaProps: ParagraphProperties? = null

        fun startNewPage() {
            currentPage++
            yPos = MARGIN_TOP
        }

        fun needsNewPage(requiredHeight: Float): Boolean {
            return yPos + requiredHeight > PAGE_HEIGHT - MARGIN_BOTTOM
        }

        fun isOnPage(targetPage: Int): Boolean = currentPage == targetPage

        fun getListCounter(listInfo: ListInfo): Int {
            val counters = listCounters.getOrPut(listInfo.numId) { mutableMapOf() }
            // Reset deeper levels when this level is encountered
            counters.keys.filter { it > listInfo.level }.forEach { counters.remove(it) }
            counters[listInfo.level] = (counters[listInfo.level] ?: 0) + 1
            return counters[listInfo.level] ?: 1
        }
    }

    private fun drawParagraph(
        state: DrawState,
        paragraph: DocxParagraph,
        doc: DocxDocument,
        targetPage: Int,
    ) {
        val props = paragraph.properties

        if (props.pageBreakBefore && state.yPos > MARGIN_TOP + 20) {
            state.startNewPage()
        }

        // Contextual spacing: suppress spacingBefore if previous paragraph has
        // contextualSpacing and the same style
        val prev = state.prevParaProps
        val suppressSpacing = prev != null &&
                (prev.contextualSpacing || props.contextualSpacing) &&
                prev.styleId != null && prev.styleId == props.styleId

        if (!suppressSpacing) {
            state.yPos += props.spacingBeforePt
        }

        val counter = paragraph.listInfo?.let { state.getListCounter(it) } ?: 1
        val listPrefix = layout.buildListPrefix(paragraph.listInfo, doc.numbering, counter)
        val indent = props.indentLeftTwips / 1440f * 72f
        val headingScale = when (props.headingLevel) {
            1 -> 1.8f; 2 -> 1.5f; 3 -> 1.3f; 4 -> 1.15f; else -> 1f
        }

        val parts = mutableListOf<Pair<String, RunProperties>>()
        if (listPrefix != null) {
            val rp = paragraph.runs.firstOrNull()?.properties ?: RunProperties()
            parts.add(listPrefix to rp)
        }
        for (run in paragraph.runs) {
            when {
                run.isPageBreak -> {
                    flushTextParts(state, parts, MARGIN_LEFT + indent, headingScale, props, targetPage)
                    parts.clear()
                    state.startNewPage()
                }
                run.isBreak -> parts.add("\n" to run.properties)
                run.text == "\t" -> parts.add("\t" to run.properties)
                else -> parts.add(run.text to run.properties)
            }
        }

        flushTextParts(state, parts, MARGIN_LEFT + indent, headingScale, props, targetPage)

        // Contextual spacing: spacingAfter is suppressed when the next paragraph
        // shares the same style — but we don't know the next paragraph yet.
        // Instead, we always add spacingAfter here, and the *next* paragraph
        // suppresses its spacingBefore (which is the same effect).
        if (!suppressSpacing) {
            state.yPos += props.spacingAfterPt
        }

        state.prevParaProps = props
    }

    private data class WordEntry(val text: String, val width: Float, val paint: Paint)
    private class LineData {
        val words = mutableListOf<WordEntry>()
        val totalWidth: Float get() = words.sumOf { it.width.toDouble() }.toFloat()
    }

    private fun flushTextParts(
        state: DrawState,
        parts: List<Pair<String, RunProperties>>,
        startX: Float,
        headingScale: Float,
        paraProps: ParagraphProperties,
        targetPage: Int,
    ) {
        if (parts.isEmpty()) return

        val indent = paraProps.indentLeftTwips / 1440f * 72f
        val rightIndent = paraProps.indentRightTwips / 1440f * 72f
        val firstLineIndent = paraProps.indentFirstLineTwips / 1440f * 72f
        val baseAvailWidth = CONTENT_WIDTH - indent - rightIndent

        // --- Pass 1: Build word entries and split into lines ---
        val lines = mutableListOf(LineData())

        for ((text, rp) in parts) {
            if (text == "\n") {
                lines.add(LineData())
                continue
            }

            val paint = layout.createPaint(rp, headingScale)

            if (text == "\t") {
                lines.last().words.add(WordEntry("    ", 36f, paint))
                continue
            }

            for (word in text.split(' ').filter { it.isNotEmpty() }) {
                val wordText = "$word "
                val wordWidth = paint.measureText(wordText)
                val lineIndex = lines.size - 1
                val availWidth = if (lineIndex == 0) baseAvailWidth + firstLineIndent else baseAvailWidth

                if (lines.last().totalWidth + wordWidth > availWidth && lines.last().words.isNotEmpty()) {
                    lines.add(LineData())
                }
                lines.last().words.add(WordEntry(wordText, wordWidth, paint))
            }
        }

        // --- Pass 2: Draw lines with alignment ---
        for ((lineIndex, line) in lines.withIndex()) {
            if (line.words.isEmpty()) {
                state.yPos += paraProps.effectiveLineHeight(14f)
                continue
            }

            val fontSize = line.words.first().paint.textSize
            if (state.needsNewPage(fontSize * 2)) {
                state.startNewPage()
            }

            val lineWidth = line.totalWidth
            val extraFirstLine = if (lineIndex == 0) firstLineIndent else 0f
            val xStart = when (paraProps.alignment) {
                ParagraphAlignment.CENTER -> MARGIN_LEFT + (CONTENT_WIDTH - lineWidth) / 2
                ParagraphAlignment.RIGHT -> PAGE_WIDTH - MARGIN_RIGHT - rightIndent - lineWidth
                else -> startX + extraFirstLine
            }

            if (state.isOnPage(targetPage)) {
                var xPos = xStart
                for (entry in line.words) {
                    state.canvas.drawText(entry.text, xPos, state.yPos + entry.paint.textSize, entry.paint)

                    state.textSegments.add(TextSegment(
                        text = entry.text.trimEnd(),
                        x = xPos / PAGE_WIDTH,
                        y = state.yPos / PAGE_HEIGHT,
                        width = entry.width / PAGE_WIDTH,
                        height = entry.paint.textSize * 1.3f / PAGE_HEIGHT,
                    ))
                    xPos += entry.width
                }
            }

            state.yPos += paraProps.effectiveLineHeight(fontSize)
        }
    }

    /**
     * Compute (x, width) positions for each cell in a row using grid column widths.
     * Falls back to cell widthTwips if no grid is available.
     */
    private fun computeCellPositions(cells: List<DocxTableCell>, gridColWidths: List<Int>): List<Pair<Float, Float>> {
        // Use tblGrid if available — authoritative column widths
        if (gridColWidths.isNotEmpty()) {
            val totalGridTwips = gridColWidths.sum()
            if (totalGridTwips > 0) {
                val result = mutableListOf<Pair<Float, Float>>()
                var x = MARGIN_LEFT
                var gridIdx = 0
                for (cell in cells) {
                    val span = cell.gridSpan.coerceAtLeast(1)
                    var cellTwips = 0
                    for (i in 0 until span) {
                        if (gridIdx + i < gridColWidths.size) cellTwips += gridColWidths[gridIdx + i]
                    }
                    gridIdx += span
                    val w = cellTwips.toFloat() / totalGridTwips * CONTENT_WIDTH
                    result.add(x to w)
                    x += w
                }
                return result
            }
        }
        // Fallback: use cell widthTwips
        val totalTwips = cells.sumOf { it.widthTwips ?: 0 }
        if (totalTwips > 0) {
            val result = mutableListOf<Pair<Float, Float>>()
            var x = MARGIN_LEFT
            for (cell in cells) {
                val w = if (cell.widthTwips != null && cell.widthTwips > 0) {
                    cell.widthTwips.toFloat() / totalTwips * CONTENT_WIDTH
                } else {
                    CONTENT_WIDTH / cells.size
                }
                result.add(x to w)
                x += w
            }
            return result
        }
        val w = CONTENT_WIDTH / cells.size.coerceAtLeast(1)
        return cells.mapIndexed { i, _ -> (MARGIN_LEFT + i * w) to w }
    }

    private fun drawTable(
        state: DrawState,
        table: DocxTable,
        doc: DocxDocument,
        targetPage: Int,
    ) {
        state.yPos += 8f

        val tableBorders = table.properties.borders
        val totalRows = table.rows.size

        for ((rowIdx, row) in table.rows.withIndex()) {
            val positions = computeCellPositions(row.cells, table.gridColWidths)
            val isFirstRow = rowIdx == 0
            val isLastRow = rowIdx == totalRows - 1

            // Check if this row is entirely vMerge CONTINUE (empty continuation row)
            val allContinue = row.cells.all { it.vMerge == VMergeType.CONTINUE }

            var maxH = 0f
            for ((ci, cell) in row.cells.withIndex()) {
                if (cell.vMerge == VMergeType.CONTINUE) continue
                val cellW = positions[ci].second
                val cellH = measureCellHeight(cell, cellW)
                maxH = maxOf(maxH, cellH)
            }
            // Minimum height: 4pt for empty/continuation rows, 8pt for content rows
            maxH = maxOf(maxH, if (allContinue) 0f else 4f)

            if (maxH < 1f) {
                // Skip entirely empty continuation rows
                continue
            }

            if (state.needsNewPage(maxH)) {
                state.startNewPage()
            }

            if (state.isOnPage(targetPage)) {
                val canvas = state.canvas

                for ((ci, cell) in row.cells.withIndex()) {
                    val (cx, cellWidth) = positions[ci]
                    val isFirstCol = ci == 0
                    val isLastCol = ci == row.cells.lastIndex

                    // Cell background
                    val bgColor = cell.shading
                    if (bgColor != null) {
                        val fillPaint = Paint().apply {
                            style = Paint.Style.FILL
                            color = bgColor
                        }
                        canvas.drawRect(cx, state.yPos, cx + cellWidth, state.yPos + maxH, fillPaint)
                    }

                    // Resolve borders: cell overrides > table defaults
                    val cb = cell.borders
                    val topB = cb?.top ?: if (isFirstRow) tableBorders.top else tableBorders.insideH
                    val bottomB = cb?.bottom ?: if (isLastRow) tableBorders.bottom else tableBorders.insideH
                    val leftB = cb?.left ?: if (isFirstCol) tableBorders.left else tableBorders.insideV
                    val rightB = cb?.right ?: if (isLastCol) tableBorders.right else tableBorders.insideV

                    // For vMerge CONTINUE cells, suppress top/bottom borders
                    if (cell.vMerge == VMergeType.CONTINUE) {
                        drawBorder(canvas, leftB, cx, state.yPos, cx, state.yPos + maxH)
                        drawBorder(canvas, rightB, cx + cellWidth, state.yPos, cx + cellWidth, state.yPos + maxH)
                        continue
                    }

                    // Draw cell borders
                    drawBorder(canvas, topB, cx, state.yPos, cx + cellWidth, state.yPos)
                    drawBorder(canvas, bottomB, cx, state.yPos + maxH, cx + cellWidth, state.yPos + maxH)
                    drawBorder(canvas, leftB, cx, state.yPos, cx, state.yPos + maxH)
                    drawBorder(canvas, rightB, cx + cellWidth, state.yPos, cx + cellWidth, state.yPos + maxH)

                    var ty = state.yPos + 2f
                    for (para in cell.paragraphs) {
                        for (run in para.runs) {
                            if (run.text.isBlank()) continue
                            val paint = layout.createPaint(run.properties, 1f)
                            if (paint.textSize > cellWidth * 0.8f) paint.textSize = 10f

                            var txPos = cx + 2f
                            for (word in run.text.split(' ').filter { it.isNotEmpty() }) {
                                val wt = "$word "
                                val ww = paint.measureText(wt)
                                if (txPos + ww > cx + cellWidth - 2f && txPos > cx + 3f) {
                                    ty += paint.textSize * 1.2f
                                    txPos = cx + 2f
                                }
                                canvas.drawText(wt, txPos, ty + paint.textSize, paint)

                                state.textSegments.add(TextSegment(
                                    text = word,
                                    x = txPos / PAGE_WIDTH,
                                    y = ty / PAGE_HEIGHT,
                                    width = ww / PAGE_WIDTH,
                                    height = paint.textSize * 1.3f / PAGE_HEIGHT,
                                ))

                                txPos += ww
                            }
                            ty += paint.textSize * 1.2f
                        }
                    }
                }
            }

            state.yPos += maxH
        }

        state.yPos += 8f
    }

    /** Draw a single border line if the style is not NONE. */
    private fun drawBorder(canvas: android.graphics.Canvas, border: CellBorder, x1: Float, y1: Float, x2: Float, y2: Float) {
        if (border.style == BorderStyle.NONE) return
        val paint = Paint().apply {
            style = Paint.Style.STROKE
            strokeWidth = (border.widthEighthPt / 8f).coerceAtLeast(0.5f)
            color = border.color ?: android.graphics.Color.BLACK
        }
        canvas.drawLine(x1, y1, x2, y2, paint)
    }

    private fun measureCellHeight(cell: DocxTableCell, cellWidth: Float): Float {
        var h = 4f // padding (2pt top + 2pt bottom)
        var hasContent = false
        for (para in cell.paragraphs) {
            for (run in para.runs) {
                if (run.text.isBlank()) continue
                hasContent = true
                val paint = layout.createPaint(run.properties, 1f)
                if (paint.textSize > cellWidth * 0.8f) paint.textSize = 10f
                var lineWidth = 2f
                var lines = 1
                for (word in run.text.split(' ').filter { it.isNotEmpty() }) {
                    val ww = paint.measureText("$word ")
                    if (lineWidth + ww > cellWidth - 4f && lineWidth > 3f) {
                        lines++
                        lineWidth = 2f
                    }
                    lineWidth += ww
                }
                h += lines * paint.textSize * 1.2f
            }
            if (cell.paragraphs.indexOf(para) < cell.paragraphs.size - 1) {
                h += 4f // paragraph spacing
            }
        }
        return if (hasContent) maxOf(h, 16f) else 4f
    }

    private fun drawImage(
        state: DrawState,
        image: DocxImage,
        doc: DocxDocument,
        targetPage: Int,
    ) {
        val imageBytes = doc.images[image.relationId] ?: return

        val bitmap = try {
            BitmapFactory.decodeByteArray(imageBytes, 0, imageBytes.size)
        } catch (_: Exception) { return } ?: return

        val maxWidth = CONTENT_WIDTH
        val widthPt = if (image.widthEmu > 0) {
            (image.widthEmu / 914400f * 72f).coerceAtMost(maxWidth)
        } else {
            maxWidth.coerceAtMost(bitmap.width.toFloat())
        }
        val heightPt = bitmap.height * (widthPt / bitmap.width)

        if (state.needsNewPage(heightPt + 8f)) {
            state.startNewPage()
        }

        state.yPos += 4f

        if (state.isOnPage(targetPage)) {
            val dest = android.graphics.Rect(
                MARGIN_LEFT.toInt(), state.yPos.toInt(),
                (MARGIN_LEFT + widthPt).toInt(), (state.yPos + heightPt).toInt(),
            )
            state.canvas.drawBitmap(bitmap, null, dest, null)
        }

        bitmap.recycle()
        state.yPos += heightPt + 4f
    }
}
