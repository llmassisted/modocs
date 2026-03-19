package com.modocs.feature.docx

import android.content.Context
import android.graphics.BitmapFactory
import android.graphics.Canvas
import android.graphics.Paint
import android.graphics.pdf.PdfDocument
import android.net.Uri
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.withContext
import java.io.OutputStream

/**
 * Converts a [DocxDocument] to PDF using Android's [PdfDocument] API.
 * Renders content to Canvas with proper font support via bundled assets.
 */
class DocxToPdfConverter(private val context: Context) {

    // Active layout values — updated from document in convert()
    private var PAGE_WIDTH = PageLayoutCalculator.PAGE_WIDTH
    private var PAGE_HEIGHT = PageLayoutCalculator.PAGE_HEIGHT
    private var MARGIN_LEFT = PageLayoutCalculator.MARGIN_LEFT
    private var MARGIN_RIGHT = PageLayoutCalculator.MARGIN_RIGHT
    private var MARGIN_TOP = PageLayoutCalculator.MARGIN_TOP
    private var MARGIN_BOTTOM = PageLayoutCalculator.MARGIN_BOTTOM
    private var CONTENT_WIDTH = PageLayoutCalculator.CONTENT_WIDTH

    private val layout = PageLayoutCalculator(context)

    /**
     * Mutable state for the current page being drawn.
     */
    private inner class PdfState(val pdfDocument: PdfDocument) {
        var pageNumber = 0
        var page: PdfDocument.Page? = null
        var canvas: Canvas? = null
        var yPos = MARGIN_TOP
        /** numId → (level → counter) for tracking numbered list positions. */
        val listCounters = mutableMapOf<String, MutableMap<Int, Int>>()
        var prevParaProps: ParagraphProperties? = null

        fun getListCounter(listInfo: ListInfo): Int {
            val counters = listCounters.getOrPut(listInfo.numId) { mutableMapOf() }
            counters.keys.filter { it > listInfo.level }.forEach { counters.remove(it) }
            counters[listInfo.level] = (counters[listInfo.level] ?: 0) + 1
            return counters[listInfo.level] ?: 1
        }

        fun startNewPage() {
            finishCurrentPage()
            pageNumber++
            val pageInfo = PdfDocument.PageInfo.Builder(PAGE_WIDTH.toInt(), PAGE_HEIGHT.toInt(), pageNumber).create()
            page = pdfDocument.startPage(pageInfo)
            canvas = page!!.canvas
            yPos = MARGIN_TOP
        }

        fun finishCurrentPage() {
            page?.let { pdfDocument.finishPage(it) }
            page = null
            canvas = null
        }

        fun needsNewPage(requiredHeight: Float): Boolean {
            return yPos + requiredHeight > PAGE_HEIGHT - MARGIN_BOTTOM
        }
    }

    suspend fun convert(document: DocxDocument, outputStream: OutputStream) =
        withContext(Dispatchers.Default) {
            // Apply document's page setup
            val setup = document.pageSetup
            PAGE_WIDTH = setup.pageWidthPt
            PAGE_HEIGHT = setup.pageHeightPt
            MARGIN_LEFT = setup.marginLeftPt
            MARGIN_RIGHT = setup.marginRightPt
            MARGIN_TOP = setup.marginTopPt
            MARGIN_BOTTOM = setup.marginBottomPt
            CONTENT_WIDTH = setup.contentWidthPt
            layout.configure(setup)

            val pdfDocument = PdfDocument()
            val state = PdfState(pdfDocument)

            try {
                state.startNewPage()

                for (element in document.body) {
                    when (element) {
                        is DocxParagraph -> drawParagraph(state, element, document)
                        is DocxTable -> drawTable(state, element, document)
                        is DocxImage -> drawImage(state, element, document)
                    }
                }

                state.finishCurrentPage()
                pdfDocument.writeTo(outputStream)
            } finally {
                pdfDocument.close()
            }
        }

    suspend fun convertToUri(document: DocxDocument, outputUri: Uri) {
        withContext(Dispatchers.IO) {
            context.contentResolver.openOutputStream(outputUri)?.use { stream ->
                convert(document, stream)
            } ?: throw IllegalStateException("Cannot open output URI")
        }
    }

    // --- Paragraph ---

    private fun drawParagraph(state: PdfState, paragraph: DocxParagraph, document: DocxDocument) {
        if (state.canvas == null) return
        val props = paragraph.properties

        // Page break before
        if (props.pageBreakBefore && state.yPos > MARGIN_TOP + 20) {
            state.startNewPage()
        }

        // Contextual spacing
        val prev = state.prevParaProps
        val suppressSpacing = prev != null &&
                (prev.contextualSpacing || props.contextualSpacing) &&
                prev.styleId != null && prev.styleId == props.styleId

        if (!suppressSpacing) state.yPos += props.spacingBeforePt

        // List prefix with counter tracking
        val counter = paragraph.listInfo?.let { state.getListCounter(it) } ?: 1
        val listPrefix = layout.buildListPrefix(paragraph.listInfo, document.numbering, counter)

        // Indentation (twips → points: twips / 1440 * 72)
        val indent = props.indentLeftTwips / 1440f * 72f

        // Heading scale
        val headingScale = when (props.headingLevel) {
            1 -> 1.8f; 2 -> 1.5f; 3 -> 1.3f; 4 -> 1.15f; else -> 1f
        }

        // Build text parts
        val parts = mutableListOf<Pair<String, RunProperties>>()
        if (listPrefix != null) {
            val rp = paragraph.runs.firstOrNull()?.properties ?: RunProperties()
            parts.add(listPrefix to rp)
        }
        for (run in paragraph.runs) {
            when {
                run.isPageBreak -> {
                    // Flush what we have, then start new page
                    flushTextParts(state, parts, MARGIN_LEFT + indent, headingScale, props)
                    parts.clear()
                    state.startNewPage()
                }
                run.isBreak -> parts.add("\n" to run.properties)
                run.text == "\t" -> parts.add("\t" to run.properties)
                else -> parts.add(run.text to run.properties)
            }
        }

        flushTextParts(state, parts, MARGIN_LEFT + indent, headingScale, props)
        if (!suppressSpacing) state.yPos += props.spacingAfterPt
        state.prevParaProps = props
    }

    private data class WordEntry(val text: String, val width: Float, val paint: Paint)
    private class LineData {
        val words = mutableListOf<WordEntry>()
        val totalWidth: Float get() = words.sumOf { it.width.toDouble() }.toFloat()
    }

    private fun flushTextParts(
        state: PdfState,
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

            var xPos = xStart
            for (entry in line.words) {
                state.canvas?.drawText(entry.text, xPos, state.yPos + entry.paint.textSize, entry.paint)
                xPos += entry.width
            }

            state.yPos += paraProps.effectiveLineHeight(fontSize)
        }
    }

    // --- Table ---

    /**
     * Compute (x, width) positions for each cell in a row using grid column widths.
     */
    private fun computeCellPositions(cells: List<DocxTableCell>, gridColWidths: List<Int>): List<Pair<Float, Float>> {
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

    private fun drawTable(state: PdfState, table: DocxTable, document: DocxDocument) {
        if (state.canvas == null) return
        state.yPos += 8f

        val tableBorders = table.properties.borders
        val totalRows = table.rows.size

        for ((rowIdx, row) in table.rows.withIndex()) {
            val allContinue = row.cells.all { it.vMerge == VMergeType.CONTINUE }
            val positions = computeCellPositions(row.cells, table.gridColWidths)
            val isFirstRow = rowIdx == 0
            val isLastRow = rowIdx == totalRows - 1

            // Measure row height — skip vMerge CONTINUE cells
            var maxH = 0f
            for ((ci, cell) in row.cells.withIndex()) {
                if (cell.vMerge == VMergeType.CONTINUE) continue
                val cellW = positions[ci].second
                maxH = maxOf(maxH, measureCellHeight(cell, cellW))
            }
            maxH = maxOf(maxH, if (allContinue) 0f else 4f)
            if (maxH < 1f) continue

            if (state.needsNewPage(maxH)) {
                state.startNewPage()
            }

            // Always read canvas from state — it changes after startNewPage()
            val canvas = state.canvas ?: return

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
                            txPos += ww
                        }
                        ty += paint.textSize * 1.2f
                    }
                }
            }

            state.yPos += maxH
        }

        state.yPos += 8f
    }

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
        var h = 4f // 2pt top + 2pt bottom padding
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
        }
        return if (hasContent) maxOf(h, 14f) else 4f
    }

    // --- Image ---

    private fun drawImage(state: PdfState, image: DocxImage, document: DocxDocument) {
        if (state.canvas == null) return
        val imageBytes = document.images[image.relationId] ?: return

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

        val canvas = state.canvas ?: return

        state.yPos += 4f
        val dest = android.graphics.Rect(
            MARGIN_LEFT.toInt(), state.yPos.toInt(),
            (MARGIN_LEFT + widthPt).toInt(), (state.yPos + heightPt).toInt(),
        )
        canvas.drawBitmap(bitmap, null, dest, null)
        bitmap.recycle()

        state.yPos += heightPt + 4f
    }

}
