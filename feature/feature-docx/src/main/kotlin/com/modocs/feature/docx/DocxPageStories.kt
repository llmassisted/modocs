package com.modocs.feature.docx

import android.graphics.Canvas
import android.graphics.Typeface
import android.text.*
import android.text.style.StyleSpan

/** Basic header/footer text and pictures, clipped to margins so they cannot obscure the body. */
fun drawPageStories(canvas: Canvas, document: DocxDocument) {
    val setup = document.pageSetup
    // Canvas units are points; a PDF canvas reports its size in points too, so keep at least 2 px per point.
    val pixelsPerPoint = maxOf(2f, canvas.width / setup.pageWidthPt)
    fun draw(elements: List<DocxElement>, top: Float, bottom: Float) {
        if (elements.isEmpty() || bottom <= top) return
        canvas.save()
        canvas.clipRect(setup.marginLeftPt, top, setup.pageWidthPt - setup.marginRightPt, bottom)
        canvas.translate(setup.marginLeftPt, top)
        var remaining = bottom - top
        for (element in elements) {
            if (remaining <= 0f) break
            if (element is DocxImage) {
                val bytes = document.images[element.partName]
                val box = imageBox(element, bytes, setup.contentWidthPt, remaining)
                val left = imageLeft(element, box, 0f, setup.contentWidthPt)
                drawDocxImage(canvas, element, bytes, left, 0f, box, pixelsPerPoint)
                canvas.translate(0f, box.height + 2f)
                remaining -= box.height + 2f
                continue
            }
            val paragraph = element as? DocxParagraph ?: continue
            val text = SpannableStringBuilder()
            for (run in paragraph.runs) {
                val start = text.length; text.append(run.text)
                val style = when { run.properties.bold && run.properties.italic -> Typeface.BOLD_ITALIC
                    run.properties.bold -> Typeface.BOLD; run.properties.italic -> Typeface.ITALIC; else -> Typeface.NORMAL }
                text.setSpan(StyleSpan(style), start, text.length, Spanned.SPAN_EXCLUSIVE_EXCLUSIVE)
            }
            val paint = TextPaint(android.graphics.Paint.ANTI_ALIAS_FLAG).apply {
                color = android.graphics.Color.BLACK
                textSize = paragraph.runs.firstOrNull()?.properties?.fontSizeSp ?: 10f
            }
            val alignment = when (paragraph.properties.alignment) {
                ParagraphAlignment.CENTER -> Layout.Alignment.ALIGN_CENTER
                ParagraphAlignment.RIGHT -> Layout.Alignment.ALIGN_OPPOSITE
                else -> Layout.Alignment.ALIGN_NORMAL
            }
            val layout = StaticLayout.Builder.obtain(text, 0, text.length, paint, setup.contentWidthPt.toInt().coerceAtLeast(1))
                .setAlignment(alignment).setIncludePad(false).build()
            layout.draw(canvas)
            val advance = layout.height + paragraph.properties.spacingAfterPt
            canvas.translate(0f, advance)
            remaining -= advance
        }
        canvas.restore()
    }
    draw(document.header, 12f, setup.marginTopPt - 6f)
    draw(document.footer, setup.pageHeightPt - setup.marginBottomPt + 6f, setup.pageHeightPt - 12f)
}
