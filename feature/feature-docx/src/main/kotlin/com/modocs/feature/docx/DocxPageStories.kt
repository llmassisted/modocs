package com.modocs.feature.docx

import android.graphics.Canvas
import android.graphics.Typeface
import android.text.*
import android.text.style.StyleSpan

/** Basic header/footer text, clipped to margins so it cannot obscure the body. */
fun drawPageStories(canvas: Canvas, document: DocxDocument) {
    val setup = document.pageSetup
    fun draw(paragraphs: List<DocxParagraph>, top: Float, bottom: Float) {
        if (paragraphs.isEmpty() || bottom <= top) return
        canvas.save()
        canvas.clipRect(setup.marginLeftPt, top, setup.pageWidthPt - setup.marginRightPt, bottom)
        canvas.translate(setup.marginLeftPt, top)
        for (paragraph in paragraphs) {
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
            canvas.translate(0f, layout.height + paragraph.properties.spacingAfterPt)
        }
        canvas.restore()
    }
    draw(document.headerParagraphs, 12f, setup.marginTopPt - 6f)
    draw(document.footerParagraphs, setup.pageHeightPt - setup.marginBottomPt + 6f, setup.pageHeightPt - 12f)
}
