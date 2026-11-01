package com.modocs.feature.docx

import android.graphics.Canvas
import android.graphics.Paint
import android.graphics.RectF
import android.text.Layout
import android.text.StaticLayout
import android.text.TextPaint
import com.modocs.core.common.bitmapPixelSize
import com.modocs.core.common.decodeSampledBitmap

private const val EMU_PER_POINT = 12700f
private const val PLACEHOLDER_HEIGHT = 48f

/** An image's laid-out size, in points. */
internal data class ImageBox(val width: Float, val height: Float)

/**
 * The size [image] is laid out at, in points, no larger than [maxWidth] × [maxHeight].
 *
 * Every layout and drawing pass sizes images through this; if pagination and drawing
 * disagreed, content would land on the wrong page. The drawing's declared size wins,
 * then the picture's own pixel size at 96 dpi, then a placeholder strip for charts and
 * pictures that are missing or cannot be decoded.
 */
internal fun imageBox(image: DocxImage, bytes: ByteArray?, maxWidth: Float, maxHeight: Float): ImageBox {
    var width = image.widthEmu / EMU_PER_POINT
    var height = image.heightEmu / EMU_PER_POINT
    if (width <= 0f || height <= 0f) {
        val pixels = bytes?.let { bitmapPixelSize(it) }
        width = pixels?.first?.times(0.75f) ?: maxWidth
        height = pixels?.second?.times(0.75f) ?: PLACEHOLDER_HEIGHT
    }
    val scale = minOf(1f, maxWidth.coerceAtLeast(1f) / width, maxHeight.coerceAtLeast(1f) / height)
    return ImageBox(width * scale, height * scale)
}

/** Left edge for [box] in a column starting at [left] that is [width] wide. */
internal fun imageLeft(image: DocxImage, box: ImageBox, left: Float, width: Float): Float = when (image.alignment) {
    ParagraphAlignment.CENTER -> left + (width - box.width) / 2f
    ParagraphAlignment.RIGHT -> left + width - box.width
    else -> left
}

/** What a placeholder says when there is no picture to draw. */
internal fun placeholderLabel(image: DocxImage, bytes: ByteArray?): String =
    image.altText.ifBlank { if (bytes == null) "Picture not available" else "Picture format not supported" }

/**
 * Draw [image] into [box] at [left], [top], decoding at [pixelsPerPoint] — the canvas's
 * resolution — so pictures cost no more memory than the pixels they cover. A picture that
 * is missing or cannot be decoded becomes a labelled placeholder rather than a blank gap.
 */
internal fun drawDocxImage(
    canvas: Canvas, image: DocxImage, bytes: ByteArray?,
    left: Float, top: Float, box: ImageBox, pixelsPerPoint: Float,
) {
    val dest = RectF(left, top, left + box.width, top + box.height)
    val bitmap = bytes?.let { decodeSampledBitmap(it, (box.width * pixelsPerPoint).toInt()) }
    if (bitmap != null) {
        canvas.drawBitmap(bitmap, null, dest, Paint(Paint.FILTER_BITMAP_FLAG))
        return
    }
    canvas.drawRect(dest, Paint().apply { color = 0xFFF1F1F1.toInt() })
    canvas.drawRect(dest, Paint(Paint.ANTI_ALIAS_FLAG).apply {
        style = Paint.Style.STROKE
        strokeWidth = 0.75f
        color = 0xFFBDBDBD.toInt()
    })
    val label = placeholderLabel(image, bytes)
    val paint = TextPaint(Paint.ANTI_ALIAS_FLAG).apply {
        color = 0xFF616161.toInt()
        textSize = 9f
    }
    val layout = StaticLayout.Builder.obtain(label, 0, label.length, paint, (box.width - 8f).toInt().coerceAtLeast(1))
        .setAlignment(Layout.Alignment.ALIGN_CENTER).setIncludePad(false).build()
    canvas.save()
    canvas.clipRect(dest)
    canvas.translate(left + 4f, top + ((box.height - layout.height) / 2f).coerceAtLeast(2f))
    layout.draw(canvas)
    canvas.restore()
}
