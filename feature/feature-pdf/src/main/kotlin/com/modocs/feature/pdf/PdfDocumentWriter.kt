package com.modocs.feature.pdf

import com.modocs.core.common.saveDocumentCopy
import com.tom_roush.pdfbox.android.PDFBoxResourceLoader
import com.tom_roush.pdfbox.pdmodel.PDDocument
import com.tom_roush.pdfbox.pdmodel.PDPageContentStream
import com.tom_roush.pdfbox.pdmodel.graphics.image.LosslessFactory
import com.tom_roush.pdfbox.util.Matrix
import android.content.Context
import android.graphics.Bitmap
import android.graphics.Canvas
import android.graphics.Color
import android.graphics.Paint
import android.graphics.Path
import android.graphics.Typeface
import android.net.Uri
import kotlin.math.sqrt

/** Append new marks without rasterizing original pages or changing their physical size. */
object PdfDocumentWriter {

    /** Longest edge we will rasterise, in pixels. */
    private const val MAX_RENDER_EDGE = 4000L

    /**
     * Total pixel budget for a single page. ARGB_8888 is 4 bytes per pixel, so
     * this is roughly a 48 MB allocation — comfortably above A4/Letter at
     * 300 DPI (~8.7 MP) but far below what an extreme page can ask for.
     */
    private const val MAX_RENDER_PIXELS = 12_000_000L

    /**
     * Pick a render size for a page, at 300 DPI where that fits in budget.
     *
     * Capping width alone left an extreme-aspect-ratio page — a long banner, a
     * spliced scan, both perfectly valid PDFs — free to demand a
     * multi-hundred-megabyte bitmap and take the process down, so the longest
     * edge and the total area are both clamped, preserving aspect ratio.
     */
    private fun renderSizeFor(pageWidth: Int, pageHeight: Int): Pair<Int, Int> {
        if (pageWidth <= 0 || pageHeight <= 0) return 1 to 1

        // Native PDF coordinates are 72 DPI points.
        var w = pageWidth.toLong() * 300 / 72
        var h = pageHeight.toLong() * 300 / 72
        if (w <= 0 || h <= 0) return 1 to 1

        val longest = maxOf(w, h)
        if (longest > MAX_RENDER_EDGE) {
            w = w * MAX_RENDER_EDGE / longest
            h = h * MAX_RENDER_EDGE / longest
        }

        if (w * h > MAX_RENDER_PIXELS) {
            val scale = sqrt(MAX_RENDER_PIXELS.toDouble() / (w.toDouble() * h.toDouble()))
            w = (w * scale).toLong()
            h = (h * scale).toLong()
        }

        return w.coerceAtLeast(1).toInt() to h.coerceAtLeast(1).toInt()
    }

    suspend fun save(
        context: Context,
        sourceUri: Uri,
        outputUri: Uri,
        pageCount: Int,
        annotations: List<PdfAnnotation>,
        formValues: Map<String, String> = emptyMap(),
    ) = saveDocumentCopy(context, outputUri) { output ->
        PDFBoxResourceLoader.init(context)
        context.contentResolver.openInputStream(sourceUri)?.use { input ->
            PDDocument.load(input).use { document ->
                require(document.numberOfPages == pageCount) { "The source PDF changed; reopen it before saving" }
                PdfForms.applyValues(document, formValues)
                for ((pageIndex, pageAnnotations) in annotations.groupBy { it.pageIndex }) {
                    val page = document.getPage(pageIndex)
                    val box = page.cropBox
                    val rotation = ((page.rotation % 360) + 360) % 360
                    val sideways = rotation == 90 || rotation == 270
                    val displayW = if (sideways) box.height else box.width
                    val displayH = if (sideways) box.width else box.height
                    val (w, h) = renderSizeFor(displayW.toInt(), displayH.toInt())
                    // Only the new marks are rasterized. Original page streams, text, links,
                    // forms, crop boxes and physical dimensions remain in the PDF.
                    val overlay = Bitmap.createBitmap(w, h, Bitmap.Config.ARGB_8888)
                    try {
                        val canvas = Canvas(overlay)
                        pageAnnotations.forEach { drawAnnotation(canvas, it, w, h) }
                        val image = LosslessFactory.createFromImage(document, overlay)
                        PDPageContentStream(document, page, PDPageContentStream.AppendMode.APPEND, true, true).use { stream ->
                            val matrix = when (rotation) {
                                90 -> Matrix(0f, box.height, -box.width, 0f, box.upperRightX, box.lowerLeftY)
                                180 -> Matrix(-box.width, 0f, 0f, -box.height, box.upperRightX, box.upperRightY)
                                270 -> Matrix(0f, -box.height, box.width, 0f, box.lowerLeftX, box.upperRightY)
                                else -> Matrix(box.width, 0f, 0f, box.height, box.lowerLeftX, box.lowerLeftY)
                            }
                            stream.drawImage(image, matrix)
                        }
                    } finally { overlay.recycle() }
                }
                document.save(output)
            }
        } ?: error("Cannot read the source PDF")
    }

    private fun drawAnnotation(
        canvas: Canvas,
        annotation: PdfAnnotation,
        pageWidth: Int,
        pageHeight: Int,
    ) {
        val px = annotation.x * pageWidth
        val py = annotation.y * pageHeight

        when (annotation) {
            is TextAnnotation -> {
                val paint = Paint(Paint.ANTI_ALIAS_FLAG).apply {
                    color = Color.BLACK
                    textSize = annotation.fontSizeSp * (pageWidth / 595f)
                    typeface = Typeface.DEFAULT
                }
                canvas.drawText(annotation.text, px, py + paint.textSize, paint)
            }

            is SignatureAnnotation -> {
                val sigWidth = annotation.width * pageWidth
                val sigHeight = annotation.height * pageHeight
                val paint = Paint(Paint.ANTI_ALIAS_FLAG).apply {
                    color = Color.BLACK
                    strokeWidth = 2.5f * (pageWidth / 595f)
                    style = Paint.Style.STROKE
                    strokeCap = Paint.Cap.ROUND
                    strokeJoin = Paint.Join.ROUND
                }

                for (stroke in annotation.strokes) {
                    if (stroke.size < 2) continue
                    val path = Path().apply {
                        val first = stroke.first()
                        moveTo(
                            px + first.x * sigWidth,
                            py + first.y * sigHeight,
                        )
                        for (i in 1 until stroke.size) {
                            lineTo(
                                px + stroke[i].x * sigWidth,
                                py + stroke[i].y * sigHeight,
                            )
                        }
                    }
                    canvas.drawPath(path, paint)
                }
            }

            is CheckmarkAnnotation -> {
                val size = annotation.sizeSp * (pageWidth / 595f)
                val paint = Paint(Paint.ANTI_ALIAS_FLAG).apply {
                    color = Color.BLACK
                    strokeWidth = 2.5f * (pageWidth / 595f)
                    style = Paint.Style.STROKE
                    strokeCap = Paint.Cap.ROUND
                    strokeJoin = Paint.Join.ROUND
                }

                if (annotation.isCheck) {
                    val path = Path().apply {
                        moveTo(px, py + size * 0.5f)
                        lineTo(px + size * 0.35f, py + size * 0.85f)
                        lineTo(px + size, py + size * 0.1f)
                    }
                    canvas.drawPath(path, paint)
                } else {
                    canvas.drawLine(px, py, px + size, py + size, paint)
                    canvas.drawLine(px + size, py, px, py + size, paint)
                }
            }

            is DateAnnotation -> {
                val paint = Paint(Paint.ANTI_ALIAS_FLAG).apply {
                    color = Color.BLACK
                    textSize = annotation.fontSizeSp * (pageWidth / 595f)
                    typeface = Typeface.DEFAULT
                }
                canvas.drawText(annotation.dateText, px, py + paint.textSize, paint)
            }
        }
    }
}
