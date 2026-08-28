package com.modocs.feature.pdf

import android.content.Context
import android.graphics.Bitmap
import android.graphics.Canvas
import android.graphics.Color
import android.graphics.Paint
import android.graphics.Path
import android.graphics.Typeface
import android.graphics.pdf.PdfDocument
import android.graphics.pdf.PdfRenderer as AndroidPdfRenderer
import android.net.Uri
import android.os.ParcelFileDescriptor
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.withContext
import java.io.File
import kotlin.math.sqrt

/**
 * Writes a filled/signed PDF by rendering each page bitmap
 * with annotations flattened on top.
 *
 * Uses its own PdfRenderer instance (separate from the viewer) to avoid
 * mutex contention and cache pollution. Renders one page at a time and
 * recycles bitmaps immediately to avoid OOM on large documents.
 */
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
    ) = withContext(Dispatchers.IO) {
        // Open a dedicated renderer for saving — separate from the viewer's renderer
        val fd = openReadDescriptor(context, sourceUri)

        val reader = AndroidPdfRenderer(fd)
        val pdfDoc = PdfDocument()

        try {
            for (pageIndex in 0 until pageCount) {
                val srcPage = reader.openPage(pageIndex)
                val pageWidth = srcPage.width
                val pageHeight = srcPage.height

                val (renderW, renderH) = renderSizeFor(pageWidth, pageHeight)

                val bitmap = Bitmap.createBitmap(renderW, renderH, Bitmap.Config.ARGB_8888)
                bitmap.eraseColor(Color.WHITE)
                srcPage.render(bitmap, null, null, AndroidPdfRenderer.Page.RENDER_MODE_FOR_PRINT)
                srcPage.close()

                // Create output page at the rendered size
                val pageInfo = PdfDocument.PageInfo.Builder(renderW, renderH, pageIndex + 1).create()
                val page = pdfDoc.startPage(pageInfo)
                val canvas = page.canvas

                canvas.drawBitmap(bitmap, 0f, 0f, null)
                bitmap.recycle() // free immediately

                // Draw annotations for this page
                val pageAnnotations = annotations.filter { it.pageIndex == pageIndex }
                for (annotation in pageAnnotations) {
                    drawAnnotation(canvas, annotation, renderW, renderH)
                }

                pdfDoc.finishPage(page)
            }

            context.contentResolver.openOutputStream(outputUri)?.use { outputStream ->
                pdfDoc.writeTo(outputStream)
            }
        } finally {
            pdfDoc.close()
            reader.close()
            fd.close()
        }
    }

    private fun openReadDescriptor(context: Context, uri: Uri): ParcelFileDescriptor {
        if (uri.scheme == "file") {
            val path = uri.path ?: throw IllegalStateException("Cannot open source PDF")
            return ParcelFileDescriptor.open(File(path), ParcelFileDescriptor.MODE_READ_ONLY)
        }
        return context.contentResolver.openFileDescriptor(uri, "r")
            ?: throw IllegalStateException("Cannot open source PDF")
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
