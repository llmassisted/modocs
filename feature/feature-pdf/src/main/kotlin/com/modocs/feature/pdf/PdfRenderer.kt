package com.modocs.feature.pdf

import android.content.Context
import android.graphics.Bitmap
import android.graphics.pdf.PdfRenderer as AndroidPdfRenderer
import android.net.Uri
import android.os.ParcelFileDescriptor
import android.util.LruCache
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.withContext

/**
 * Wrapper around Android's PdfRenderer that handles page rendering
 * with caching and thread safety.
 */
sealed class PdfOpenResult {
    data class Success(val wrapper: PdfRendererWrapper) : PdfOpenResult()
    data object PasswordRequired : PdfOpenResult()
    data class Failed(val message: String) : PdfOpenResult()
}

class PdfRendererWrapper private constructor(
    private val fileDescriptor: ParcelFileDescriptor,
    private val renderer: AndroidPdfRenderer,
) : AutoCloseable {

    val pageCount: Int = renderer.pageCount

    private val lock = Any()
    private var closed = false

    // LruCache synchronizes lookups independently of native page rendering.
    private val pageCache = LruCache<CacheKey, Bitmap>(8)

    data class CacheKey(val pageIndex: Int, val width: Int)

    internal fun cachedPage(pageIndex: Int, renderWidth: Int): Bitmap? =
        pageCache.get(CacheKey(pageIndex, renderWidth))

    /**
     * Render a page at a given width, maintaining aspect ratio.
     * Cached for repeated access (scrolling back and forth).
     */
    suspend fun renderPage(pageIndex: Int, renderWidth: Int): Bitmap? {
        if (pageIndex < 0 || pageIndex >= pageCount) return null

        val cacheKey = CacheKey(pageIndex, renderWidth)
        pageCache.get(cacheKey)?.let { return it }

        return withContext(Dispatchers.IO) {
            synchronized(lock) {
                if (closed) return@withContext null
                pageCache.get(cacheKey)?.let { return@withContext it }

                try {
                    // A page left open makes every later openPage() fail, so close it even if rendering throws.
                    val bitmap = renderer.openPage(pageIndex).use { page ->
                        val aspectRatio = page.height.toFloat() / page.width.toFloat()
                        val renderHeight = (renderWidth * aspectRatio).toInt()

                        val bitmap = Bitmap.createBitmap(
                            renderWidth,
                            renderHeight,
                            Bitmap.Config.ARGB_8888,
                        )
                        // White background
                        bitmap.eraseColor(android.graphics.Color.WHITE)

                        page.render(
                            bitmap,
                            null,
                            null,
                            AndroidPdfRenderer.Page.RENDER_MODE_FOR_DISPLAY,
                        )
                        bitmap
                    }

                    pageCache.put(cacheKey, bitmap)

                    bitmap
                } catch (e: Exception) {
                    null
                }
            }
        }
    }

    /**
     * Get page dimensions without rendering (for layout calculations).
     */
    suspend fun getPageDimensions(pageIndex: Int): Pair<Int, Int>? {
        if (pageIndex < 0 || pageIndex >= pageCount) return null
        return withContext(Dispatchers.IO) {
            synchronized(lock) {
                if (closed) return@withContext null
                try {
                    renderer.openPage(pageIndex).use { page -> page.width to page.height }
                } catch (e: Exception) {
                    null
                }
            }
        }
    }

    override fun close() = synchronized(lock) {
        if (closed) return@synchronized
        closed = true
        // An outgoing page may still be drawing a cached bitmap.
        pageCache.evictAll()
        try { renderer.close() } catch (_: Exception) {}
        try { fileDescriptor.close() } catch (_: Exception) {}
    }

    companion object {
        /**
         * Open a PDF from a content URI.
         */
        suspend fun open(context: Context, uri: Uri): PdfOpenResult {
            return withContext(Dispatchers.IO) {
                // Both calls below can raise SecurityException but they mean
                // completely different things, so they are opened separately:
                // the resolver throws when the URI grant is gone, while
                // AndroidPdfRenderer throws when the document is encrypted.
                val fd = try {
                    context.contentResolver.openFileDescriptor(uri, "r")
                        ?: return@withContext PdfOpenResult.Failed("Cannot open file")
                } catch (_: SecurityException) {
                    return@withContext PdfOpenResult.Failed(
                        "No longer have permission to read this file. " +
                            "Open it again from the app or folder it came from."
                    )
                } catch (e: Exception) {
                    return@withContext PdfOpenResult.Failed(e.message ?: "Cannot open file")
                }

                try {
                    PdfOpenResult.Success(PdfRendererWrapper(fd, AndroidPdfRenderer(fd)))
                } catch (_: SecurityException) {
                    fd.closeQuietly()
                    PdfOpenResult.PasswordRequired
                } catch (e: Exception) {
                    fd.closeQuietly()
                    PdfOpenResult.Failed(e.message ?: "Failed to open PDF")
                }
            }
        }

        private fun ParcelFileDescriptor.closeQuietly() {
            try { close() } catch (_: Exception) {}
        }

        /**
         * Open a PDF from a [File] (used for decrypted temp files).
         */
        suspend fun openFile(file: java.io.File): PdfOpenResult {
            return withContext(Dispatchers.IO) {
                val fd = try {
                    ParcelFileDescriptor.open(file, ParcelFileDescriptor.MODE_READ_ONLY)
                } catch (e: Exception) {
                    return@withContext PdfOpenResult.Failed(e.message ?: "Cannot open file")
                }

                try {
                    PdfOpenResult.Success(PdfRendererWrapper(fd, AndroidPdfRenderer(fd)))
                } catch (e: Exception) {
                    fd.closeQuietly()
                    PdfOpenResult.Failed(e.message ?: "Failed to open PDF")
                }
            }
        }
    }
}
