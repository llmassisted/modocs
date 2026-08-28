package com.modocs.feature.pdf

import android.content.Context
import android.graphics.Bitmap
import android.graphics.pdf.PdfRenderer as AndroidPdfRenderer
import android.net.Uri
import android.os.ParcelFileDescriptor
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.sync.Mutex
import kotlinx.coroutines.sync.withLock
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

    val pageCount: Int get() = renderer.pageCount

    private val mutex = Mutex()

    // LRU-style cache: pageIndex -> rendered bitmap
    private val pageCache = LinkedHashMap<CacheKey, Bitmap>(16, 0.75f, true)
    private val maxCacheSize = 8

    data class CacheKey(val pageIndex: Int, val width: Int)

    /**
     * Render a page at a given width, maintaining aspect ratio.
     * Cached for repeated access (scrolling back and forth).
     */
    suspend fun renderPage(pageIndex: Int, renderWidth: Int): Bitmap? {
        if (pageIndex < 0 || pageIndex >= pageCount) return null

        val cacheKey = CacheKey(pageIndex, renderWidth)
        pageCache[cacheKey]?.let { return it }

        return withContext(Dispatchers.IO) {
            mutex.withLock {
                // Double-check after acquiring lock
                pageCache[cacheKey]?.let { return@withContext it }

                try {
                    val page = renderer.openPage(pageIndex)
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
                    page.close()

                    // Evict oldest if cache full
                    if (pageCache.size >= maxCacheSize) {
                        val oldestKey = pageCache.keys.first()
                        pageCache.remove(oldestKey)?.recycle()
                    }
                    pageCache[cacheKey] = bitmap

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
            mutex.withLock {
                try {
                    val page = renderer.openPage(pageIndex)
                    val dims = page.width to page.height
                    page.close()
                    dims
                } catch (e: Exception) {
                    null
                }
            }
        }
    }

    override fun close() {
        pageCache.values.forEach { it.recycle() }
        pageCache.clear()
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
