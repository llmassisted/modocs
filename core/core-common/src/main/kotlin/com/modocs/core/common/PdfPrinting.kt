package com.modocs.core.common

import android.content.Context
import android.os.*
import android.print.*
import com.tom_roush.pdfbox.android.PDFBoxResourceLoader
import com.tom_roush.pdfbox.pdmodel.PDDocument
import java.io.File
import java.util.concurrent.Executors

/** Print a prepared local copy, honoring the page ranges requested by Android. */
fun printPdf(context: Context, source: File, name: String) {
    PDFBoxResourceLoader.init(context)
    val manager = context.getSystemService(Context.PRINT_SERVICE) as PrintManager
    manager.print(name, object : PrintDocumentAdapter() {
        private val executor = Executors.newSingleThreadExecutor()
        private val main = Handler(Looper.getMainLooper())
        override fun onLayout(oldAttributes: PrintAttributes?, newAttributes: PrintAttributes?, signal: CancellationSignal,
            callback: LayoutResultCallback, extras: Bundle?) {
            executor.execute {
                try {
                    val count = PDDocument.load(source).use { it.numberOfPages }
                    main.post {
                        if (signal.isCanceled) callback.onLayoutCancelled()
                        else callback.onLayoutFinished(PrintDocumentInfo.Builder(name).setContentType(PrintDocumentInfo.CONTENT_TYPE_DOCUMENT).setPageCount(count).build(), true)
                    }
                } catch (e: Exception) { main.post { callback.onLayoutFailed(e.message ?: "Cannot prepare document") } }
            }
        }
        override fun onWrite(pages: Array<out PageRange>, destination: ParcelFileDescriptor, signal: CancellationSignal,
            callback: WriteResultCallback) {
            executor.execute {
                try {
                    PDDocument.load(source).use { original ->
                        PDDocument().use { selected ->
                            for (index in 0 until original.numberOfPages) {
                                if (signal.isCanceled) break
                                if (pages.any { index in it.start..it.end }) {
                                    val page = original.getPage(index)
                                    selected.importPage(page).resources = page.resources
                                }
                            }
                            if (!signal.isCanceled) ParcelFileDescriptor.AutoCloseOutputStream(destination).use { selected.save(it) }
                        }
                    }
                    main.post { if (signal.isCanceled) callback.onWriteCancelled() else callback.onWriteFinished(pages) }
                } catch (e: Exception) { main.post { callback.onWriteFailed(e.message ?: "Cannot print document") } }
            }
        }
        override fun onFinish() { executor.execute { source.delete() }; executor.shutdown() }
    }, null)
}
