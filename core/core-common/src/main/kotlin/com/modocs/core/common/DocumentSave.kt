package com.modocs.core.common

import android.content.Context
import android.net.Uri
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.withContext
import java.io.File
import java.io.OutputStream

/** Finish serialization before touching the destination. SAF providers need not support atomic rename. */
suspend fun saveDocumentCopy(context: Context, uri: Uri, write: (OutputStream) -> Unit) =
    withContext(Dispatchers.IO) {
        val staged = File.createTempFile("document_save_", ".tmp", context.cacheDir)
        try {
            staged.outputStream().use(write)
            context.contentResolver.openOutputStream(uri, "wt")?.use { destination ->
                staged.inputStream().use { it.copyTo(destination) }
            } ?: error("Cannot open destination for writing")
        } finally {
            staged.delete()
        }
    }
