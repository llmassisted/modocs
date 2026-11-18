package com.modocs.core.storage

import android.content.Context
import android.content.Intent
import android.net.Uri
import com.modocs.core.common.DocumentType
import com.modocs.core.common.MAX_FILE_SIZE_BYTES
import com.modocs.core.common.getFileSize
import dagger.hilt.android.qualifiers.ApplicationContext
import java.io.InputStream
import java.io.OutputStream
import javax.inject.Inject
import javax.inject.Singleton

@Singleton
class FileAccessManager @Inject constructor(
    @ApplicationContext private val context: Context,
) {

    fun openInputStream(uri: Uri): InputStream? {
        return context.contentResolver.openInputStream(uri)
    }

    fun openOutputStream(uri: Uri): OutputStream? {
        return context.contentResolver.openOutputStream(uri, "wt")
    }

    fun getMimeType(uri: Uri): String? {
        return context.contentResolver.getType(uri)
    }

    /** True when a later launch can reopen this URI without the sender's temporary grant. */
    fun persistPermission(uri: Uri): Boolean {
        if (uri.scheme == "file") return uri.path?.let { java.io.File(it).canRead() } == true
        if (uri.scheme != "content") return false
        return try {
            context.contentResolver.takePersistableUriPermission(
                uri,
                Intent.FLAG_GRANT_READ_URI_PERMISSION or Intent.FLAG_GRANT_WRITE_URI_PERMISSION,
            )
            true
        } catch (_: SecurityException) {
            // Read-only permission fallback
            try {
                context.contentResolver.takePersistableUriPermission(
                    uri,
                    Intent.FLAG_GRANT_READ_URI_PERMISSION,
                )
                true
            } catch (_: SecurityException) {
                context.contentResolver.persistedUriPermissions.any { it.uri == uri && it.isReadPermission }
            }
        }
    }

    fun validateFileSize(uri: Uri): Boolean {
        val size = uri.getFileSize(context) ?: return true // allow if unknown
        return size <= MAX_FILE_SIZE_BYTES
    }

    companion object {
        val SUPPORTED_MIME_TYPES = DocumentType.entries
            .filter { it != DocumentType.UNKNOWN }
            .flatMap { listOf(it.mimeType) + it.alternateMimeTypes }
            .toTypedArray()
    }
}
