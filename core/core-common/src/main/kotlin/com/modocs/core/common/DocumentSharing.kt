package com.modocs.core.common

import android.content.ClipData
import android.content.ContentResolver
import android.content.Context
import android.content.Intent
import android.net.Uri
import androidx.core.content.FileProvider
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.withContext
import java.io.File
import java.io.IOException

/**
 * Cache subdirectory holding copies of documents made for sharing.
 *
 * Shared so the code that writes these copies and the startup sweep that
 * deletes them cannot drift apart — a rename on one side alone would leave
 * copies of the user's documents on disk indefinitely.
 */
const val SHARE_CACHE_DIR = "shared"

/**
 * Build a chooser Intent that sends [uri] to another app.
 *
 * A `content://` document is forwarded as it stands, with a read grant attached,
 * so nothing is duplicated even for a 100 MB spreadsheet. Anything else — in
 * practice a `file://` URI arriving from a VIEW intent — is copied into
 * [SHARE_CACHE_DIR] and handed over through our own FileProvider, because
 * handing a raw `file://` URI to another app throws FileUriExposedException on
 * API 24+.
 *
 * The copy runs on [Dispatchers.IO], so call this from a coroutine rather than
 * building the intent on the main thread.
 */
suspend fun buildShareIntent(
    context: Context,
    uri: Uri,
    displayName: String? = null,
): Intent {
    val name = shareFileName(context, uri, displayName)
    val mimeType = context.contentResolver.getType(uri)
        ?: DocumentType.fromFileName(name).mimeType

    val shareUri = if (uri.scheme == ContentResolver.SCHEME_CONTENT) {
        uri
    } else {
        copyToShareCache(context, uri, name)
    }

    val send = Intent(Intent.ACTION_SEND).apply {
        type = mimeType
        putExtra(Intent.EXTRA_STREAM, shareUri)
        putExtra(Intent.EXTRA_TITLE, name)
        // Receivers and the chooser preview read the grant off clipData; with
        // EXTRA_STREAM alone some of them are handed a URI they cannot open.
        clipData = ClipData.newRawUri(name, shareUri)
        addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
    }

    return Intent.createChooser(send, "Share $name").apply {
        addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
    }
}

/**
 * The name the receiving app should see, with anything that could escape the
 * cache directory stripped out — a provider is free to return whatever it likes
 * as a display name, including something with a path separator in it.
 */
private fun shareFileName(context: Context, uri: Uri, displayName: String?): String {
    val raw = displayName?.takeIf { it.isNotBlank() }
        ?: uri.getFileName(context)
        ?: uri.lastPathSegment
        ?: "document"

    return raw.substringAfterLast('/')
        .substringAfterLast('\\')
        .replace(Regex("""[\x00-\x1f:*?"<>|]"""), "_")
        .trim()
        .ifEmpty { "document" }
}

private suspend fun copyToShareCache(context: Context, uri: Uri, name: String): Uri =
    withContext(Dispatchers.IO) {
        val dir = File(context.cacheDir, SHARE_CACHE_DIR)
        if (!dir.isDirectory && !dir.mkdirs()) {
            throw IOException("Could not create the share cache directory")
        }

        val target = File(dir, name)
        try {
            val input = context.contentResolver.openInputStream(uri)
                ?: throw IOException("Could not read \"$name\"")
            input.use { source ->
                target.outputStream().use { destination -> source.copyTo(destination) }
            }
        } catch (t: Throwable) {
            // A half-written copy is worse than none: it would be shared as if
            // whole, and the sweep only runs at process start.
            target.delete()
            throw t
        }

        FileProvider.getUriForFile(context, "${context.packageName}.fileprovider", target)
    }
