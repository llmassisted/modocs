package com.modocs.app

import android.content.Intent
import android.os.Bundle
import android.widget.Toast
import androidx.activity.ComponentActivity
import androidx.activity.compose.setContent
import androidx.activity.enableEdgeToEdge
import androidx.compose.runtime.getValue
import androidx.compose.runtime.mutableStateOf
import androidx.compose.runtime.setValue
import com.modocs.core.common.DocumentType
import com.modocs.core.common.MAX_FILE_SIZE_BYTES
import com.modocs.core.common.getFileSize
import com.modocs.core.ui.theme.MoDocsTheme
import com.modocs.app.navigation.DocumentRequest
import com.modocs.app.navigation.MoDocsApp
import dagger.hilt.android.AndroidEntryPoint

@AndroidEntryPoint
class MainActivity : ComponentActivity() {

    // A one-shot request to open an externally-supplied document. Each incoming
    // intent produces a request with a fresh [DocumentRequest.token] so navigation
    // fires even when the same file is reopened, and survives activity recreation.
    private var documentRequest by mutableStateOf<DocumentRequest?>(null)
    private var requestCounter = 0

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        enableEdgeToEdge()

        handleIntent(intent)

        setContent {
            MoDocsTheme {
                MoDocsApp(documentRequest = documentRequest)
            }
        }
    }

    override fun onNewIntent(intent: Intent) {
        super.onNewIntent(intent)
        // Keep getIntent() current so a later activity recreation restores the
        // most recently opened document, not the one this activity launched with.
        setIntent(intent)
        handleIntent(intent)
    }

    private fun handleIntent(intent: Intent) {
        val uri = when (intent.action) {
            Intent.ACTION_VIEW, Intent.ACTION_EDIT -> intent.data
            else -> null
        } ?: return

        // Reject oversized files up front (external opens otherwise bypass the
        // picker's size check). Size is best-effort; allow through if unknown.
        val size = uri.getFileSize(this)
        if (size != null && size > MAX_FILE_SIZE_BYTES) {
            val maxMb = MAX_FILE_SIZE_BYTES / (1024 * 1024)
            Toast.makeText(this, "File too large (max ${maxMb}MB)", Toast.LENGTH_LONG).show()
            return
        }

        // Detect document type from MIME type or file name
        val mimeType = intent.type ?: contentResolver.getType(uri)
        val fileName = uri.lastPathSegment ?: ""
        val type = DocumentType.fromMimeType(mimeType)
            .takeIf { it != DocumentType.UNKNOWN }
            ?: DocumentType.fromFileName(fileName)

        documentRequest = DocumentRequest(uri = uri, type = type, token = ++requestCounter)
    }
}
