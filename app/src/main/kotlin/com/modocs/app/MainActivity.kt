package com.modocs.app

import android.content.Intent
import android.net.Uri
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
import androidx.lifecycle.lifecycleScope
import com.modocs.core.common.AppPreferences
import com.modocs.core.storage.FileAccessManager
import com.modocs.core.storage.RecentFilesRepository
import dagger.hilt.android.AndroidEntryPoint
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.launch
import javax.inject.Inject

@AndroidEntryPoint
class MainActivity : ComponentActivity() {

    @Inject lateinit var recentFiles: RecentFilesRepository
    @Inject lateinit var fileAccess: FileAccessManager

    // A one-shot request to open an externally-supplied document. Each incoming
    // intent produces a request with a fresh [DocumentRequest.token] so navigation
    // fires even when the same file is reopened, and survives activity recreation.
    private var documentRequest by mutableStateOf<DocumentRequest?>(null)
    private var requestCounter = 0

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        enableEdgeToEdge()

        val pendingUri = savedInstanceState?.getString("pendingUri")
        val pendingType = savedInstanceState?.getString("pendingType")
        if (pendingUri != null && pendingType != null) {
            documentRequest = DocumentRequest(Uri.parse(pendingUri), DocumentType.valueOf(pendingType), ++requestCounter)
        } else if (savedInstanceState == null) {
            handleIntent(intent)
        }

        setContent {
            val prefs = com.modocs.core.ui.components.rememberAppPreferences()
            val dark = when (prefs.theme) { "Dark" -> true; "Light" -> false; else -> androidx.compose.foundation.isSystemInDarkTheme() }
            MoDocsTheme(darkTheme = dark, dynamicColor = prefs.dynamicColor) {
                MoDocsApp(documentRequest, onConsumed = { request ->
                    if (documentRequest == request) documentRequest = null
                }, onOpened = { uri, name, type ->
                    if (AppPreferences.get(this).state.value.keepRecents) {
                        lifecycleScope.launch(Dispatchers.IO) {
                            fileAccess.persistPermission(uri)
                            recentFiles.addRecentFile(uri.toString(), name, type, null)
                        }
                    }
                })
            }
        }
    }

    override fun onSaveInstanceState(outState: Bundle) {
        super.onSaveInstanceState(outState)
        documentRequest?.let {
            outState.putString("pendingUri", it.uri.toString())
            outState.putString("pendingType", it.type.name)
        }
    }

    override fun onNewIntent(intent: Intent) {
        super.onNewIntent(intent)
        // The navigation stack restores accepted opens; only pending requests need replay.
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
        val size = try {
            uri.getFileSize(this)
        } catch (_: SecurityException) {
            Toast.makeText(this, "File unavailable. Open it again from the app that shared it.", Toast.LENGTH_LONG).show()
            return
        }
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

        if (type == DocumentType.UNKNOWN) {
            Toast.makeText(this, "Unsupported file type", Toast.LENGTH_LONG).show()
            return
        }
        documentRequest = DocumentRequest(uri = uri, type = type, token = ++requestCounter)
    }
}
