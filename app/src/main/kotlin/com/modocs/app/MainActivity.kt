package com.modocs.app

import android.content.Intent
import android.net.Uri
import android.os.Bundle
import androidx.activity.ComponentActivity
import androidx.activity.compose.setContent
import androidx.activity.enableEdgeToEdge
import androidx.compose.runtime.getValue
import androidx.compose.runtime.mutableStateOf
import androidx.compose.runtime.setValue
import com.modocs.core.common.DocumentType
import com.modocs.core.ui.theme.MoDocsTheme
import com.modocs.app.navigation.MoDocsApp
import dagger.hilt.android.AndroidEntryPoint

@AndroidEntryPoint
class MainActivity : ComponentActivity() {

    private var documentUri by mutableStateOf<Uri?>(null)
    private var documentType by mutableStateOf<DocumentType?>(null)
    private var isOpenedExternally by mutableStateOf(false)

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        enableEdgeToEdge()

        handleIntent(intent)

        setContent {
            MoDocsTheme {
                MoDocsApp(
                    initialDocumentUri = documentUri,
                    initialDocumentType = documentType,
                    isOpenedExternally = isOpenedExternally,
                )
            }
        }
    }

    override fun onNewIntent(intent: Intent) {
        super.onNewIntent(intent)
        handleIntent(intent)
    }

    private fun handleIntent(intent: Intent) {
        val uri = when (intent.action) {
            Intent.ACTION_VIEW, Intent.ACTION_EDIT -> intent.data
            else -> null
        }
        if (uri != null) {
            documentUri = uri
            isOpenedExternally = true
            // Detect document type from MIME type or file name
            val mimeType = intent.type ?: contentResolver.getType(uri)
            val fileName = uri.lastPathSegment ?: ""
            documentType = DocumentType.fromMimeType(mimeType)
                .takeIf { it != DocumentType.UNKNOWN }
                ?: DocumentType.fromFileName(fileName)
        }
    }
}
