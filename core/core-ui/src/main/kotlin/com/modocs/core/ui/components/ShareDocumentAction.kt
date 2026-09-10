package com.modocs.core.ui.components

import android.content.ActivityNotFoundException
import android.net.Uri
import android.widget.Toast
import androidx.compose.foundation.layout.size
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.filled.Share
import androidx.compose.material3.CircularProgressIndicator
import androidx.compose.material3.Icon
import androidx.compose.material3.IconButton
import androidx.compose.runtime.Composable
import androidx.compose.runtime.getValue
import androidx.compose.runtime.mutableStateOf
import androidx.compose.runtime.remember
import androidx.compose.runtime.rememberCoroutineScope
import androidx.compose.runtime.setValue
import androidx.compose.ui.Modifier
import androidx.compose.ui.platform.LocalContext
import androidx.compose.ui.unit.dp
import com.modocs.core.common.buildShareIntent
import kotlinx.coroutines.CancellationException
import kotlinx.coroutines.launch

/**
 * Top-bar action that sends the open document to another app.
 *
 * The document is shared as it exists on disk. A viewer with unsaved edits
 * therefore shares the last saved version, which is why this sits next to the
 * save actions rather than replacing them.
 *
 * Errors surface as a toast: the failure modes here are "no app can receive
 * this" and "the copy failed", both of which are momentary and need no action
 * beyond telling the user it didn't happen.
 */
@Composable
fun ShareDocumentAction(
    uri: Uri,
    displayName: String?,
    modifier: Modifier = Modifier,
    enabled: Boolean = true,
) {
    val context = LocalContext.current
    val scope = rememberCoroutineScope()
    var isSharing by remember { mutableStateOf(false) }

    IconButton(
        onClick = {
            // Building the intent can take a moment for a large file:// copy,
            // and a second chooser on top of the first helps nobody.
            if (isSharing) return@IconButton
            isSharing = true
            scope.launch {
                try {
                    context.startActivity(buildShareIntent(context, uri, displayName))
                } catch (e: CancellationException) {
                    throw e
                } catch (e: ActivityNotFoundException) {
                    Toast.makeText(
                        context,
                        "No app available to share this document",
                        Toast.LENGTH_LONG,
                    ).show()
                } catch (e: Exception) {
                    val reason = e.message?.takeIf { it.isNotBlank() }
                    Toast.makeText(
                        context,
                        reason?.let { "Couldn't share this document: $it" }
                            ?: "Couldn't share this document",
                        Toast.LENGTH_LONG,
                    ).show()
                } finally {
                    isSharing = false
                }
            }
        },
        modifier = modifier,
        enabled = enabled && !isSharing,
    ) {
        if (isSharing) {
            CircularProgressIndicator(modifier = Modifier.size(20.dp), strokeWidth = 2.dp)
        } else {
            Icon(Icons.Default.Share, contentDescription = "Share document")
        }
    }
}
