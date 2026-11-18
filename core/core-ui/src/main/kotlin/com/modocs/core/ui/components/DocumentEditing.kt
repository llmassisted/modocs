package com.modocs.core.ui.components

import android.net.Uri
import android.widget.Toast
import androidx.activity.compose.BackHandler
import androidx.activity.compose.rememberLauncherForActivityResult
import androidx.activity.result.contract.ActivityResultContracts
import androidx.compose.foundation.layout.*
import androidx.compose.foundation.rememberScrollState
import androidx.compose.foundation.horizontalScroll
import androidx.compose.material3.*
import androidx.compose.runtime.*
import androidx.compose.runtime.saveable.rememberSaveable
import androidx.compose.ui.Modifier
import androidx.compose.ui.platform.LocalContext
import androidx.compose.ui.unit.dp
import com.modocs.core.common.buildShareIntent
import kotlinx.coroutines.launch

class DocumentSwitch(val token: Int, val open: () -> Unit, val cancel: () -> Unit)

@Composable
fun ReportDocumentOpened(uri: Uri, loaded: Boolean, name: String, onOpened: (Uri, String) -> Unit) {
    var lastReported by rememberSaveable { mutableStateOf<Uri?>(null) }
    if (loaded && uri != lastReported) SideEffect {
        onOpened(uri, name)
        lastReported = uri
    }
}

class DocumentActions(val saveCopy: () -> Unit, val saveAs: (String) -> Unit, val share: () -> Unit, val back: () -> Unit)

/** A copy becomes the active document only after saving succeeds. */
@Composable
fun rememberDocumentActions(
    uri: Uri, name: String, mimeType: String, dirty: Boolean, saving: Boolean,
    saveRevision: Int, onSave: (Uri) -> Unit, onBack: () -> Unit, documentSwitch: DocumentSwitch? = null,
): DocumentActions {
    val context = LocalContext.current
    val scope = rememberCoroutineScope()
    var pending by rememberSaveable { mutableStateOf("") }
    var pickerOpen by rememberSaveable { mutableStateOf(false) }
    var askLeave by remember { mutableStateOf(false) }
    var askShare by remember { mutableStateOf(false) }
    var seenRevision by rememberSaveable { mutableIntStateOf(saveRevision) }
    var wasSaving by remember { mutableStateOf(false) }
    val share: () -> Unit = {
        scope.launch {
            try { context.startActivity(buildShareIntent(context, uri, name)) }
            catch (e: kotlinx.coroutines.CancellationException) { throw e }
            catch (e: Exception) { Toast.makeText(context, "Could not share: ${e.message}", Toast.LENGTH_LONG).show() }
        }
    }
    val launcher = rememberLauncherForActivityResult(ActivityResultContracts.CreateDocument(mimeType)) { output ->
        pickerOpen = false
        if (output != null) {
            if (pending == "switch") pending = ""
            onSave(output)
        } else {
            if (pending == "switch") documentSwitch?.cancel?.invoke()
            pending = ""
        }
    }
    fun save(action: String, suggestedName: String? = null) {
        if (!saving && !pickerOpen) {
            pickerOpen = true
            pending = action
            val extension = when {
                mimeType.contains("wordprocessingml") -> "docx"
                mimeType.contains("spreadsheetml") -> "xlsx"
                else -> "pdf"
            }
            launcher.launch(suggestedName ?: name.substringBeforeLast('.', name) + "_copy." + extension)
        }
    }
    LaunchedEffect(saveRevision, saving) {
        if (saveRevision != seenRevision) {
            seenRevision = saveRevision
            val action = pending
            pending = ""
            if (action == "share") share()
            if (action == "back") onBack()
        } else if (wasSaving && !saving) {
            pending = ""
        }
        wasSaving = saving
    }
    val back: () -> Unit = { if (!saving) { if (dirty) askLeave = true else onBack() } }
    BackHandler(enabled = dirty || saving, onBack = back)
    LaunchedEffect(documentSwitch?.token, dirty, saving, pending, pickerOpen) {
        if (documentSwitch != null && !dirty && !saving && !pickerOpen && pending.isEmpty()) documentSwitch.open()
    }
    val switching = documentSwitch != null && dirty && !saving && !pickerOpen && pending.isEmpty()
    if (switching || askLeave) {
        val keepEditing = {
            askLeave = false
            documentSwitch?.cancel?.invoke()
            Unit
        }
        AlertDialog(
            onDismissRequest = keepEditing, title = { Text("Save your changes?") },
            text = { Text("Save a copy to keep your changes and preserve the original.") },
            confirmButton = { TextButton(onClick = { askLeave = false; save(if (switching) "switch" else "back") }) {
                Text(if (switching) "Save and open" else "Save and close")
            } },
            dismissButton = { Row {
                TextButton(onClick = keepEditing) { Text("Keep editing") }
                TextButton(onClick = {
                    askLeave = false
                    if (switching) documentSwitch?.open?.invoke() else onBack()
                }) { Text("Discard") }
            } },
        )
    }
    if (askShare) AlertDialog(
        onDismissRequest = { askShare = false }, title = { Text("Share your changes") },
        text = { Text("Save an edited copy, then choose where to send it.") },
        confirmButton = { TextButton(onClick = { askShare = false; save("share") }) { Text("Save and share") } },
        dismissButton = { TextButton(onClick = { askShare = false }) { Text("Cancel") } },
    )
    return DocumentActions(
        saveCopy = { save("") },
        saveAs = { suggested -> save("", suggested) },
        share = { if (!saving) { if (dirty) askShare = true else share() } },
        back = back,
    )
}

@Composable
fun DocumentEditBar(dirty: Boolean, saving: Boolean, canUndo: Boolean, onUndo: () -> Unit, onSave: () -> Unit) {
    if (dirty || canUndo || saving) Surface(color = MaterialTheme.colorScheme.secondaryContainer) {
        Row(Modifier.fillMaxWidth().horizontalScroll(rememberScrollState()).padding(horizontal = 12.dp),
            verticalAlignment = androidx.compose.ui.Alignment.CenterVertically) {
            Text(if (saving) "Saving…" else if (dirty) "Unsaved changes" else "Saved copy", style = MaterialTheme.typography.labelMedium)
            Spacer(Modifier.width(12.dp))
            TextButton(onClick = onUndo, enabled = canUndo && !saving) { Text("Undo") }
            TextButton(onClick = onSave, enabled = !saving) { Text("Save a copy") }
        }
    }
}

@Composable
fun DocumentWarnings(warnings: List<String>) {
    var expanded by remember { mutableStateOf(false) }
    if (warnings.isEmpty()) return
    Surface(color = MaterialTheme.colorScheme.tertiaryContainer) {
        Row(Modifier.fillMaxWidth().padding(horizontal = 12.dp), verticalAlignment = androidx.compose.ui.Alignment.CenterVertically) {
            Text("Some content could not be loaded", Modifier.weight(1f), style = MaterialTheme.typography.bodySmall)
            TextButton(onClick = { expanded = true }) { Text("Details") }
        }
    }
    if (expanded) AlertDialog(onDismissRequest = { expanded = false },
        title = { Text("Document details") },
        text = { androidx.compose.foundation.lazy.LazyColumn { items(warnings.size) { Text(warnings[it], Modifier.padding(vertical = 4.dp)) } } },
        confirmButton = { TextButton(onClick = { expanded = false }) { Text("OK") } })
}
