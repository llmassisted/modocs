package com.modocs.core.ui.components

import androidx.compose.foundation.clickable
import androidx.compose.foundation.lazy.LazyListState
import androidx.compose.foundation.text.KeyboardOptions
import androidx.compose.material3.*
import androidx.compose.runtime.*
import androidx.compose.runtime.saveable.rememberSaveable
import androidx.compose.ui.Modifier
import androidx.compose.ui.text.input.KeyboardType
import kotlinx.coroutines.launch

/** Page occupying the largest vertical portion of the viewport. */
fun LazyListState.mostVisiblePage(): Int {
    val layout = layoutInfo
    return layout.visibleItemsInfo.maxByOrNull { item ->
        minOf(layout.viewportEndOffset, item.offset + item.size) - maxOf(layout.viewportStartOffset, item.offset)
    }?.index ?: firstVisibleItemIndex
}

@Composable
fun ReaderPageNumber(currentPage: Int, pageCount: Int, list: LazyListState) {
    var showJump by rememberSaveable { mutableStateOf(false) }
    val scope = rememberCoroutineScope()
    Text(
        "Page ${currentPage + 1} of $pageCount",
        modifier = Modifier.clickable(onClickLabel = "Go to page") { showJump = true },
        style = MaterialTheme.typography.bodySmall,
        color = MaterialTheme.colorScheme.onSurfaceVariant,
    )
    if (showJump) {
        var input by rememberSaveable { mutableStateOf("") }
        val target = input.toIntOrNull()?.takeIf { it in 1..pageCount }
        AlertDialog(
            onDismissRequest = { showJump = false },
            title = { Text("Go to page") },
            text = { OutlinedTextField(
                value = input, onValueChange = { input = it }, singleLine = true,
                label = { Text("Page (1–$pageCount)") },
                keyboardOptions = KeyboardOptions(keyboardType = KeyboardType.Number),
                isError = input.isNotEmpty() && target == null,
            ) },
            confirmButton = { TextButton(enabled = target != null, onClick = {
                if (target != null) scope.launch { list.scrollToItem(target - 1) }
                showJump = false
            }) { Text("Go") } },
            dismissButton = { TextButton(onClick = { showJump = false }) { Text("Cancel") } },
        )
    }
}
