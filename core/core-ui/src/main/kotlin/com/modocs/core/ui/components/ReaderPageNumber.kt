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
private fun LazyListState.mostVisiblePage(): Int {
    val layout = layoutInfo
    return layout.visibleItemsInfo.maxByOrNull { item ->
        minOf(layout.viewportEndOffset, item.offset + item.size) - maxOf(layout.viewportStartOffset, item.offset)
    }?.index ?: firstVisibleItemIndex
}

/** A jump selects its page even when the end of the document constrains its placement. */
class ReaderPagePosition(private val list: LazyListState) {
    private var selected by mutableStateOf<Int?>(null)

    val currentPage by derivedStateOf {
        selected?.takeIf { page -> list.layoutInfo.visibleItemsInfo.any { it.index == page } }
            ?: list.mostVisiblePage()
    }

    suspend fun jumpTo(page: Int) {
        list.scrollToItem(page)
        selected = page
    }

    suspend fun followLayout() {
        // Replacing a loading placeholder with a bitmap can move the target again after
        // scrollToItem returns. Correct that reflow until the reader starts scrolling.
        snapshotFlow {
            val layout = list.layoutInfo
            val displaced = selected?.takeIf { page ->
                page < layout.totalItemsCount && layout.visibleItemsInfo.none { it.index == page }
            }
            list.isScrollInProgress to displaced
        }.collect { (scrolling, displaced) ->
            if (scrolling) selected = null
            else if (displaced != null) list.requestScrollToItem(displaced)
        }
    }
}

@Composable
fun rememberReaderPagePosition(list: LazyListState): ReaderPagePosition {
    val position = remember(list) { ReaderPagePosition(list) }
    LaunchedEffect(position) { position.followLayout() }
    return position
}

@Composable
fun ReaderPageNumber(position: ReaderPagePosition, pageCount: Int) {
    var showJump by rememberSaveable { mutableStateOf(false) }
    val scope = rememberCoroutineScope()
    Text(
        "Page ${position.currentPage + 1} of $pageCount",
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
                if (target != null) scope.launch { position.jumpTo(target - 1) }
                showJump = false
            }) { Text("Go") } },
            dismissButton = { TextButton(onClick = { showJump = false }) { Text("Cancel") } },
        )
    }
}
