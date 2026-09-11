package com.modocs.core.ui.components

import androidx.compose.runtime.*
import androidx.compose.foundation.lazy.LazyListState
import androidx.compose.ui.platform.LocalContext
import androidx.compose.ui.platform.LocalView
import com.modocs.core.common.AppPreferences
import com.modocs.core.common.ReaderPreferences
import kotlinx.coroutines.flow.distinctUntilChanged
import kotlinx.coroutines.flow.filter
import kotlinx.coroutines.flow.first

@Composable
fun rememberAppPreferences(): ReaderPreferences {
    val context = LocalContext.current
    return remember(context) { AppPreferences.get(context) }.state.collectAsState().value
}

@Composable
fun ReaderKeepAwake() {
    val enabled = rememberAppPreferences().keepAwake
    val view = LocalView.current
    DisposableEffect(view, enabled) {
        view.keepScreenOn = enabled
        onDispose { view.keepScreenOn = false }
    }
}

@Composable
fun RememberReadingPosition(document: String, list: LazyListState, ready: Boolean) {
    ReaderKeepAwake()
    val context = LocalContext.current
    val store = remember(context) { AppPreferences.get(context) }
    val enabled = rememberAppPreferences().rememberPosition
    LaunchedEffect(document, ready, enabled, list) {
        if (!ready || !enabled) return@LaunchedEffect
        val saved = store.position(document)
        val count = snapshotFlow { list.layoutInfo.totalItemsCount }.filter { it > 0 }.first()
        list.scrollToItem(saved.coerceIn(0, count - 1))
        snapshotFlow { list.firstVisibleItemIndex }.distinctUntilChanged().collect { store.savePosition(document, it) }
    }
}
