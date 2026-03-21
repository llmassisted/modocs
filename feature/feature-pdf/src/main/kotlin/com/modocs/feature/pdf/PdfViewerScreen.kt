package com.modocs.feature.pdf

import android.graphics.Bitmap
import android.net.Uri
import androidx.activity.compose.rememberLauncherForActivityResult
import androidx.activity.result.contract.ActivityResultContracts
import androidx.compose.animation.AnimatedVisibility
import androidx.compose.animation.expandVertically
import androidx.compose.animation.shrinkVertically
import androidx.compose.foundation.Canvas
import androidx.compose.foundation.Image
import androidx.compose.foundation.background
import androidx.compose.foundation.border
import androidx.compose.foundation.clickable
import androidx.compose.foundation.gestures.detectDragGestures
import androidx.compose.foundation.gestures.detectTapGestures
import com.modocs.core.ui.components.ZoomableContainer
import androidx.compose.foundation.layout.Arrangement
import androidx.compose.foundation.layout.Box
import androidx.compose.foundation.layout.Column
import androidx.compose.foundation.layout.PaddingValues
import androidx.compose.foundation.layout.Row
import androidx.compose.foundation.layout.Spacer
import androidx.compose.foundation.layout.fillMaxSize
import androidx.compose.foundation.layout.fillMaxWidth
import androidx.compose.foundation.layout.height
import androidx.compose.foundation.layout.padding
import androidx.compose.foundation.layout.size
import androidx.compose.foundation.layout.width
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.lazy.LazyListState
import androidx.compose.foundation.lazy.rememberLazyListState
import androidx.compose.foundation.text.KeyboardActions
import androidx.compose.foundation.text.KeyboardOptions
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.automirrored.filled.ArrowBack
import androidx.compose.material.icons.filled.CalendarToday
import androidx.compose.material.icons.filled.Check
import androidx.compose.material.icons.filled.CheckBox
import androidx.compose.material.icons.filled.Close
import androidx.compose.material.icons.filled.Draw
import androidx.compose.material.icons.filled.Edit
import androidx.compose.material.icons.filled.KeyboardArrowDown
import androidx.compose.material.icons.filled.KeyboardArrowUp
import androidx.compose.material.icons.filled.PictureAsPdf
import androidx.compose.material.icons.filled.Save
import androidx.compose.material.icons.filled.Search
import androidx.compose.material.icons.filled.TextFields
import androidx.compose.material.icons.automirrored.filled.Undo
import androidx.compose.material3.CircularProgressIndicator
import androidx.compose.material3.ExperimentalMaterial3Api
import androidx.compose.material3.FilterChip
import androidx.compose.material3.FilterChipDefaults
import androidx.compose.material3.Icon
import androidx.compose.material3.IconButton
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.OutlinedTextField
import androidx.compose.material3.OutlinedTextFieldDefaults
import androidx.compose.material3.Scaffold
import androidx.compose.material3.SnackbarHost
import androidx.compose.material3.SnackbarHostState
import androidx.compose.material3.Surface
import androidx.compose.material3.Text
import androidx.compose.material3.TopAppBar
import androidx.compose.material3.TopAppBarDefaults
import androidx.compose.runtime.Composable
import androidx.compose.runtime.LaunchedEffect
import androidx.compose.runtime.getValue
import androidx.compose.runtime.mutableIntStateOf
import androidx.compose.runtime.mutableStateOf
import androidx.compose.runtime.produceState
import androidx.compose.runtime.remember
import androidx.compose.runtime.rememberCoroutineScope
import androidx.compose.runtime.setValue
import androidx.compose.runtime.snapshotFlow
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.focus.FocusRequester
import androidx.compose.ui.focus.focusRequester
import androidx.compose.ui.geometry.Offset
import androidx.compose.ui.geometry.Size
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.graphics.Path
import androidx.compose.ui.graphics.StrokeCap
import androidx.compose.ui.graphics.StrokeJoin
import androidx.compose.ui.graphics.asImageBitmap
import androidx.compose.ui.graphics.drawscope.Stroke
import androidx.compose.ui.graphics.nativeCanvas
import androidx.compose.ui.graphics.vector.ImageVector
import androidx.compose.ui.input.nestedscroll.nestedScroll
import androidx.compose.ui.input.pointer.pointerInput
import androidx.compose.ui.layout.ContentScale
import androidx.compose.ui.layout.onSizeChanged
import androidx.compose.ui.platform.LocalConfiguration
import androidx.compose.ui.platform.LocalDensity
import androidx.compose.ui.platform.LocalSoftwareKeyboardController
import androidx.compose.ui.text.input.ImeAction
import androidx.compose.ui.text.style.TextOverflow
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.toSize
import androidx.hilt.navigation.compose.hiltViewModel
import androidx.lifecycle.compose.collectAsStateWithLifecycle
import kotlinx.coroutines.flow.distinctUntilChanged
import kotlinx.coroutines.launch

// Highlight colors
private val HighlightYellow = Color(0x66FFEB3B)
private val HighlightOrange = Color(0x99FF9800)
private val HighlightBorderOrange = Color(0xFFFF9800)

@OptIn(ExperimentalMaterial3Api::class)
@Composable
fun PdfViewerScreen(
    uri: Uri,
    displayName: String?,
    onNavigateBack: () -> Unit,
    viewModel: PdfViewerViewModel = hiltViewModel(),
) {
    val state by viewModel.state.collectAsStateWithLifecycle()
    val searchState by viewModel.searchState.collectAsStateWithLifecycle()
    val fillSignState by viewModel.fillSignState.collectAsStateWithLifecycle()

    val snackbarHostState = remember { SnackbarHostState() }
    val scope = rememberCoroutineScope()

    LaunchedEffect(uri) {
        viewModel.loadPdf(uri, displayName)
    }

    val scrollBehavior = TopAppBarDefaults.pinnedScrollBehavior()
    val listState = rememberLazyListState()

    // Scroll to page when current search match changes
    LaunchedEffect(searchState.currentMatch) {
        val match = searchState.currentMatch ?: return@LaunchedEffect
        listState.animateScrollToItem(match.pageIndex)
    }

    // Save-As launcher
    val saveAsLauncher = rememberLauncherForActivityResult(
        ActivityResultContracts.CreateDocument("application/pdf"),
    ) { outputUri ->
        if (outputUri != null) {
            viewModel.saveFilled(outputUri)
            scope.launch {
                snackbarHostState.showSnackbar("PDF saved successfully")
            }
        }
    }

    // Signature pad dialog
    if (fillSignState.showSignaturePad) {
        SignaturePadDialog(
            onDismiss = { viewModel.dismissSignaturePad() },
            onConfirm = { strokes -> viewModel.onSignatureDrawn(strokes) },
        )
    }

    Scaffold(
        modifier = Modifier.nestedScroll(scrollBehavior.nestedScrollConnection),
        topBar = {
            TopAppBar(
                title = {
                    Column {
                        Text(
                            text = state.fileName.ifEmpty { "PDF Viewer" },
                            maxLines = 1,
                            overflow = TextOverflow.Ellipsis,
                            style = MaterialTheme.typography.titleMedium,
                        )
                        if (state.pageCount > 0) {
                            Text(
                                text = "Page ${state.currentPage + 1} of ${state.pageCount}",
                                style = MaterialTheme.typography.bodySmall,
                                color = MaterialTheme.colorScheme.onSurfaceVariant,
                            )
                        }
                    }
                },
                navigationIcon = {
                    IconButton(onClick = onNavigateBack) {
                        Icon(
                            imageVector = Icons.AutoMirrored.Filled.ArrowBack,
                            contentDescription = "Back",
                        )
                    }
                },
                actions = {
                    if (!fillSignState.isActive) {
                        IconButton(onClick = { viewModel.toggleSearch() }) {
                            Icon(
                                imageVector = if (searchState.isSearchActive) Icons.Default.Close else Icons.Default.Search,
                                contentDescription = if (searchState.isSearchActive) "Close search" else "Search",
                            )
                        }
                    }
                    IconButton(onClick = { viewModel.toggleFillSign() }) {
                        Icon(
                            imageVector = if (fillSignState.isActive) Icons.Default.Close else Icons.Default.Edit,
                            contentDescription = if (fillSignState.isActive) "Exit fill & sign" else "Fill & Sign",
                            tint = if (fillSignState.isActive) MaterialTheme.colorScheme.primary else MaterialTheme.colorScheme.onSurface,
                        )
                    }
                },
                scrollBehavior = scrollBehavior,
            )
        },
        snackbarHost = { SnackbarHost(snackbarHostState) },
    ) { innerPadding ->
        Column(modifier = Modifier.padding(innerPadding)) {
            // Search bar
            AnimatedVisibility(
                visible = searchState.isSearchActive && !fillSignState.isActive,
                enter = expandVertically(),
                exit = shrinkVertically(),
            ) {
                SearchBar(
                    searchState = searchState,
                    onQueryChange = { viewModel.updateSearchQuery(it) },
                    onNext = { viewModel.nextMatch() },
                    onPrevious = { viewModel.previousMatch() },
                )
            }

            // Fill & Sign toolbar
            AnimatedVisibility(
                visible = fillSignState.isActive,
                enter = expandVertically(),
                exit = shrinkVertically(),
            ) {
                FillSignToolbar(
                    fillSignState = fillSignState,
                    onSelectTool = { viewModel.selectTool(it) },
                    onUndo = { viewModel.undoLastAnnotation() },
                    onSave = {
                        val name = state.fileName.removeSuffix(".pdf") + "_filled.pdf"
                        saveAsLauncher.launch(name)
                    },
                    onNewSignature = { viewModel.showSignaturePad() },
                    onSizeChange = { delta ->
                        fillSignState.selectedAnnotationId?.let { viewModel.changeAnnotationSize(it, delta) }
                    },
                    onDeleteSelected = {
                        fillSignState.selectedAnnotationId?.let {
                            viewModel.deleteAnnotation(it)
                            viewModel.selectAnnotation(null)
                        }
                    },
                    onDeselectAnnotation = { viewModel.selectAnnotation(null) },
                )
            }

            // Text input bar
            AnimatedVisibility(
                visible = fillSignState.editingAnnotationId != null,
                enter = expandVertically(),
                exit = shrinkVertically(),
            ) {
                TextInputOverlay(
                    text = fillSignState.editingText,
                    onTextChange = { viewModel.updateEditingText(it) },
                    onConfirm = { viewModel.confirmTextEdit() },
                    onCancel = { viewModel.cancelTextEdit() },
                )
            }

            when {
                state.isLoading -> LoadingContent()
                state.errorMessage != null -> ErrorContent(message = state.errorMessage!!)
                else -> PdfPagesContent(
                    viewModel = viewModel,
                    state = state,
                    searchState = searchState,
                    fillSignState = fillSignState,
                    listState = listState,
                )
            }
        }
    }
}

@Composable
private fun FillSignToolbar(
    fillSignState: FillSignState,
    onSelectTool: (FillSignTool) -> Unit,
    onUndo: () -> Unit,
    onSave: () -> Unit,
    onNewSignature: () -> Unit,
    onSizeChange: (delta: Float) -> Unit,
    onDeleteSelected: () -> Unit,
    onDeselectAnnotation: () -> Unit,
) {
    val selectedAnnotation = fillSignState.selectedAnnotation

    Surface(
        tonalElevation = 2.dp,
        modifier = Modifier.fillMaxWidth(),
    ) {
        Column(
            modifier = Modifier
                .fillMaxWidth()
                .padding(horizontal = 8.dp, vertical = 6.dp),
        ) {
            // Row 1: Tool chips (scrollable)
            Row(
                modifier = Modifier.fillMaxWidth(),
                horizontalArrangement = Arrangement.spacedBy(4.dp),
                verticalAlignment = Alignment.CenterVertically,
            ) {
                ToolChip(
                    label = "Text",
                    icon = Icons.Default.TextFields,
                    selected = fillSignState.selectedTool == FillSignTool.TEXT,
                    onClick = { onDeselectAnnotation(); onSelectTool(FillSignTool.TEXT) },
                )
                ToolChip(
                    label = "Sign",
                    icon = Icons.Default.Draw,
                    selected = fillSignState.selectedTool == FillSignTool.SIGNATURE,
                    onClick = {
                        onDeselectAnnotation()
                        if (fillSignState.selectedTool == FillSignTool.SIGNATURE) {
                            onNewSignature()
                        } else {
                            onSelectTool(FillSignTool.SIGNATURE)
                        }
                    },
                )
                ToolChip(
                    label = "\u2713",
                    icon = Icons.Default.CheckBox,
                    selected = fillSignState.selectedTool == FillSignTool.CHECKMARK,
                    onClick = { onDeselectAnnotation(); onSelectTool(FillSignTool.CHECKMARK) },
                )
                ToolChip(
                    label = "\u2717",
                    icon = Icons.Default.Close,
                    selected = fillSignState.selectedTool == FillSignTool.CROSS,
                    onClick = { onDeselectAnnotation(); onSelectTool(FillSignTool.CROSS) },
                )
                ToolChip(
                    label = "Date",
                    icon = Icons.Default.CalendarToday,
                    selected = fillSignState.selectedTool == FillSignTool.DATE,
                    onClick = { onDeselectAnnotation(); onSelectTool(FillSignTool.DATE) },
                )
            }

            // Row 2: Actions, size controls, or hint
            Row(
                modifier = Modifier.fillMaxWidth(),
                verticalAlignment = Alignment.CenterVertically,
            ) {
                IconButton(
                    onClick = onUndo,
                    enabled = fillSignState.hasAnnotations,
                    modifier = Modifier.size(36.dp),
                ) {
                    Icon(Icons.AutoMirrored.Filled.Undo, contentDescription = "Undo", modifier = Modifier.size(20.dp))
                }

                IconButton(
                    onClick = onSave,
                    enabled = fillSignState.hasAnnotations && !fillSignState.isSaving,
                    modifier = Modifier.size(36.dp),
                ) {
                    if (fillSignState.isSaving) {
                        CircularProgressIndicator(modifier = Modifier.size(18.dp), strokeWidth = 2.dp)
                    } else {
                        Icon(Icons.Default.Save, contentDescription = "Save", modifier = Modifier.size(20.dp))
                    }
                }

                if (selectedAnnotation != null) {
                    Spacer(modifier = Modifier.width(4.dp))

                    // Size controls for selected annotation
                    Surface(
                        shape = MaterialTheme.shapes.small,
                        tonalElevation = 4.dp,
                    ) {
                        Row(verticalAlignment = Alignment.CenterVertically) {
                            IconButton(
                                onClick = { onSizeChange(-2f) },
                                modifier = Modifier.size(32.dp),
                            ) {
                                Text("A", style = MaterialTheme.typography.labelSmall)
                            }

                            val sizeLabel = when (selectedAnnotation) {
                                is TextAnnotation -> "${selectedAnnotation.fontSizeSp.toInt()}"
                                is DateAnnotation -> "${selectedAnnotation.fontSizeSp.toInt()}"
                                is CheckmarkAnnotation -> "${selectedAnnotation.sizeSp.toInt()}"
                                is SignatureAnnotation -> "${(selectedAnnotation.width * 100).toInt()}%"
                            }
                            Text(
                                text = sizeLabel,
                                style = MaterialTheme.typography.labelSmall,
                                modifier = Modifier.padding(horizontal = 2.dp),
                            )

                            IconButton(
                                onClick = { onSizeChange(2f) },
                                modifier = Modifier.size(32.dp),
                            ) {
                                Text("A", style = MaterialTheme.typography.labelMedium)
                            }
                        }
                    }

                    Spacer(modifier = Modifier.width(4.dp))

                    // Delete selected annotation
                    IconButton(
                        onClick = onDeleteSelected,
                        modifier = Modifier.size(36.dp),
                    ) {
                        Icon(
                            Icons.Default.Close,
                            contentDescription = "Delete",
                            modifier = Modifier.size(20.dp),
                            tint = MaterialTheme.colorScheme.error,
                        )
                    }

                    Spacer(modifier = Modifier.weight(1f))
                    Text(
                        text = "Drag to move",
                        style = MaterialTheme.typography.bodySmall,
                        color = MaterialTheme.colorScheme.primary,
                    )
                } else {
                    Spacer(modifier = Modifier.width(8.dp))

                    if (fillSignState.selectedTool == FillSignTool.SIGNATURE && fillSignState.pendingSignature != null) {
                        Text(
                            text = "Tap to place signature",
                            style = MaterialTheme.typography.bodySmall,
                            color = MaterialTheme.colorScheme.primary,
                        )
                    } else if (fillSignState.selectedTool == FillSignTool.TEXT) {
                        Text(
                            text = "Tap to add text",
                            style = MaterialTheme.typography.bodySmall,
                            color = MaterialTheme.colorScheme.onSurfaceVariant,
                        )
                    }
                }
            }
        }
    }
}

@Composable
private fun ToolChip(
    label: String,
    icon: ImageVector,
    selected: Boolean,
    onClick: () -> Unit,
) {
    FilterChip(
        selected = selected,
        onClick = onClick,
        label = { Text(label, style = MaterialTheme.typography.labelSmall) },
        leadingIcon = {
            Icon(icon, contentDescription = null, modifier = Modifier.size(16.dp))
        },
        colors = FilterChipDefaults.filterChipColors(
            selectedContainerColor = MaterialTheme.colorScheme.primaryContainer,
        ),
    )
}

@Composable
private fun TextInputOverlay(
    text: String,
    onTextChange: (String) -> Unit,
    onConfirm: () -> Unit,
    onCancel: () -> Unit,
) {
    val focusRequester = remember { FocusRequester() }

    LaunchedEffect(Unit) {
        focusRequester.requestFocus()
    }

    Surface(
        modifier = Modifier.fillMaxWidth(),
        tonalElevation = 4.dp,
    ) {
        Row(
            modifier = Modifier
                .fillMaxWidth()
                .padding(horizontal = 12.dp, vertical = 8.dp),
            verticalAlignment = Alignment.CenterVertically,
        ) {
            OutlinedTextField(
                value = text,
                onValueChange = onTextChange,
                modifier = Modifier
                    .weight(1f)
                    .focusRequester(focusRequester),
                placeholder = { Text("Type text here...") },
                singleLine = true,
                keyboardOptions = KeyboardOptions(imeAction = ImeAction.Done),
                keyboardActions = KeyboardActions(onDone = { onConfirm() }),
                colors = OutlinedTextFieldDefaults.colors(
                    focusedContainerColor = MaterialTheme.colorScheme.surface,
                    unfocusedContainerColor = MaterialTheme.colorScheme.surface,
                ),
            )
            IconButton(onClick = onConfirm) {
                Icon(Icons.Default.Check, contentDescription = "Confirm", tint = MaterialTheme.colorScheme.primary)
            }
            IconButton(onClick = onCancel) {
                Icon(Icons.Default.Close, contentDescription = "Cancel")
            }
        }
    }
}

@Composable
private fun SearchBar(
    searchState: SearchState,
    onQueryChange: (String) -> Unit,
    onNext: () -> Unit,
    onPrevious: () -> Unit,
) {
    val focusRequester = remember { FocusRequester() }
    val keyboardController = LocalSoftwareKeyboardController.current

    LaunchedEffect(Unit) {
        focusRequester.requestFocus()
    }

    Surface(
        tonalElevation = 2.dp,
        modifier = Modifier.fillMaxWidth(),
    ) {
        Column(
            modifier = Modifier
                .fillMaxWidth()
                .padding(horizontal = 12.dp, vertical = 8.dp),
        ) {
            Row(
                verticalAlignment = Alignment.CenterVertically,
                modifier = Modifier.fillMaxWidth(),
            ) {
                OutlinedTextField(
                    value = searchState.query,
                    onValueChange = onQueryChange,
                    modifier = Modifier
                        .weight(1f)
                        .focusRequester(focusRequester),
                    placeholder = { Text("Search in document...") },
                    singleLine = true,
                    leadingIcon = {
                        Icon(
                            Icons.Default.Search,
                            contentDescription = null,
                            modifier = Modifier.size(20.dp),
                        )
                    },
                    trailingIcon = {
                        if (searchState.query.isNotEmpty()) {
                            IconButton(onClick = { onQueryChange("") }) {
                                Icon(
                                    Icons.Default.Close,
                                    contentDescription = "Clear",
                                    modifier = Modifier.size(20.dp),
                                )
                            }
                        }
                    },
                    keyboardOptions = KeyboardOptions(imeAction = ImeAction.Search),
                    keyboardActions = KeyboardActions(
                        onSearch = {
                            keyboardController?.hide()
                            if (searchState.hasMatches) onNext()
                        },
                    ),
                    colors = OutlinedTextFieldDefaults.colors(
                        focusedContainerColor = MaterialTheme.colorScheme.surface,
                        unfocusedContainerColor = MaterialTheme.colorScheme.surface,
                    ),
                )

                Spacer(modifier = Modifier.width(4.dp))

                IconButton(
                    onClick = onPrevious,
                    enabled = searchState.hasMatches,
                ) {
                    Icon(Icons.Default.KeyboardArrowUp, contentDescription = "Previous match")
                }

                IconButton(
                    onClick = onNext,
                    enabled = searchState.hasMatches,
                ) {
                    Icon(Icons.Default.KeyboardArrowDown, contentDescription = "Next match")
                }
            }

            if (searchState.query.length >= 2) {
                val statusText = when {
                    searchState.isSearching -> "Searching..."
                    !searchState.textExtracted -> "Extracting text..."
                    searchState.hasMatches -> {
                        val current = searchState.currentMatchIndex + 1
                        val total = searchState.totalMatches
                        val pageNum = (searchState.currentMatch?.pageIndex ?: 0) + 1
                        "$current of $total matches (page $pageNum)"
                    }
                    else -> "No matches found"
                }
                Text(
                    text = statusText,
                    style = MaterialTheme.typography.bodySmall,
                    color = if (searchState.hasMatches || searchState.isSearching || !searchState.textExtracted) {
                        MaterialTheme.colorScheme.onSurfaceVariant
                    } else {
                        MaterialTheme.colorScheme.error
                    },
                    modifier = Modifier.padding(start = 4.dp, top = 4.dp),
                )
            }
        }
    }
}

@Composable
private fun PdfPagesContent(
    viewModel: PdfViewerViewModel,
    state: PdfViewerState,
    searchState: SearchState,
    fillSignState: FillSignState,
    listState: LazyListState = rememberLazyListState(),
    modifier: Modifier = Modifier,
) {
    val renderer = viewModel.getRenderer() ?: return

    LaunchedEffect(listState) {
        snapshotFlow { listState.firstVisibleItemIndex }
            .distinctUntilChanged()
            .collect { viewModel.updateCurrentPage(it) }
    }

    ZoomableContainer(
        modifier = modifier.fillMaxSize(),
        enabled = true,
        maxScale = 4f,
        contentModifier = Modifier.fillMaxSize(),
    ) {
        val highlightVersion = remember(searchState.currentMatchIndex, searchState.matches.size) {
            searchState.currentMatchIndex to searchState.matches.size
        }

        LazyColumn(
            state = listState,
            modifier = Modifier.fillMaxSize(),
            contentPadding = PaddingValues(vertical = 8.dp),
            verticalArrangement = Arrangement.spacedBy(8.dp),
        ) {
            items(state.pageCount) { pageIndex ->
                val pageAnnotations = fillSignState.annotations.filter { it.pageIndex == pageIndex }
                PdfPage(
                    renderer = renderer,
                    pageIndex = pageIndex,
                    highlights = viewModel.getHighlightsForPage(pageIndex),
                    highlightVersion = highlightVersion,
                    fillSignActive = fillSignState.isActive,
                    annotations = pageAnnotations,
                    selectedAnnotationId = fillSignState.selectedAnnotationId,
                    onPageTap = { normX, normY, widthPx, heightPx ->
                        viewModel.onPageTap(pageIndex, normX, normY, widthPx, heightPx)
                    },
                    onAnnotationDrag = { id, normX, normY ->
                        viewModel.moveAnnotation(id, normX, normY)
                    },
                )
            }
        }
    }
}

@Composable
private fun PdfPage(
    renderer: PdfRendererWrapper,
    pageIndex: Int,
    highlights: List<PageHighlight>,
    highlightVersion: Pair<Int, Int>,
    fillSignActive: Boolean,
    annotations: List<PdfAnnotation>,
    selectedAnnotationId: String?,
    onPageTap: (normalizedX: Float, normalizedY: Float, widthPx: Float, heightPx: Float) -> Unit,
    onAnnotationDrag: (id: String, newNormX: Float, newNormY: Float) -> Unit,
) {
    val configuration = LocalConfiguration.current
    val density = LocalDensity.current

    val displayWidth = with(density) {
        (configuration.screenWidthDp.dp - 16.dp).toPx().toInt()
    }
    val renderWidth = (displayWidth * 2).coerceAtMost(3000)

    val bitmap by produceState<Bitmap?>(null, pageIndex, renderWidth) {
        value = renderer.renderPage(pageIndex, renderWidth)
    }

    val bmp = bitmap
    if (bmp != null) {
        var composableWidth by remember { mutableIntStateOf(0) }
        var composableHeight by remember { mutableIntStateOf(0) }

        Box(
            modifier = Modifier
                .fillMaxWidth()
                .padding(horizontal = 8.dp)
                .onSizeChanged {
                    composableWidth = it.width
                    composableHeight = it.height
                }
                .then(
                    if (fillSignActive) {
                        Modifier.pointerInput(selectedAnnotationId) {
                            detectTapGestures { offset ->
                                if (composableWidth > 0) {
                                    val aspectRatio = bmp.height.toFloat() / bmp.width.toFloat()
                                    val imageHeight = composableWidth * aspectRatio
                                    val normX = (offset.x / composableWidth).coerceIn(0f, 1f)
                                    val normY = (offset.y / imageHeight).coerceIn(0f, 1f)
                                    onPageTap(normX, normY, composableWidth.toFloat(), imageHeight)
                                }
                            }
                        }
                    } else Modifier
                )
                .then(
                    if (fillSignActive && selectedAnnotationId != null &&
                        annotations.any { it.id == selectedAnnotationId }) {
                        Modifier.pointerInput(selectedAnnotationId) {
                            detectDragGestures { change, _ ->
                                change.consume()
                                if (composableWidth > 0) {
                                    val aspectRatio = bmp.height.toFloat() / bmp.width.toFloat()
                                    val imageHeight = composableWidth * aspectRatio
                                    val normX = (change.position.x / composableWidth).coerceIn(0f, 1f)
                                    val normY = (change.position.y / imageHeight).coerceIn(0f, 1f)
                                    onAnnotationDrag(selectedAnnotationId, normX, normY)
                                }
                            }
                        }
                    } else Modifier
                ),
        ) {
            Image(
                bitmap = bmp.asImageBitmap(),
                contentDescription = "Page ${pageIndex + 1}",
                modifier = Modifier
                    .fillMaxWidth()
                    .background(Color.White),
                contentScale = ContentScale.FillWidth,
            )

            // Search highlight overlays
            if (highlights.isNotEmpty() && composableWidth > 0) {
                val aspectRatio = bmp.height.toFloat() / bmp.width.toFloat()
                val imageHeight = composableWidth * aspectRatio

                Canvas(
                    modifier = Modifier
                        .fillMaxWidth()
                        .height(with(density) { imageHeight.toDp() }),
                ) {
                    for (highlight in highlights) {
                        val rect = highlight.rect
                        val left = rect.left * size.width
                        val top = rect.top * size.height
                        val width = (rect.right - rect.left) * size.width
                        val height = (rect.bottom - rect.top) * size.height

                        drawRect(
                            color = if (highlight.isCurrent) HighlightOrange else HighlightYellow,
                            topLeft = Offset(left, top),
                            size = Size(width, height),
                        )

                        if (highlight.isCurrent) {
                            drawRect(
                                color = HighlightBorderOrange,
                                topLeft = Offset(left, top),
                                size = Size(width, height),
                                style = Stroke(width = 2.dp.toPx()),
                            )
                        }
                    }
                }
            }

            // Annotation overlays
            if (annotations.isNotEmpty() && composableWidth > 0) {
                val aspectRatio = bmp.height.toFloat() / bmp.width.toFloat()
                val imageHeight = composableWidth * aspectRatio

                Canvas(
                    modifier = Modifier
                        .fillMaxWidth()
                        .height(with(density) { imageHeight.toDp() }),
                ) {
                    val w = size.width
                    val h = size.height
                    // Use page-relative scaling so preview matches saved output exactly
                    val scaleFactor = w / 595f
                    val selectionColor = Color(0xFF1976D2)
                    val selectionStroke = Stroke(width = 2.dp.toPx(), pathEffect = androidx.compose.ui.graphics.PathEffect.dashPathEffect(floatArrayOf(8f, 4f)))

                    for (annotation in annotations) {
                        val px = annotation.x * w
                        val py = annotation.y * h
                        val isSelected = annotation.id == selectedAnnotationId

                        when (annotation) {
                            is TextAnnotation -> {
                                if (annotation.text.isNotEmpty()) {
                                    val paint = android.graphics.Paint(android.graphics.Paint.ANTI_ALIAS_FLAG).apply {
                                        color = android.graphics.Color.BLACK
                                        textSize = annotation.fontSizeSp * scaleFactor
                                    }
                                    drawContext.canvas.nativeCanvas.drawText(
                                        annotation.text, px, py + paint.textSize, paint,
                                    )
                                    if (isSelected) {
                                        val textW = paint.measureText(annotation.text)
                                        drawRect(selectionColor, Offset(px - 2, py), Size(textW + 4, paint.textSize * 1.3f), style = selectionStroke)
                                    }
                                }
                            }

                            is SignatureAnnotation -> {
                                val sigW = annotation.width * w
                                val sigH = annotation.height * h
                                val strokeStyle = Stroke(
                                    width = 2.5f * scaleFactor,
                                    cap = StrokeCap.Round,
                                    join = StrokeJoin.Round,
                                )
                                for (stroke in annotation.strokes) {
                                    if (stroke.size < 2) continue
                                    val path = Path().apply {
                                        moveTo(
                                            px + stroke.first().x * sigW,
                                            py + stroke.first().y * sigH,
                                        )
                                        for (i in 1 until stroke.size) {
                                            lineTo(
                                                px + stroke[i].x * sigW,
                                                py + stroke[i].y * sigH,
                                            )
                                        }
                                    }
                                    drawPath(path, Color.Black, style = strokeStyle)
                                }
                                if (isSelected) {
                                    drawRect(selectionColor, Offset(px, py), Size(sigW, sigH), style = selectionStroke)
                                }
                            }

                            is CheckmarkAnnotation -> {
                                val sz = annotation.sizeSp * scaleFactor
                                val strokeStyle = Stroke(
                                    width = 2.5f * scaleFactor,
                                    cap = StrokeCap.Round,
                                    join = StrokeJoin.Round,
                                )
                                if (annotation.isCheck) {
                                    val path = Path().apply {
                                        moveTo(px, py + sz * 0.5f)
                                        lineTo(px + sz * 0.35f, py + sz * 0.85f)
                                        lineTo(px + sz, py + sz * 0.1f)
                                    }
                                    drawPath(path, Color.Black, style = strokeStyle)
                                } else {
                                    drawLine(
                                        Color.Black,
                                        Offset(px, py),
                                        Offset(px + sz, py + sz),
                                        strokeWidth = 2.5f * scaleFactor,
                                    )
                                    drawLine(
                                        Color.Black,
                                        Offset(px + sz, py),
                                        Offset(px, py + sz),
                                        strokeWidth = 2.5f * scaleFactor,
                                    )
                                }
                                if (isSelected) {
                                    drawRect(selectionColor, Offset(px - 2, py - 2), Size(sz + 4, sz + 4), style = selectionStroke)
                                }
                            }

                            is DateAnnotation -> {
                                val paint = android.graphics.Paint(android.graphics.Paint.ANTI_ALIAS_FLAG).apply {
                                    color = android.graphics.Color.BLACK
                                    textSize = annotation.fontSizeSp * scaleFactor
                                }
                                drawContext.canvas.nativeCanvas.drawText(
                                    annotation.dateText, px, py + paint.textSize, paint,
                                )
                                if (isSelected) {
                                    val textW = paint.measureText(annotation.dateText)
                                    drawRect(selectionColor, Offset(px - 2, py), Size(textW + 4, paint.textSize * 1.3f), style = selectionStroke)
                                }
                            }
                        }
                    }
                }
            }
        }
    } else {
        Box(
            modifier = Modifier
                .fillMaxWidth()
                .height(400.dp)
                .padding(horizontal = 8.dp)
                .background(Color.White),
            contentAlignment = Alignment.Center,
        ) {
            CircularProgressIndicator(modifier = Modifier.size(32.dp))
        }
    }
}

@Composable
private fun LoadingContent(modifier: Modifier = Modifier) {
    Box(
        modifier = modifier.fillMaxSize(),
        contentAlignment = Alignment.Center,
    ) {
        Column(horizontalAlignment = Alignment.CenterHorizontally) {
            CircularProgressIndicator()
            Spacer(modifier = Modifier.height(16.dp))
            Text("Opening PDF...")
        }
    }
}

@Composable
private fun ErrorContent(
    message: String,
    modifier: Modifier = Modifier,
) {
    Box(
        modifier = modifier.fillMaxSize(),
        contentAlignment = Alignment.Center,
    ) {
        Column(horizontalAlignment = Alignment.CenterHorizontally) {
            Icon(
                imageVector = Icons.Filled.PictureAsPdf,
                contentDescription = null,
                modifier = Modifier.size(64.dp),
                tint = MaterialTheme.colorScheme.error,
            )
            Spacer(modifier = Modifier.height(16.dp))
            Text(
                text = message,
                style = MaterialTheme.typography.bodyLarge,
                color = MaterialTheme.colorScheme.error,
            )
        }
    }
}
