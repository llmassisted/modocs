package com.modocs.feature.docx

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
import com.modocs.core.ui.components.ZoomableContainer
import androidx.compose.foundation.layout.Arrangement
import androidx.compose.foundation.layout.Box
import androidx.compose.foundation.layout.Column
import androidx.compose.foundation.horizontalScroll
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
import androidx.compose.foundation.rememberScrollState
import androidx.compose.foundation.lazy.LazyListState
import androidx.compose.foundation.lazy.itemsIndexed
import androidx.compose.foundation.lazy.rememberLazyListState
import androidx.compose.foundation.text.KeyboardActions
import androidx.compose.foundation.text.KeyboardOptions
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.automirrored.filled.ArrowBack
import androidx.compose.material.icons.filled.Close
import androidx.compose.material.icons.filled.Description
import androidx.compose.material.icons.filled.Edit
import androidx.compose.material.icons.filled.EditOff
import androidx.compose.material.icons.filled.FormatBold
import androidx.compose.material.icons.filled.FormatItalic
import androidx.compose.material.icons.filled.FormatUnderlined
import androidx.compose.material.icons.filled.KeyboardArrowDown
import androidx.compose.material.icons.filled.KeyboardArrowUp
import androidx.compose.material.icons.filled.PictureAsPdf
import androidx.compose.material.icons.filled.Save
import androidx.compose.material.icons.filled.SaveAs
import androidx.compose.material.icons.filled.Search
import androidx.compose.material3.BottomAppBar
import androidx.compose.material3.CircularProgressIndicator
import androidx.compose.material3.ExperimentalMaterial3Api
import androidx.compose.material3.Icon
import androidx.compose.material3.IconButton
import androidx.compose.material3.IconButtonDefaults
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.OutlinedTextField
import androidx.compose.material3.OutlinedTextFieldDefaults
import androidx.compose.material3.Scaffold
import androidx.compose.material3.Surface
import androidx.compose.material3.Text
import androidx.compose.material3.TopAppBar
import androidx.compose.material3.SnackbarHost
import androidx.compose.material3.SnackbarHostState
import androidx.compose.material3.TopAppBarDefaults
import androidx.compose.runtime.Composable
import androidx.compose.runtime.LaunchedEffect
import androidx.compose.runtime.getValue
import androidx.compose.runtime.mutableIntStateOf
import androidx.compose.runtime.produceState
import androidx.compose.runtime.remember
import androidx.compose.runtime.setValue
import androidx.compose.runtime.snapshotFlow
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.focus.FocusRequester
import androidx.compose.ui.focus.focusRequester
import androidx.compose.ui.geometry.Offset
import androidx.compose.ui.geometry.Size
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.graphics.asImageBitmap
import androidx.compose.ui.graphics.drawscope.Stroke
import androidx.compose.ui.input.nestedscroll.nestedScroll
import androidx.compose.ui.layout.ContentScale
import androidx.compose.ui.layout.onSizeChanged
import androidx.compose.ui.platform.LocalConfiguration
import androidx.compose.ui.platform.LocalContext
import androidx.compose.ui.platform.LocalDensity
import androidx.compose.ui.platform.LocalSoftwareKeyboardController
import androidx.compose.ui.text.input.ImeAction
import androidx.compose.ui.text.style.TextOverflow
import androidx.compose.ui.unit.dp
import androidx.hilt.navigation.compose.hiltViewModel
import androidx.lifecycle.compose.collectAsStateWithLifecycle
import com.modocs.fonts.FontResolver
import kotlinx.coroutines.flow.distinctUntilChanged

// Highlight colors for search
private val HighlightYellow = Color(0x66FFEB3B)
private val HighlightOrange = Color(0x99FF9800)
private val HighlightBorderOrange = Color(0xFFFF9800)

@OptIn(ExperimentalMaterial3Api::class)
@Composable
fun DocxViewerScreen(
    uri: Uri,
    displayName: String?,
    onNavigateBack: () -> Unit,
    onExportPdf: ((Uri) -> Unit)? = null,
    viewModel: DocxViewerViewModel = hiltViewModel(),
) {
    val state by viewModel.state.collectAsStateWithLifecycle()
    val searchState by viewModel.searchState.collectAsStateWithLifecycle()
    val autoPageBreaks by viewModel.autoPageBreaks.collectAsStateWithLifecycle()
    val snackbarHostState = remember { SnackbarHostState() }

    // SAF launcher for "Save As PDF"
    val pdfSaveLauncher = rememberLauncherForActivityResult(
        contract = ActivityResultContracts.CreateDocument("application/pdf"),
    ) { outputUri: Uri? ->
        if (outputUri != null) {
            viewModel.exportToPdf(outputUri)
        }
    }

    LaunchedEffect(uri) {
        viewModel.loadDocx(uri, displayName)
    }

    // SAF launcher for "Save As"
    val saveAsLauncher = rememberLauncherForActivityResult(
        contract = ActivityResultContracts.CreateDocument("application/vnd.openxmlformats-officedocument.wordprocessingml.document"),
    ) { outputUri: Uri? ->
        if (outputUri != null) {
            viewModel.saveDocumentAs(outputUri)
        }
    }

    // Collect events (export/save success/failure)
    LaunchedEffect(Unit) {
        viewModel.events.collect { event ->
            when (event) {
                is DocxViewerViewModel.DocxEvent.ExportPdfReady ->
                    snackbarHostState.showSnackbar(event.message)
                is DocxViewerViewModel.DocxEvent.ExportPdfError ->
                    snackbarHostState.showSnackbar(event.message)
                is DocxViewerViewModel.DocxEvent.SaveSuccess ->
                    snackbarHostState.showSnackbar(event.message)
                is DocxViewerViewModel.DocxEvent.SaveError ->
                    snackbarHostState.showSnackbar(event.message)
            }
        }
    }

    val scrollBehavior = TopAppBarDefaults.pinnedScrollBehavior()
    val listState = rememberLazyListState()

    // Scroll to match when current search match changes
    LaunchedEffect(searchState.currentMatch, state.viewMode) {
        val match = searchState.currentMatch ?: return@LaunchedEffect
        if (state.viewMode == DocxViewMode.CANVAS && match.pageIndex >= 0) {
            listState.animateScrollToItem(match.pageIndex)
        } else {
            listState.animateScrollToItem(match.elementIndex)
        }
    }

    Scaffold(
        modifier = Modifier.nestedScroll(scrollBehavior.nestedScrollConnection),
        snackbarHost = { SnackbarHost(snackbarHostState) },
        topBar = {
            TopAppBar(
                title = {
                    Column {
                        Text(
                            text = state.fileName.ifEmpty { "Word Document" },
                            maxLines = 1,
                            overflow = TextOverflow.Ellipsis,
                            style = MaterialTheme.typography.titleMedium,
                        )
                        if (state.document != null) {
                            val subtitle = if (state.viewMode == DocxViewMode.CANVAS && state.pageCount > 0) {
                                "Page ${state.currentPage + 1} of ${state.pageCount}"
                            } else {
                                "${state.document!!.body.size} elements"
                            }
                            Text(
                                text = subtitle,
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
                    if (state.document != null) {
                        // Save button (visible when dirty)
                        if (state.isDirty) {
                            IconButton(onClick = { viewModel.saveDocument() }, enabled = !state.isSaving) {
                                if (state.isSaving) {
                                    CircularProgressIndicator(modifier = Modifier.size(20.dp), strokeWidth = 2.dp)
                                } else {
                                    Icon(Icons.Default.Save, contentDescription = "Save")
                                }
                            }
                            IconButton(onClick = {
                                val docxName = state.fileName
                                saveAsLauncher.launch(docxName)
                            }) {
                                Icon(Icons.Default.SaveAs, contentDescription = "Save As")
                            }
                        }

                        // Edit toggle
                        IconButton(onClick = { viewModel.toggleEditMode() }) {
                            Icon(
                                imageVector = if (state.isEditing) Icons.Default.EditOff else Icons.Default.Edit,
                                contentDescription = if (state.isEditing) "Stop editing" else "Edit document",
                            )
                        }

                        // Export to PDF button
                        IconButton(onClick = {
                            val pdfName = state.fileName
                                .removeSuffix(".docx").removeSuffix(".doc") + ".pdf"
                            pdfSaveLauncher.launch(pdfName)
                        }) {
                            Icon(Icons.Default.PictureAsPdf, contentDescription = "Export to PDF")
                        }
                    }

                    IconButton(onClick = { viewModel.toggleSearch() }) {
                        Icon(
                            imageVector = if (searchState.isSearchActive) Icons.Default.Close else Icons.Default.Search,
                            contentDescription = if (searchState.isSearchActive) "Close search" else "Search in document",
                        )
                    }
                },
                scrollBehavior = scrollBehavior,
            )
        },
        bottomBar = {
            if (state.isEditing && state.editingElementIndex >= 0) {
                val editingPara = state.document?.body?.getOrNull(state.editingElementIndex) as? DocxParagraph
                val runProps = editingPara?.runs?.firstOrNull()?.properties

                FormattingToolbar(
                    isBold = runProps?.bold == true,
                    isItalic = runProps?.italic == true,
                    isUnderline = runProps?.underline == true,
                    onBoldClick = { viewModel.toggleBold(state.editingElementIndex) },
                    onItalicClick = { viewModel.toggleItalic(state.editingElementIndex) },
                    onUnderlineClick = { viewModel.toggleUnderline(state.editingElementIndex) },
                )
            }
        },
    ) { innerPadding ->
        Column(modifier = Modifier.padding(innerPadding)) {
            AnimatedVisibility(
                visible = searchState.isSearchActive,
                enter = expandVertically(),
                exit = shrinkVertically(),
            ) {
                DocxSearchBar(
                    searchState = searchState,
                    onQueryChange = { viewModel.updateSearchQuery(it) },
                    onNext = { viewModel.nextMatch() },
                    onPrevious = { viewModel.previousMatch() },
                )
            }

            when {
                state.isLoading -> LoadingContent()
                state.errorMessage != null -> ErrorContent(message = state.errorMessage!!)
                state.document != null && state.viewMode == DocxViewMode.CANVAS ->
                    DocxPagesContent(
                        viewModel = viewModel,
                        state = state,
                        searchState = searchState,
                        listState = listState,
                    )
                state.document != null -> DocxContent(
                    document = state.document!!,
                    viewModel = viewModel,
                    searchState = searchState,
                    listState = listState,
                    isEditing = state.isEditing,
                    editingElementIndex = state.editingElementIndex,
                    autoPageBreaks = autoPageBreaks,
                )
            }
        }
    }
}

@Composable
private fun DocxSearchBar(
    searchState: DocxSearchState,
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
                    searchState.hasMatches -> {
                        val current = searchState.currentMatchIndex + 1
                        val total = searchState.totalMatches
                        val pageInfo = searchState.currentMatch?.pageIndex?.let { "(page ${it + 1})" } ?: ""
                        "$current of $total matches $pageInfo"
                    }
                    else -> "No matches found"
                }
                Text(
                    text = statusText,
                    style = MaterialTheme.typography.bodySmall,
                    color = if (searchState.hasMatches || searchState.isSearching) {
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

// ---------------------------------------------------------------
// CANVAS mode: Bitmap page rendering with zoom
// ---------------------------------------------------------------

@Composable
private fun DocxPagesContent(
    viewModel: DocxViewerViewModel,
    state: DocxViewerState,
    searchState: DocxSearchState,
    listState: LazyListState = rememberLazyListState(),
    modifier: Modifier = Modifier,
) {
    val renderer = viewModel.getPageRenderer() ?: return

    LaunchedEffect(listState) {
        snapshotFlow { listState.firstVisibleItemIndex }
            .distinctUntilChanged()
            .collect { viewModel.updateCurrentPage(it) }
    }

    ZoomableContainer(
        modifier = modifier.fillMaxSize(),
        maxScale = 4f,
        contentModifier = Modifier.fillMaxSize(),
    ) { _ ->
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
                DocxPage(
                    renderer = renderer,
                    pageIndex = pageIndex,
                    highlights = viewModel.getHighlightsForPage(pageIndex),
                    highlightVersion = highlightVersion,
                )
            }
        }
    }
}

@Composable
private fun DocxPage(
    renderer: DocxPageRenderer,
    pageIndex: Int,
    highlights: List<DocxPageHighlight>,
    highlightVersion: Pair<Int, Int>,
) {
    val configuration = LocalConfiguration.current
    val density = LocalDensity.current

    // Render at 2x display resolution for crisp text
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

        Box(
            modifier = Modifier
                .fillMaxWidth()
                .padding(horizontal = 8.dp)
                .onSizeChanged { composableWidth = it.width },
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
        }
    } else {
        // Loading placeholder
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

// ---------------------------------------------------------------
// COMPOSE mode: Editable element-based rendering
// ---------------------------------------------------------------

@Composable
private fun DocxContent(
    document: DocxDocument,
    viewModel: DocxViewerViewModel,
    searchState: DocxSearchState,
    listState: LazyListState,
    isEditing: Boolean,
    editingElementIndex: Int,
    autoPageBreaks: Set<Int>,
) {
    val context = LocalContext.current
    val fontResolver = remember { FontResolver(context) }
    val configuration = LocalConfiguration.current

    // Compute scale factor: maps page points to screen dp proportionally.
    // This ensures margins, fonts, and spacing in COMPOSE mode match the
    // CANVAS bitmap view where the full page is scaled to screen width.
    val pageSetup = document.pageSetup
    val screenWidthDp = configuration.screenWidthDp.toFloat()
    val pageScale = screenWidthDp / pageSetup.pageWidthPt

    // Layout-level zoom: scale all dimensions (including content width) by the
    // zoom factor so text wraps at the exact same points at every zoom level,
    // preserving the fixed PDF-like layout. No graphicsLayer means
    // BasicTextField cursor/handle positioning is always correct.
    val horizontalScrollState = rememberScrollState()

    ZoomableContainer(
        modifier = Modifier.fillMaxSize(),
        applyTransform = false,
        maxScale = 4f,
        contentModifier = Modifier.fillMaxSize(),
    ) { zoomScale ->
        val effectivePageScale = pageScale * zoomScale
        val contentWidth = (screenWidthDp * zoomScale).dp

        Box(
            modifier = Modifier
                .fillMaxSize()
                .horizontalScroll(horizontalScrollState),
        ) {
            LazyColumn(
                state = listState,
                modifier = Modifier
                    .width(contentWidth)
                    .background(Color.White),
                contentPadding = PaddingValues(
                    start = (pageSetup.marginLeftPt * effectivePageScale).dp,
                    end = (pageSetup.marginRightPt * effectivePageScale).dp,
                    top = (pageSetup.marginTopPt * effectivePageScale).dp,
                    bottom = (12 * zoomScale).dp,
                ),
                verticalArrangement = Arrangement.spacedBy(0.dp),
            ) {
                itemsIndexed(
                    items = document.body,
                    key = { index, _ -> index },
                ) { index, element ->
                    if (index in autoPageBreaks) {
                        AutoPageBreakIndicator(pageNumber = computePageNumber(index, autoPageBreaks))
                    }

                    val highlights = viewModel.getHighlightsForElement(index)
                    val highlightRanges = highlights.map { it.first }
                    val currentGlobalIndex = searchState.currentMatchIndex
                    val currentLocalIndex = highlights.indexOfFirst { it.second == currentGlobalIndex }

                    val isThisEditing = isEditing && editingElementIndex == index

                    DocxElementRenderer(
                        element = element,
                        document = document,
                        fontResolver = fontResolver,
                        pageScale = effectivePageScale,
                        searchHighlights = highlightRanges,
                        currentHighlightIndex = currentLocalIndex,
                        isEditing = isEditing,
                        isActivelyEditing = isThisEditing,
                        onTapToEdit = { viewModel.startEditingElement(index) },
                        onTextChanged = { newText -> viewModel.updateParagraphText(index, newText) },
                    )
                }
            }
        }
    }
}

private fun computePageNumber(elementIndex: Int, autoPageBreaks: Set<Int>): Int {
    return autoPageBreaks.count { it <= elementIndex } + 1
}

@Composable
private fun AutoPageBreakIndicator(pageNumber: Int) {
    Row(
        modifier = Modifier
            .fillMaxWidth()
            .padding(vertical = 20.dp),
        verticalAlignment = Alignment.CenterVertically,
        horizontalArrangement = Arrangement.Center,
    ) {
        androidx.compose.foundation.Canvas(
            modifier = Modifier.weight(1f).height(2.dp),
        ) {
            drawLine(
                color = androidx.compose.ui.graphics.Color.Black,
                start = androidx.compose.ui.geometry.Offset(0f, size.height / 2),
                end = androidx.compose.ui.geometry.Offset(size.width, size.height / 2),
                strokeWidth = 2f,
                pathEffect = androidx.compose.ui.graphics.PathEffect.dashPathEffect(
                    floatArrayOf(10f, 6f),
                    0f,
                ),
            )
        }
        Text(
            text = "  Page $pageNumber  ",
            style = MaterialTheme.typography.labelSmall,
            color = androidx.compose.ui.graphics.Color.Black,
            fontWeight = androidx.compose.ui.text.font.FontWeight.Bold,
        )
        androidx.compose.foundation.Canvas(
            modifier = Modifier.weight(1f).height(2.dp),
        ) {
            drawLine(
                color = androidx.compose.ui.graphics.Color.Black,
                start = androidx.compose.ui.geometry.Offset(0f, size.height / 2),
                end = androidx.compose.ui.geometry.Offset(size.width, size.height / 2),
                strokeWidth = 2f,
                pathEffect = androidx.compose.ui.graphics.PathEffect.dashPathEffect(
                    floatArrayOf(10f, 6f),
                    0f,
                ),
            )
        }
    }
}

@Composable
private fun FormattingToolbar(
    isBold: Boolean,
    isItalic: Boolean,
    isUnderline: Boolean,
    onBoldClick: () -> Unit,
    onItalicClick: () -> Unit,
    onUnderlineClick: () -> Unit,
) {
    BottomAppBar(
        tonalElevation = 4.dp,
    ) {
        Row(
            modifier = Modifier.fillMaxWidth().padding(horizontal = 8.dp),
            horizontalArrangement = Arrangement.Start,
            verticalAlignment = Alignment.CenterVertically,
        ) {
            val activeColors = IconButtonDefaults.iconButtonColors(
                containerColor = MaterialTheme.colorScheme.primaryContainer,
                contentColor = MaterialTheme.colorScheme.onPrimaryContainer,
            )
            val defaultColors = IconButtonDefaults.iconButtonColors()

            IconButton(
                onClick = onBoldClick,
                colors = if (isBold) activeColors else defaultColors,
            ) {
                Icon(Icons.Default.FormatBold, contentDescription = "Bold")
            }

            IconButton(
                onClick = onItalicClick,
                colors = if (isItalic) activeColors else defaultColors,
            ) {
                Icon(Icons.Default.FormatItalic, contentDescription = "Italic")
            }

            IconButton(
                onClick = onUnderlineClick,
                colors = if (isUnderline) activeColors else defaultColors,
            ) {
                Icon(Icons.Default.FormatUnderlined, contentDescription = "Underline")
            }
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
            Text("Opening Word document...")
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
                imageVector = Icons.Filled.Description,
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
