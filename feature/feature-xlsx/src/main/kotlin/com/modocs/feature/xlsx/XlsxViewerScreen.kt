package com.modocs.feature.xlsx

import android.app.Activity
import android.net.Uri
import androidx.activity.compose.rememberLauncherForActivityResult
import androidx.activity.result.contract.ActivityResultContracts
import androidx.compose.animation.AnimatedVisibility
import androidx.compose.animation.expandVertically
import androidx.compose.animation.shrinkVertically
import androidx.compose.foundation.background
import androidx.compose.foundation.border
import androidx.compose.foundation.clickable
import androidx.compose.foundation.gestures.detectTransformGestures
import androidx.compose.foundation.horizontalScroll
import androidx.compose.foundation.layout.Box
import androidx.compose.foundation.layout.Column
import androidx.compose.foundation.layout.IntrinsicSize
import androidx.compose.foundation.layout.Row
import androidx.compose.foundation.layout.fillMaxHeight
import androidx.compose.foundation.layout.fillMaxSize
import androidx.compose.foundation.layout.fillMaxWidth
import androidx.compose.foundation.layout.height
import androidx.compose.foundation.layout.padding
import androidx.compose.foundation.layout.width
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.lazy.rememberLazyListState
import androidx.compose.foundation.rememberScrollState
import androidx.compose.foundation.text.KeyboardActions
import androidx.compose.foundation.text.KeyboardOptions
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.automirrored.filled.ArrowBack
import androidx.compose.material.icons.filled.Check
import androidx.compose.material.icons.filled.Close
import androidx.compose.material.icons.filled.Edit
import androidx.compose.material.icons.filled.KeyboardArrowDown
import androidx.compose.material.icons.filled.KeyboardArrowUp
import androidx.compose.material.icons.filled.Search
import androidx.compose.material.icons.filled.ZoomIn
import androidx.compose.material.icons.filled.ZoomOut
import androidx.compose.material3.ExperimentalMaterial3Api
import androidx.compose.material3.Icon
import androidx.compose.material3.IconButton
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.OutlinedTextField
import androidx.compose.material3.OutlinedTextFieldDefaults
import androidx.compose.material3.Scaffold
import androidx.compose.material3.ScrollableTabRow
import androidx.compose.material3.SnackbarHost
import androidx.compose.material3.SnackbarHostState
import androidx.compose.material3.Tab
import androidx.compose.material3.Text
import androidx.compose.material3.TextField
import androidx.compose.material3.TextFieldDefaults
import androidx.compose.material3.TopAppBar
import androidx.compose.material3.TopAppBarDefaults
import androidx.compose.runtime.Composable
import androidx.compose.runtime.LaunchedEffect
import androidx.compose.runtime.getValue
import androidx.compose.runtime.mutableFloatStateOf
import androidx.compose.runtime.mutableStateOf
import androidx.compose.runtime.remember
import androidx.compose.runtime.setValue
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.focus.FocusRequester
import androidx.compose.ui.focus.focusRequester
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.input.nestedscroll.nestedScroll
import androidx.compose.ui.input.pointer.pointerInput
import androidx.compose.ui.platform.LocalSoftwareKeyboardController
import androidx.compose.ui.text.font.FontStyle
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.text.input.ImeAction
import androidx.compose.ui.text.style.TextDecoration
import androidx.compose.ui.text.style.TextOverflow
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp
import androidx.hilt.navigation.compose.hiltViewModel
import androidx.lifecycle.compose.collectAsStateWithLifecycle
import com.modocs.core.ui.components.ErrorMessage
import com.modocs.core.ui.components.LoadingIndicator
import com.modocs.core.ui.components.*
import androidx.compose.material.icons.filled.Share
import androidx.compose.material3.TextButton

private val HighlightYellow = Color(0x66FFEB3B)
private val HighlightOrange = Color(0x99FF9800)
private val EditingCellBorder = Color(0xFF1976D2)

// Fixed spreadsheet colors — never overridden by dark/light theme
private val SheetBackground = Color.White
private val HeaderBackground = Color(0xFFF2F2F2)
private val HeaderTextColor = Color(0xFF333333)
private val CellBorderColor = Color(0xFFD4D4D4)
private val DefaultCellTextColor = Color(0xFF000000)

@OptIn(ExperimentalMaterial3Api::class)
@Composable
fun XlsxViewerScreen(
    uri: Uri,
    displayName: String? = null,
    onNavigateBack: () -> Unit,
    viewModel: XlsxViewerViewModel = hiltViewModel(),
) {
    val state by viewModel.state.collectAsStateWithLifecycle()
    val searchState by viewModel.searchState.collectAsStateWithLifecycle()
    val preferences = com.modocs.core.ui.components.rememberAppPreferences()
    var zoomScale by remember { mutableFloatStateOf(preferences.spreadsheetZoom) }
    val snackbarHostState = remember { SnackbarHostState() }
    val documentActions = rememberDocumentActions(state.activeUri ?: uri, state.fileName.ifEmpty { "document.xlsx" },
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", state.isDirty || state.draftDirty, state.isSaving, state.saveRevision,
        viewModel::saveDocumentAs, onNavigateBack)


    // Collect events for snackbar
    LaunchedEffect(Unit) {
        viewModel.events.collect { event ->
            when (event) {
                is XlsxViewerViewModel.XlsxEvent.SaveSuccess -> snackbarHostState.showSnackbar(event.message)
                is XlsxViewerViewModel.XlsxEvent.SaveError -> snackbarHostState.showSnackbar(event.message)
                is XlsxViewerViewModel.XlsxEvent.Error -> snackbarHostState.showSnackbar(event.message)
            }
        }
    }

    LaunchedEffect(uri) {
        viewModel.loadXlsx(uri, displayName)
    }

    // Password dialog
    if (state.isPasswordRequired) {
        com.modocs.core.ui.components.PasswordDialog(
            onPasswordSubmit = { password -> viewModel.submitPassword(password) },
            onDismiss = onNavigateBack,
            isLoading = state.isLoading,
            errorMessage = state.passwordError,
        )
    }

    // Save-As launcher
    val saveAsLauncher = rememberLauncherForActivityResult(
        contract = ActivityResultContracts.CreateDocument("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
    ) { outputUri ->
        if (outputUri != null) {
            viewModel.saveDocumentAs(outputUri)
        }
    }

    val scrollBehavior = TopAppBarDefaults.pinnedScrollBehavior()

    Scaffold(
        modifier = Modifier.nestedScroll(scrollBehavior.nestedScrollConnection),
        snackbarHost = { SnackbarHost(snackbarHostState) },
        topBar = {
            Column {
                TopAppBar(
                    title = {
                        Text(
                            text = state.fileName,
                            maxLines = 1,
                            overflow = TextOverflow.Ellipsis,
                        )
                    },
                    navigationIcon = {
                        IconButton(onClick = documentActions.back) {
                            Icon(Icons.AutoMirrored.Filled.ArrowBack, contentDescription = "Back")
                        }
                    },
                    actions = {
                        IconButton(onClick = { viewModel.toggleSearch() }) {
                            Icon(Icons.Filled.Search, contentDescription = "Search")
                        }
                        // Edit toggle
                        if (state.document != null) {
                            IconButton(onClick = { viewModel.toggleEditMode() }, enabled = !state.isSaving) {
                                Icon(
                                    if (state.isEditing) Icons.Filled.Close else Icons.Filled.Edit,
                                    contentDescription = if (state.isEditing) "Exit edit mode" else "Edit",
                                    tint = if (state.isEditing) MaterialTheme.colorScheme.primary else MaterialTheme.colorScheme.onSurface,
                                )
                            }
                        }
                        IconButton(onClick = documentActions.share, enabled = !state.isSaving) { Icon(Icons.Default.Share, contentDescription = "Share document") }
                    },
                    scrollBehavior = scrollBehavior,
                )

                // Search bar
                AnimatedVisibility(
                    visible = searchState.isSearchActive,
                    enter = expandVertically(),
                    exit = shrinkVertically(),
                ) {
                    SearchBar(
                        query = searchState.query,
                        onQueryChange = viewModel::updateSearchQuery,
                        totalMatches = searchState.totalMatches,
                        currentMatchIndex = searchState.currentMatchIndex,
                        onNext = viewModel::nextMatch,
                        onPrevious = viewModel::previousMatch,
                        onClose = viewModel::toggleSearch,
                    )
                }

                DocumentEditBar(state.isDirty || state.draftDirty, state.isSaving, state.canUndo, viewModel::undo, documentActions.saveCopy)
                DocumentWarnings(state.document?.warnings.orEmpty())
                // Cell editing bar (at top so keyboard doesn't cover it)
                AnimatedVisibility(
                    visible = state.isEditing && state.editingCell != null,
                    enter = expandVertically(),
                    exit = shrinkVertically(),
                ) {
                    if (state.editingCell != null) {
                        val (editRow, editCol) = state.editingCell!!
                        val sheet = state.document?.sheets?.getOrNull(state.activeSheetIndex)
                        val currentValue = sheet?.rows
                            ?.find { it.rowIndex == editRow }
                            ?.cells?.find { it.columnIndex == editCol }
                            ?.value ?: ""
                        val editText = state.draftText
                        val editFocusRequester = remember { FocusRequester() }

                        LaunchedEffect(editRow, editCol) {
                            editFocusRequester.requestFocus()
                        }

                        Row(
                            modifier = Modifier
                                .fillMaxWidth()
                                .background(MaterialTheme.colorScheme.surface)
                                .padding(horizontal = 8.dp, vertical = 4.dp),
                            verticalAlignment = Alignment.CenterVertically,
                        ) {
                            Text(
                                text = "${columnLetter(editCol)}${editRow + 1}",
                                style = MaterialTheme.typography.labelLarge,
                                fontWeight = FontWeight.Bold,
                                modifier = Modifier.padding(end = 8.dp),
                            )

                            TextField(
                                value = editText,
                                onValueChange = viewModel::updateDraft,
                                modifier = Modifier
                                    .weight(1f)
                                    .focusRequester(editFocusRequester),
                                singleLine = true,
                                textStyle = MaterialTheme.typography.bodyMedium,
                                colors = TextFieldDefaults.colors(
                                    focusedContainerColor = MaterialTheme.colorScheme.surface,
                                    unfocusedContainerColor = MaterialTheme.colorScheme.surface,
                                ),
                                keyboardOptions = KeyboardOptions(imeAction = ImeAction.Done),
                                keyboardActions = KeyboardActions(
                                    onDone = {
                                        viewModel.updateCellValue(editRow, editCol, editText)
                                        viewModel.stopEditingCell()
                                    },
                                ),
                            )

                            IconButton(onClick = {
                                viewModel.updateCellValue(editRow, editCol, editText)
                                viewModel.stopEditingCell()
                            }) {
                                Icon(Icons.Filled.Check, contentDescription = "Confirm")
                            }
                            IconButton(onClick = { viewModel.stopEditingCell() }) {
                                Icon(Icons.Filled.Close, contentDescription = "Cancel")
                            }
                        }
                    }
                }
            }
        },
    ) { innerPadding ->
        when {
            state.isLoading -> {
                LoadingIndicator(modifier = Modifier.padding(innerPadding))
            }
            state.errorMessage != null -> {
                ErrorMessage(
                    message = state.errorMessage!!,
                    modifier = Modifier.padding(innerPadding),
                )
            }
            state.document != null -> {
                val document = state.document!!
                Column(modifier = Modifier.padding(innerPadding)) {
                    // Sheet tabs
                    if (document.sheets.size > 1) {
                        ScrollableTabRow(
                            selectedTabIndex = state.activeSheetIndex,
                            edgePadding = 8.dp,
                        ) {
                            document.sheets.forEachIndexed { index, sheet ->
                                Tab(
                                    selected = index == state.activeSheetIndex,
                                    onClick = { viewModel.selectSheet(index) },
                                    text = {
                                        Text(
                                            text = sheet.name,
                                            maxLines = 1,
                                            overflow = TextOverflow.Ellipsis,
                                        )
                                    },
                                )
                            }
                        }
                    }

                    // Sheet content
                    val sheet = document.sheets.getOrNull(state.activeSheetIndex)
                    if (sheet != null) {
                        SheetContent(
                            sheet = sheet,
                            documentKey = (state.activeUri ?: uri).toString(),
                            styles = document.styles,
                            searchState = searchState,
                            activeSheetIndex = state.activeSheetIndex,
                            zoomScale = zoomScale,
                            onZoomChange = { zoomScale = it },
                            isEditing = state.isEditing,
                            editingCell = state.editingCell,
                            onCellClick = { rowIndex, colIndex ->
                                viewModel.startEditingCell(rowIndex, colIndex)
                            },
                        )
                    }
                }
            }
        }
    }
}

@Composable
private fun SearchBar(
    query: String,
    onQueryChange: (String) -> Unit,
    totalMatches: Int,
    currentMatchIndex: Int,
    onNext: () -> Unit,
    onPrevious: () -> Unit,
    onClose: () -> Unit,
) {
    val focusRequester = remember { FocusRequester() }
    val keyboardController = LocalSoftwareKeyboardController.current

    LaunchedEffect(Unit) {
        focusRequester.requestFocus()
    }

    Row(
        modifier = Modifier
            .fillMaxWidth()
            .background(MaterialTheme.colorScheme.surface)
            .padding(horizontal = 8.dp, vertical = 4.dp),
        verticalAlignment = Alignment.CenterVertically,
    ) {
        OutlinedTextField(
            value = query,
            onValueChange = onQueryChange,
            modifier = Modifier
                .weight(1f)
                .focusRequester(focusRequester),
            placeholder = { Text("Search cells...") },
            singleLine = true,
            textStyle = MaterialTheme.typography.bodyMedium,
            colors = OutlinedTextFieldDefaults.colors(
                focusedBorderColor = MaterialTheme.colorScheme.primary,
            ),
            keyboardOptions = KeyboardOptions(imeAction = ImeAction.Search),
            keyboardActions = KeyboardActions(
                onSearch = {
                    keyboardController?.hide()
                    if (totalMatches > 0) onNext()
                },
            ),
        )

        if (totalMatches > 0) {
            Text(
                text = "${currentMatchIndex + 1}/$totalMatches",
                style = MaterialTheme.typography.labelMedium,
                modifier = Modifier.padding(horizontal = 8.dp),
            )
        }

        IconButton(onClick = onPrevious, enabled = totalMatches > 0) {
            Icon(Icons.Filled.KeyboardArrowUp, contentDescription = "Previous")
        }
        IconButton(onClick = onNext, enabled = totalMatches > 0) {
            Icon(Icons.Filled.KeyboardArrowDown, contentDescription = "Next")
        }
        IconButton(onClick = onClose) {
            Icon(Icons.Filled.Close, contentDescription = "Close search")
        }
    }
}

/** Convert 0-based column index to Excel-style letter (A, B, ..., Z, AA, AB, ...). */
internal fun columnLetter(index: Int): String {
    val sb = StringBuilder()
    var n = index
    do {
        sb.insert(0, ('A' + n % 26))
        n = n / 26 - 1
    } while (n >= 0)
    return sb.toString()
}
