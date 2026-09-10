package com.modocs.feature.pptx

import android.graphics.Bitmap
import android.net.Uri
import androidx.compose.animation.AnimatedVisibility
import androidx.compose.animation.expandVertically
import androidx.compose.animation.shrinkVertically
import androidx.compose.foundation.Image
import androidx.compose.foundation.background
import com.modocs.core.ui.components.ZoomableContainer
import androidx.compose.foundation.layout.Arrangement
import androidx.compose.foundation.layout.Box
import androidx.compose.foundation.layout.Column
import androidx.compose.foundation.layout.Row
import androidx.compose.foundation.layout.Spacer
import androidx.compose.foundation.layout.aspectRatio
import androidx.compose.foundation.layout.fillMaxSize
import androidx.compose.foundation.layout.fillMaxWidth
import androidx.compose.foundation.layout.padding
import androidx.compose.foundation.layout.width
import androidx.compose.foundation.text.KeyboardActions
import androidx.compose.foundation.text.KeyboardOptions
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.automirrored.filled.ArrowBack
import androidx.compose.material.icons.filled.Close
import androidx.compose.material.icons.filled.KeyboardArrowDown
import androidx.compose.material.icons.filled.KeyboardArrowUp
import androidx.compose.material.icons.automirrored.filled.NavigateBefore
import androidx.compose.material.icons.automirrored.filled.NavigateNext
import androidx.compose.material.icons.filled.Search
import androidx.compose.material3.BottomAppBar
import androidx.compose.material3.ExperimentalMaterial3Api
import androidx.compose.material3.Icon
import androidx.compose.material3.IconButton
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.OutlinedTextField
import androidx.compose.material3.OutlinedTextFieldDefaults
import androidx.compose.material3.Scaffold
import androidx.compose.material3.Text
import androidx.compose.material3.TopAppBar
import androidx.compose.material3.TopAppBarDefaults
import androidx.compose.runtime.Composable
import androidx.compose.runtime.LaunchedEffect
import androidx.compose.runtime.getValue
import androidx.compose.runtime.mutableIntStateOf
import androidx.compose.runtime.produceState
import androidx.compose.runtime.remember
import androidx.compose.runtime.setValue
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.focus.FocusRequester
import androidx.compose.ui.focus.focusRequester
import androidx.compose.ui.graphics.asImageBitmap
import androidx.compose.ui.layout.ContentScale
import androidx.compose.ui.layout.onSizeChanged
import androidx.compose.ui.platform.LocalSoftwareKeyboardController
import androidx.compose.ui.text.input.ImeAction
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.text.style.TextOverflow
import androidx.compose.ui.unit.dp
import androidx.hilt.navigation.compose.hiltViewModel
import androidx.lifecycle.compose.collectAsStateWithLifecycle
import com.modocs.core.ui.components.ErrorMessage
import com.modocs.core.ui.components.LoadingIndicator
import com.modocs.core.ui.components.ShareDocumentAction

// Fixed slide canvas background — neutral gray, not affected by dark/light theme
private val SlideCanvasBackground = Color(0xFFE0E0E0)

@OptIn(ExperimentalMaterial3Api::class)
@Composable
fun PptxViewerScreen(
    uri: Uri,
    displayName: String? = null,
    onNavigateBack: () -> Unit,
    viewModel: PptxViewerViewModel = hiltViewModel(),
) {
    val state by viewModel.state.collectAsStateWithLifecycle()
    val searchState by viewModel.searchState.collectAsStateWithLifecycle()

    LaunchedEffect(uri) {
        viewModel.loadPptx(uri, displayName)
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

    val scrollBehavior = TopAppBarDefaults.pinnedScrollBehavior()

    Scaffold(
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
                        IconButton(onClick = onNavigateBack) {
                            Icon(Icons.AutoMirrored.Filled.ArrowBack, contentDescription = "Back")
                        }
                    },
                    actions = {
                        IconButton(onClick = { viewModel.toggleSearch() }) {
                            Icon(Icons.Filled.Search, contentDescription = "Search")
                        }
                        ShareDocumentAction(
                            uri = uri,
                            displayName = state.fileName.ifEmpty { null } ?: displayName,
                        )
                    },
                    scrollBehavior = scrollBehavior,
                )

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
            }
        },
        bottomBar = {
            if (state.slideCount > 0) {
                BottomAppBar {
                    Row(
                        modifier = Modifier.fillMaxWidth(),
                        horizontalArrangement = Arrangement.Center,
                        verticalAlignment = Alignment.CenterVertically,
                    ) {
                        IconButton(
                            onClick = { viewModel.previousSlide() },
                            enabled = state.currentSlide > 0,
                        ) {
                            Icon(Icons.AutoMirrored.Filled.NavigateBefore, contentDescription = "Previous slide")
                        }

                        Text(
                            text = "Slide ${state.currentSlide + 1} of ${state.slideCount}",
                            style = MaterialTheme.typography.bodyMedium,
                        )

                        IconButton(
                            onClick = { viewModel.nextSlide() },
                            enabled = state.currentSlide < state.slideCount - 1,
                        ) {
                            Icon(Icons.AutoMirrored.Filled.NavigateNext, contentDescription = "Next slide")
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
                SlideViewer(
                    viewModel = viewModel,
                    document = state.document!!,
                    currentSlide = state.currentSlide,
                    modifier = Modifier.padding(innerPadding),
                )
            }
        }
    }
}

@Composable
private fun SlideViewer(
    viewModel: PptxViewerViewModel,
    document: PptxDocument,
    currentSlide: Int,
    modifier: Modifier = Modifier,
) {
    var containerWidth by remember { mutableIntStateOf(0) }

    val aspectRatio = document.slideWidth.toFloat() / document.slideHeight.toFloat()

    // Render slide bitmap — use density-aware sizing for tablets
    val slideBitmap by produceState<Bitmap?>(null, currentSlide, containerWidth) {
        if (containerWidth > 0) {
            // Render at 2x for crisp display, higher cap for tablets
            val renderWidth = (containerWidth * 2).coerceAtMost(3840)
            value = viewModel.renderSlide(currentSlide, renderWidth)
        }
    }

    ZoomableContainer(
        modifier = modifier
            .fillMaxSize()
            .background(SlideCanvasBackground)
            .onSizeChanged { containerWidth = it.width },
        maxScale = 5f,
        contentModifier = Modifier.fillMaxSize(),
    ) { _ ->
        val bitmap = slideBitmap
        if (bitmap != null) {
            Image(
                bitmap = bitmap.asImageBitmap(),
                contentDescription = "Slide ${currentSlide + 1}",
                modifier = Modifier
                    .align(Alignment.Center)
                    .fillMaxWidth()
                    .padding(16.dp)
                    .aspectRatio(aspectRatio),
                contentScale = ContentScale.Fit,
            )
        } else if (containerWidth > 0) {
            Box(modifier = Modifier.align(Alignment.Center)) {
                LoadingIndicator()
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
            placeholder = { Text("Search slides...") },
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
