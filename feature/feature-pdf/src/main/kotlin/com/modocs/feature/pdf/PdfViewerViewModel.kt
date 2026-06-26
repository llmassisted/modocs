package com.modocs.feature.pdf

import android.content.Context
import android.net.Uri
import android.provider.OpenableColumns
import androidx.compose.ui.geometry.Offset
import androidx.compose.ui.geometry.Rect
import androidx.lifecycle.SavedStateHandle
import androidx.lifecycle.ViewModel
import androidx.lifecycle.viewModelScope
import dagger.hilt.android.lifecycle.HiltViewModel
import dagger.hilt.android.qualifiers.ApplicationContext
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.Job
import kotlinx.coroutines.flow.MutableSharedFlow
import kotlinx.coroutines.flow.MutableStateFlow
import kotlinx.coroutines.flow.asSharedFlow
import kotlinx.coroutines.flow.StateFlow
import kotlinx.coroutines.flow.asStateFlow
import kotlinx.coroutines.launch
import kotlinx.coroutines.withContext
import java.text.SimpleDateFormat
import java.util.Date
import java.util.Locale
import java.util.UUID
import javax.inject.Inject

data class SearchMatch(
    val pageIndex: Int,
    val startIndex: Int,
    val endIndex: Int,
    val contextSnippet: String,
    /** Normalized highlight rects (0..1) for this match on the page. */
    val highlightRects: List<Rect>,
)

data class SearchState(
    val isSearchActive: Boolean = false,
    val query: String = "",
    val isSearching: Boolean = false,
    val matches: List<SearchMatch> = emptyList(),
    val currentMatchIndex: Int = -1,
    val textExtracted: Boolean = false,
) {
    val totalMatches: Int get() = matches.size
    val hasMatches: Boolean get() = matches.isNotEmpty()
    val currentMatch: SearchMatch? get() =
        if (currentMatchIndex in matches.indices) matches[currentMatchIndex] else null
}

data class PdfViewerState(
    val isLoading: Boolean = true,
    val pageCount: Int = 0,
    val currentPage: Int = 0,
    val fileName: String = "",
    val errorMessage: String? = null,
    val isPasswordRequired: Boolean = false,
    val passwordError: String? = null,
)

data class FillSignState(
    val isActive: Boolean = false,
    val selectedTool: FillSignTool = FillSignTool.TEXT,
    val annotations: List<PdfAnnotation> = emptyList(),
    val showSignaturePad: Boolean = false,
    /** Pending signature strokes waiting to be placed. */
    val pendingSignature: List<List<Offset>>? = null,
    val editingAnnotationId: String? = null,
    val editingText: String = "",
    val isSaving: Boolean = false,
    val isDirty: Boolean = false,
    /** Currently selected annotation for moving/resizing. */
    val selectedAnnotationId: String? = null,
) {
    val hasAnnotations: Boolean get() = annotations.isNotEmpty()
    val selectedAnnotation: PdfAnnotation? get() = annotations.find { it.id == selectedAnnotationId }
}

@HiltViewModel
class PdfViewerViewModel @Inject constructor(
    @ApplicationContext private val context: Context,
    savedStateHandle: SavedStateHandle,
) : ViewModel() {

    private val _state = MutableStateFlow(PdfViewerState())
    val state: StateFlow<PdfViewerState> = _state.asStateFlow()

    private val _searchState = MutableStateFlow(SearchState())
    val searchState: StateFlow<SearchState> = _searchState.asStateFlow()

    private val _fillSignState = MutableStateFlow(FillSignState())
    val fillSignState: StateFlow<FillSignState> = _fillSignState.asStateFlow()

    sealed interface PdfEvent {
        data class SaveSuccess(val message: String) : PdfEvent
        data class SaveError(val message: String) : PdfEvent
    }

    private val _events = MutableSharedFlow<PdfEvent>()
    val events = _events.asSharedFlow()

    private var pdfRenderer: PdfRendererWrapper? = null
    private var documentUri: Uri? = null
    private var renderSourceUri: Uri? = null
    private var pageTexts: List<PdfTextExtractor.PageText> = emptyList()
    private var searchJob: Job? = null
    private var decryptedTempFile: java.io.File? = null

    fun getRenderer(): PdfRendererWrapper? = pdfRenderer

    fun loadPdf(uri: Uri, displayName: String?) {
        if (pdfRenderer != null) return

        documentUri = uri
        renderSourceUri = uri

        viewModelScope.launch {
            _state.value = _state.value.copy(isLoading = true, errorMessage = null)

            val name = displayName
                ?: resolveFileName(uri)
                ?: "Document"

            when (val result = PdfRendererWrapper.open(context, uri)) {
                is PdfOpenResult.Success -> {
                    pdfRenderer = result.wrapper
                    _state.value = _state.value.copy(
                        isLoading = false,
                        pageCount = result.wrapper.pageCount,
                        fileName = name,
                    )
                    extractTextAsync(uri)
                }
                is PdfOpenResult.PasswordRequired -> {
                    _state.value = _state.value.copy(
                        isLoading = false,
                        isPasswordRequired = true,
                        fileName = name,
                    )
                }
                is PdfOpenResult.Failed -> {
                    _state.value = _state.value.copy(
                        isLoading = false,
                        errorMessage = "Failed to open PDF: ${result.message}",
                    )
                }
            }
        }
    }

    fun submitPassword(password: String) {
        val uri = documentUri ?: return

        viewModelScope.launch {
            _state.value = _state.value.copy(isLoading = true, passwordError = null)

            when (val result = PdfDecryptor.decrypt(context, uri, password)) {
                is PdfDecryptor.DecryptResult.Success -> {
                    decryptedTempFile = result.tempFile
                    val decryptedUri = Uri.fromFile(result.tempFile)

                    when (val openResult = PdfRendererWrapper.openFile(result.tempFile)) {
                        is PdfOpenResult.Success -> {
                            pdfRenderer = openResult.wrapper
                            renderSourceUri = decryptedUri
                            _state.value = _state.value.copy(
                                isLoading = false,
                                isPasswordRequired = false,
                                passwordError = null,
                                pageCount = openResult.wrapper.pageCount,
                            )
                            // Extract from the decrypted file, not the still-encrypted source,
                            // otherwise search finds nothing on password-protected PDFs.
                            extractTextAsync(decryptedUri)
                        }
                        else -> {
                            _state.value = _state.value.copy(
                                isLoading = false,
                                isPasswordRequired = false,
                                errorMessage = "Failed to open decrypted PDF",
                            )
                        }
                    }
                }
                is PdfDecryptor.DecryptResult.WrongPassword -> {
                    _state.value = _state.value.copy(
                        isLoading = false,
                        passwordError = "Incorrect password",
                    )
                }
                is PdfDecryptor.DecryptResult.Failed -> {
                    _state.value = _state.value.copy(
                        isLoading = false,
                        isPasswordRequired = false,
                        errorMessage = "Failed to decrypt: ${result.message}",
                    )
                }
            }
        }
    }

    private fun extractTextAsync(uri: Uri) {
        viewModelScope.launch {
            pageTexts = PdfTextExtractor.extractText(context, uri)
            _searchState.value = _searchState.value.copy(textExtracted = true)
        }
    }

    fun updateCurrentPage(page: Int) {
        _state.value = _state.value.copy(currentPage = page)
    }

    /**
     * Get highlight rects for a specific page (normalized 0..1 coordinates).
     */
    fun getHighlightsForPage(pageIndex: Int): List<PageHighlight> {
        val search = _searchState.value
        if (!search.hasMatches || search.query.length < 2) return emptyList()

        return search.matches
            .filter { it.pageIndex == pageIndex }
            .flatMap { match ->
                val isCurrent = search.matches.indexOf(match) == search.currentMatchIndex
                match.highlightRects.map { rect ->
                    PageHighlight(rect = rect, isCurrent = isCurrent)
                }
            }
    }

    // --- Search ---

    fun toggleSearch() {
        val current = _searchState.value
        if (current.isSearchActive) {
            _searchState.value = SearchState(textExtracted = current.textExtracted)
        } else {
            _searchState.value = current.copy(isSearchActive = true)
        }
    }

    fun updateSearchQuery(query: String) {
        _searchState.value = _searchState.value.copy(query = query)
        performSearch(query)
    }

    private fun performSearch(query: String) {
        searchJob?.cancel()

        if (query.length < 2) {
            _searchState.value = _searchState.value.copy(
                matches = emptyList(),
                currentMatchIndex = -1,
                isSearching = false,
            )
            return
        }

        searchJob = viewModelScope.launch {
            _searchState.value = _searchState.value.copy(isSearching = true)

            val matches = withContext(Dispatchers.Default) {
                val results = mutableListOf<SearchMatch>()
                val queryLower = query.lowercase()

                for (page in pageTexts) {
                    val textLower = page.text.lowercase()
                    var searchFrom = 0

                    while (searchFrom < textLower.length) {
                        val index = textLower.indexOf(queryLower, searchFrom)
                        if (index == -1) break

                        val snippetStart = (index - 30).coerceAtLeast(0)
                        val snippetEnd = (index + query.length + 30).coerceAtMost(page.text.length)
                        val prefix = if (snippetStart > 0) "..." else ""
                        val suffix = if (snippetEnd < page.text.length) "..." else ""
                        val snippet = prefix + page.text.substring(snippetStart, snippetEnd) + suffix

                        val highlightRects = findHighlightRects(
                            page.segments,
                            page.text,
                            index,
                            index + query.length,
                        )

                        results.add(
                            SearchMatch(
                                pageIndex = page.pageIndex,
                                startIndex = index,
                                endIndex = index + query.length,
                                contextSnippet = snippet,
                                highlightRects = highlightRects,
                            )
                        )

                        searchFrom = index + 1
                    }
                }

                results
            }

            _searchState.value = _searchState.value.copy(
                matches = matches,
                currentMatchIndex = if (matches.isNotEmpty()) 0 else -1,
                isSearching = false,
            )
        }
    }

    private fun findHighlightRects(
        segments: List<PdfTextExtractor.TextSegment>,
        fullText: String,
        matchStart: Int,
        matchEnd: Int,
    ): List<Rect> {
        val rects = mutableListOf<Rect>()

        var charPos = 0
        for (segment in segments) {
            val segStart = charPos
            val segEnd = charPos + segment.text.length

            if (segEnd > matchStart && segStart < matchEnd) {
                val overlapStart = (matchStart - segStart).coerceAtLeast(0)
                val overlapEnd = (matchEnd - segStart).coerceAtMost(segment.text.length)

                val charWidth = if (segment.text.isNotEmpty()) segment.width / segment.text.length else 0f
                val rectX = segment.x + overlapStart * charWidth
                val rectW = (overlapEnd - overlapStart) * charWidth

                rects.add(
                    Rect(
                        left = rectX,
                        top = segment.y,
                        right = (rectX + rectW).coerceAtMost(1f),
                        bottom = (segment.y + segment.height).coerceAtMost(1f),
                    )
                )
            }

            charPos = segEnd
        }

        if (rects.isEmpty()) {
            val matchText = fullText.substring(matchStart, matchEnd).lowercase()
            for (segment in segments) {
                if (segment.text.lowercase().contains(matchText)) {
                    rects.add(
                        Rect(
                            left = segment.x,
                            top = segment.y,
                            right = (segment.x + segment.width).coerceAtMost(1f),
                            bottom = (segment.y + segment.height).coerceAtMost(1f),
                        )
                    )
                    break
                }
            }
        }

        return rects
    }

    fun nextMatch() {
        val current = _searchState.value
        if (current.matches.isEmpty()) return
        val nextIndex = (current.currentMatchIndex + 1) % current.matches.size
        _searchState.value = current.copy(currentMatchIndex = nextIndex)
    }

    fun previousMatch() {
        val current = _searchState.value
        if (current.matches.isEmpty()) return
        val prevIndex = if (current.currentMatchIndex <= 0) {
            current.matches.size - 1
        } else {
            current.currentMatchIndex - 1
        }
        _searchState.value = current.copy(currentMatchIndex = prevIndex)
    }

    // --- Fill & Sign ---

    fun toggleFillSign() {
        val current = _fillSignState.value
        _fillSignState.value = current.copy(isActive = !current.isActive)
        // Close search when entering fill/sign mode
        if (!current.isActive) {
            val search = _searchState.value
            if (search.isSearchActive) {
                _searchState.value = SearchState(textExtracted = search.textExtracted)
            }
        }
    }

    fun selectTool(tool: FillSignTool) {
        val current = _fillSignState.value
        // Cancel any pending editing and deselect annotation
        _fillSignState.value = current.copy(
            selectedTool = tool,
            editingAnnotationId = null,
            editingText = "",
            selectedAnnotationId = null,
        )
        if (tool == FillSignTool.SIGNATURE && current.pendingSignature == null) {
            _fillSignState.value = _fillSignState.value.copy(showSignaturePad = true)
        }
    }

    fun showSignaturePad() {
        _fillSignState.value = _fillSignState.value.copy(showSignaturePad = true)
    }

    fun dismissSignaturePad() {
        _fillSignState.value = _fillSignState.value.copy(showSignaturePad = false)
    }

    fun onSignatureDrawn(strokes: List<List<Offset>>) {
        // Normalize strokes relative to pad canvas size — strokes come in pixel coords
        // They'll be re-normalized when placed on the page
        _fillSignState.value = _fillSignState.value.copy(
            showSignaturePad = false,
            pendingSignature = strokes,
            selectedTool = FillSignTool.SIGNATURE,
        )
    }

    /**
     * Hit-test: check if a tap hits an existing annotation on this page.
     * Returns the annotation ID if hit, null otherwise.
     */
    fun hitTestAnnotation(pageIndex: Int, normX: Float, normY: Float, pageWidthPx: Float, pageHeightPx: Float): String? {
        val annotations = _fillSignState.value.annotations.filter { it.pageIndex == pageIndex }
        val scaleFactor = pageWidthPx / 595f

        // Check in reverse order (topmost first)
        for (ann in annotations.reversed()) {
            val ax = ann.x
            val ay = ann.y
            when (ann) {
                is TextAnnotation -> {
                    if (ann.text.isEmpty()) continue
                    val textH = ann.fontSizeSp * scaleFactor / pageHeightPx
                    val textW = ann.text.length * ann.fontSizeSp * 0.6f * scaleFactor / pageWidthPx
                    if (normX in ax..(ax + textW) && normY in ay..(ay + textH * 1.5f)) return ann.id
                }
                is DateAnnotation -> {
                    val textH = ann.fontSizeSp * scaleFactor / pageHeightPx
                    val textW = ann.dateText.length * ann.fontSizeSp * 0.6f * scaleFactor / pageWidthPx
                    if (normX in ax..(ax + textW) && normY in ay..(ay + textH * 1.5f)) return ann.id
                }
                is CheckmarkAnnotation -> {
                    val sz = ann.sizeSp * scaleFactor
                    val szNormW = sz / pageWidthPx
                    val szNormH = sz / pageHeightPx
                    if (normX in (ax - szNormW * 0.2f)..(ax + szNormW * 1.2f) &&
                        normY in (ay - szNormH * 0.2f)..(ay + szNormH * 1.2f)) return ann.id
                }
                is SignatureAnnotation -> {
                    if (normX in ax..(ax + ann.width) && normY in ay..(ay + ann.height)) return ann.id
                }
            }
        }
        return null
    }

    fun onPageTap(pageIndex: Int, normalizedX: Float, normalizedY: Float, pageWidthPx: Float, pageHeightPx: Float) {
        val current = _fillSignState.value
        if (!current.isActive) return

        // Hit-test existing annotations first
        val hitId = hitTestAnnotation(pageIndex, normalizedX, normalizedY, pageWidthPx, pageHeightPx)
        if (hitId != null) {
            _fillSignState.value = current.copy(selectedAnnotationId = hitId)
            return
        }

        // Deselect if tapping empty space while something is selected
        if (current.selectedAnnotationId != null) {
            _fillSignState.value = current.copy(selectedAnnotationId = null)
            return
        }

        when (current.selectedTool) {
            FillSignTool.TEXT -> {
                val id = UUID.randomUUID().toString()
                val annotation = TextAnnotation(
                    id = id,
                    pageIndex = pageIndex,
                    x = normalizedX,
                    y = normalizedY,
                    text = "",
                )
                _fillSignState.value = current.copy(
                    annotations = current.annotations + annotation,
                    editingAnnotationId = id,
                    editingText = "",
                    isDirty = true,
                )
            }

            FillSignTool.SIGNATURE -> {
                val strokes = current.pendingSignature ?: run {
                    _fillSignState.value = current.copy(showSignaturePad = true)
                    return
                }
                // Normalize the pixel-coordinate strokes to 0..1 relative to the signature pad size
                // The pad is the full canvas, so find bounds
                val allPoints = strokes.flatten()
                if (allPoints.isEmpty()) return
                val maxX = allPoints.maxOf { it.x }
                val maxY = allPoints.maxOf { it.y }
                if (maxX <= 0f || maxY <= 0f) return

                val normalizedStrokes = strokes.map { stroke ->
                    stroke.map { pt -> Offset(pt.x / maxX, pt.y / maxY) }
                }

                val annotation = SignatureAnnotation(
                    id = UUID.randomUUID().toString(),
                    pageIndex = pageIndex,
                    x = normalizedX,
                    y = normalizedY,
                    width = 0.3f,
                    height = 0.08f,
                    strokes = normalizedStrokes,
                )
                _fillSignState.value = current.copy(
                    annotations = current.annotations + annotation,
                    isDirty = true,
                )
            }

            FillSignTool.CHECKMARK -> {
                val annotation = CheckmarkAnnotation(
                    id = UUID.randomUUID().toString(),
                    pageIndex = pageIndex,
                    x = normalizedX,
                    y = normalizedY,
                    isCheck = true,
                )
                _fillSignState.value = current.copy(
                    annotations = current.annotations + annotation,
                    isDirty = true,
                )
            }

            FillSignTool.CROSS -> {
                val annotation = CheckmarkAnnotation(
                    id = UUID.randomUUID().toString(),
                    pageIndex = pageIndex,
                    x = normalizedX,
                    y = normalizedY,
                    isCheck = false,
                )
                _fillSignState.value = current.copy(
                    annotations = current.annotations + annotation,
                    isDirty = true,
                )
            }

            FillSignTool.DATE -> {
                val dateStr = SimpleDateFormat("MM/dd/yyyy", Locale.US).format(Date())
                val annotation = DateAnnotation(
                    id = UUID.randomUUID().toString(),
                    pageIndex = pageIndex,
                    x = normalizedX,
                    y = normalizedY,
                    dateText = dateStr,
                )
                _fillSignState.value = current.copy(
                    annotations = current.annotations + annotation,
                    isDirty = true,
                )
            }
        }
    }

    fun updateEditingText(text: String) {
        _fillSignState.value = _fillSignState.value.copy(editingText = text)
    }

    fun confirmTextEdit() {
        val current = _fillSignState.value
        val editId = current.editingAnnotationId ?: return
        val text = current.editingText.trim()

        if (text.isEmpty()) {
            // Remove empty text annotation
            _fillSignState.value = current.copy(
                annotations = current.annotations.filter { it.id != editId },
                editingAnnotationId = null,
                editingText = "",
            )
        } else {
            _fillSignState.value = current.copy(
                annotations = current.annotations.map { ann ->
                    if (ann.id == editId && ann is TextAnnotation) {
                        ann.copy(text = text)
                    } else ann
                },
                editingAnnotationId = null,
                editingText = "",
            )
        }
    }

    fun cancelTextEdit() {
        val current = _fillSignState.value
        val editId = current.editingAnnotationId ?: return
        val ann = current.annotations.find { it.id == editId }
        // If it's a new empty text annotation, remove it
        if (ann is TextAnnotation && ann.text.isEmpty()) {
            _fillSignState.value = current.copy(
                annotations = current.annotations.filter { it.id != editId },
                editingAnnotationId = null,
                editingText = "",
            )
        } else {
            _fillSignState.value = current.copy(
                editingAnnotationId = null,
                editingText = "",
            )
        }
    }

    fun deleteAnnotation(id: String) {
        val current = _fillSignState.value
        _fillSignState.value = current.copy(
            annotations = current.annotations.filter { it.id != id },
            isDirty = current.annotations.any { it.id == id },
        )
    }

    fun selectAnnotation(id: String?) {
        _fillSignState.value = _fillSignState.value.copy(selectedAnnotationId = id)
    }

    fun moveAnnotation(id: String, newX: Float, newY: Float) {
        val current = _fillSignState.value
        _fillSignState.value = current.copy(
            annotations = current.annotations.map { ann ->
                if (ann.id != id) return@map ann
                when (ann) {
                    is TextAnnotation -> ann.copy(x = newX, y = newY)
                    is SignatureAnnotation -> ann.copy(x = newX, y = newY)
                    is CheckmarkAnnotation -> ann.copy(x = newX, y = newY)
                    is DateAnnotation -> ann.copy(x = newX, y = newY)
                }
            },
            isDirty = true,
        )
    }

    fun changeAnnotationSize(id: String, delta: Float) {
        val current = _fillSignState.value
        _fillSignState.value = current.copy(
            annotations = current.annotations.map { ann ->
                if (ann.id != id) return@map ann
                when (ann) {
                    is TextAnnotation -> ann.copy(fontSizeSp = (ann.fontSizeSp + delta).coerceIn(6f, 72f))
                    is DateAnnotation -> ann.copy(fontSizeSp = (ann.fontSizeSp + delta).coerceIn(6f, 72f))
                    is CheckmarkAnnotation -> ann.copy(sizeSp = (ann.sizeSp + delta).coerceIn(8f, 60f))
                    is SignatureAnnotation -> {
                        val scaleFactor = 1f + (delta / 20f)
                        ann.copy(
                            width = (ann.width * scaleFactor).coerceIn(0.05f, 0.8f),
                            height = (ann.height * scaleFactor).coerceIn(0.02f, 0.3f),
                        )
                    }
                }
            },
            isDirty = true,
        )
    }

    fun undoLastAnnotation() {
        val current = _fillSignState.value
        if (current.annotations.isEmpty()) return
        _fillSignState.value = current.copy(
            annotations = current.annotations.dropLast(1),
            selectedAnnotationId = null,
            isDirty = current.annotations.size > 1,
        )
    }

    fun getAnnotationsForPage(pageIndex: Int): List<PdfAnnotation> {
        return _fillSignState.value.annotations.filter { it.pageIndex == pageIndex }
    }

    fun saveFilled(outputUri: Uri) {
        val srcUri = renderSourceUri ?: documentUri ?: return
        val pageCount = _state.value.pageCount
        val annotations = _fillSignState.value.annotations

        viewModelScope.launch {
            _fillSignState.value = _fillSignState.value.copy(isSaving = true)
            try {
                PdfDocumentWriter.save(
                    context = context,
                    sourceUri = srcUri,
                    outputUri = outputUri,
                    pageCount = pageCount,
                    annotations = annotations,
                )
                _fillSignState.value = _fillSignState.value.copy(
                    isSaving = false,
                    isDirty = false,
                )
                _events.emit(PdfEvent.SaveSuccess("PDF saved successfully"))
            } catch (e: Exception) {
                _fillSignState.value = _fillSignState.value.copy(isSaving = false)
                _events.emit(PdfEvent.SaveError("Failed to save: ${e.message ?: "Unknown error"}"))
            }
        }
    }

    // --- Utility ---

    private suspend fun resolveFileName(uri: Uri): String? = withContext(Dispatchers.IO) {
        try {
            context.contentResolver.query(uri, null, null, null, null)?.use { cursor ->
                if (cursor.moveToFirst()) {
                    val nameIndex = cursor.getColumnIndex(OpenableColumns.DISPLAY_NAME)
                    if (nameIndex >= 0) cursor.getString(nameIndex) else null
                } else null
            }
        } catch (_: Exception) {
            null
        }
    }

    override fun onCleared() {
        super.onCleared()
        searchJob?.cancel()
        pdfRenderer?.close()
        pdfRenderer = null
        decryptedTempFile?.delete()
        decryptedTempFile = null
    }
}

data class PageHighlight(
    val rect: Rect,
    val isCurrent: Boolean,
)
