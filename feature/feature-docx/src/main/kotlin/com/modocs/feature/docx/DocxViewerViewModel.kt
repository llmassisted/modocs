package com.modocs.feature.docx

import android.content.Context
import android.net.Uri
import android.provider.OpenableColumns
import androidx.lifecycle.ViewModel
import androidx.lifecycle.viewModelScope
import dagger.hilt.android.lifecycle.HiltViewModel
import dagger.hilt.android.qualifiers.ApplicationContext
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.Job
import kotlinx.coroutines.flow.MutableStateFlow
import kotlinx.coroutines.flow.StateFlow
import kotlinx.coroutines.flow.asStateFlow
import kotlinx.coroutines.flow.MutableSharedFlow
import kotlinx.coroutines.flow.asSharedFlow
import kotlinx.coroutines.launch
import kotlinx.coroutines.withContext
import com.modocs.core.common.OoxmlDecryptor
import javax.inject.Inject

enum class DocxViewMode { CANVAS, COMPOSE }

data class DocxViewerState(
    val isLoading: Boolean = true,
    val fileName: String = "",
    val errorMessage: String? = null,
    val document: DocxDocument? = null,
    val isEditing: Boolean = false,
    val isDirty: Boolean = false,
    val isSaving: Boolean = false,
    val editingElementIndex: Int = -1,
    val viewMode: DocxViewMode = DocxViewMode.CANVAS,
    val pageCount: Int = 0,
    val currentPage: Int = 0,
    val formattingVersion: Int = 0,
    val isPasswordRequired: Boolean = false,
    val passwordError: String? = null,
)

data class DocxSearchState(
    val isSearchActive: Boolean = false,
    val query: String = "",
    val isSearching: Boolean = false,
    val matches: List<DocxSearchMatch> = emptyList(),
    val currentMatchIndex: Int = -1,
) {
    val totalMatches: Int get() = matches.size
    val hasMatches: Boolean get() = matches.isNotEmpty()
    val currentMatch: DocxSearchMatch? get() =
        if (currentMatchIndex in matches.indices) matches[currentMatchIndex] else null
}

data class DocxSearchMatch(
    /** Index of the element in the document body. */
    val elementIndex: Int,
    /** Character range within the element's text. */
    val range: IntRange,
    /** Short context snippet around the match. */
    val snippet: String,
    /** Page index for CANVAS mode navigation. */
    val pageIndex: Int = -1,
)

/**
 * A highlight rectangle for drawing on a page bitmap.
 * Coordinates are normalized (0..1 relative to page dimensions).
 */
data class DocxPageHighlight(
    val rect: DocxHighlightRect,
    val isCurrent: Boolean,
)

data class DocxHighlightRect(
    val left: Float,
    val top: Float,
    val right: Float,
    val bottom: Float,
)

@HiltViewModel
class DocxViewerViewModel @Inject constructor(
    @ApplicationContext private val context: Context,
) : ViewModel() {

    private val _state = MutableStateFlow(DocxViewerState())
    val state: StateFlow<DocxViewerState> = _state.asStateFlow()

    private val _searchState = MutableStateFlow(DocxSearchState())
    val searchState: StateFlow<DocxSearchState> = _searchState.asStateFlow()

    private val _formattingAtCursor = MutableStateFlow(RunProperties())
    val formattingAtCursor: StateFlow<RunProperties> = _formattingAtCursor.asStateFlow()

    sealed interface DocxEvent {
        data class ExportPdfReady(val message: String) : DocxEvent
        data class ExportPdfError(val message: String) : DocxEvent
        data class SaveSuccess(val message: String) : DocxEvent
        data class SaveError(val message: String) : DocxEvent
    }

    private val _events = MutableSharedFlow<DocxEvent>()
    val events = _events.asSharedFlow()

    private val layoutCalculator = PageLayoutCalculator(context)
    private var pageRenderer: DocxPageRenderer? = null
    private var documentUri: Uri? = null
    private var searchJob: Job? = null
    private var editingSelectionStart: Int = 0
    private var editingSelectionEnd: Int = 0

    /**
     * Precomputed plain text per element for fast search.
     * Index matches document.body index.
     */
    private var elementTexts: List<String> = emptyList()
    /**
     * Cumulative character offset for each element (for global search positioning).
     */
    private var elementOffsets: List<Int> = emptyList()

    /**
     * Set of element indices where an automatic page break occurs BEFORE that element.
     * Computed based on estimated element heights and A4 page dimensions.
     */
    private val _autoPageBreaks = MutableStateFlow<Set<Int>>(emptySet())
    val autoPageBreaks: StateFlow<Set<Int>> = _autoPageBreaks.asStateFlow()

    fun loadDocx(uri: Uri, displayName: String?) {
        if (_state.value.document != null) return

        documentUri = uri

        viewModelScope.launch {
            _state.value = _state.value.copy(isLoading = true, errorMessage = null)

            val name = displayName
                ?: resolveFileName(uri)
                ?: "Document"

            // Check if file is password-protected (OLE2 container)
            if (OoxmlDecryptor.isEncryptedOoxmlFile(context, uri)) {
                _state.value = _state.value.copy(
                    isLoading = false,
                    isPasswordRequired = true,
                    fileName = name,
                )
                return@launch
            }

            try {
                val document = DocxParser.parse(context, uri)
                onDocumentLoaded(document, name)
            } catch (e: Exception) {
                _state.value = _state.value.copy(
                    isLoading = false,
                    errorMessage = "Failed to open document: ${e.message ?: "Unknown error"}",
                )
            }
        }
    }

    fun submitPassword(password: String) {
        val uri = documentUri ?: return

        viewModelScope.launch {
            _state.value = _state.value.copy(isLoading = true, passwordError = null)

            when (val result = OoxmlDecryptor.decrypt(context, uri, password)) {
                is OoxmlDecryptor.DecryptResult.Success -> {
                    try {
                        val document = DocxParser.parse(result.inputStream)
                        _state.value = _state.value.copy(
                            isPasswordRequired = false,
                            passwordError = null,
                        )
                        onDocumentLoaded(document, _state.value.fileName)
                    } catch (e: Exception) {
                        _state.value = _state.value.copy(
                            isLoading = false,
                            isPasswordRequired = false,
                            errorMessage = "Failed to open document: ${e.message ?: "Unknown error"}",
                        )
                    }
                }
                is OoxmlDecryptor.DecryptResult.WrongPassword -> {
                    _state.value = _state.value.copy(
                        isLoading = false,
                        passwordError = "Incorrect password",
                    )
                }
                is OoxmlDecryptor.DecryptResult.Failed -> {
                    _state.value = _state.value.copy(
                        isLoading = false,
                        isPasswordRequired = false,
                        errorMessage = "Failed to decrypt: ${result.message}",
                    )
                }
            }
        }
    }

    private fun onDocumentLoaded(document: DocxDocument, name: String) {
        // Precompute text for search
        val texts = document.body.map { element ->
            when (element) {
                is DocxParagraph -> element.text
                is DocxTable -> element.rows.joinToString(" ") { row ->
                    row.cells.joinToString(" ") { cell ->
                        cell.paragraphs.joinToString(" ") { it.text }
                    }
                }
                is DocxImage -> element.altText
            }
        }
        elementTexts = texts

        var offset = 0
        val offsets = mutableListOf<Int>()
        for (text in texts) {
            offsets.add(offset)
            offset += text.length + 1 // +1 for separator
        }
        elementOffsets = offsets

        // Compute automatic page break positions using real font metrics
        _autoPageBreaks.value = layoutCalculator.computeAutoPageBreaks(document)

        // Prepare bitmap page renderer for CANVAS view mode
        val renderer = DocxPageRenderer(context)
        renderer.prepare(document)
        pageRenderer = renderer

        _state.value = _state.value.copy(
            isLoading = false,
            fileName = name,
            document = document,
            pageCount = renderer.pageCount,
        )
    }

    // --- Export to PDF ---

    fun exportToPdf(outputUri: Uri) {
        val document = _state.value.document ?: return

        viewModelScope.launch {
            try {
                val converter = DocxToPdfConverter(context)
                converter.convertToUri(document, outputUri)
                _events.emit(DocxEvent.ExportPdfReady("PDF exported successfully"))
            } catch (e: Exception) {
                _events.emit(DocxEvent.ExportPdfError("Export failed: ${e.message ?: "Unknown error"}"))
            }
        }
    }

    // --- Search ---

    fun toggleSearch() {
        val current = _searchState.value
        if (current.isSearchActive) {
            _searchState.value = DocxSearchState()
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

            val isCanvasMode = _state.value.viewMode == DocxViewMode.CANVAS

            val matches = withContext(Dispatchers.Default) {
                if (isCanvasMode) {
                    searchCanvasMode(query)
                } else {
                    searchComposeMode(query)
                }
            }

            _searchState.value = _searchState.value.copy(
                matches = matches,
                currentMatchIndex = if (matches.isNotEmpty()) 0 else -1,
                isSearching = false,
            )
        }
    }

    /**
     * Search per-page text from the page renderer (CANVAS mode).
     */
    private fun searchCanvasMode(query: String): List<DocxSearchMatch> {
        val renderer = pageRenderer ?: return emptyList()
        val results = mutableListOf<DocxSearchMatch>()
        val queryLower = query.lowercase()

        for (pageIdx in 0 until renderer.pageCount) {
            val pageText = renderer.getPageText(pageIdx)
            if (pageText.isEmpty()) continue

            val textLower = pageText.lowercase()
            var searchFrom = 0

            while (searchFrom < textLower.length) {
                val matchIndex = textLower.indexOf(queryLower, searchFrom)
                if (matchIndex == -1) break

                val snippetStart = (matchIndex - 30).coerceAtLeast(0)
                val snippetEnd = (matchIndex + query.length + 30).coerceAtMost(pageText.length)
                val prefix = if (snippetStart > 0) "..." else ""
                val suffix = if (snippetEnd < pageText.length) "..." else ""
                val snippet = prefix + pageText.substring(snippetStart, snippetEnd) + suffix

                results.add(DocxSearchMatch(
                    elementIndex = 0,
                    range = matchIndex until (matchIndex + query.length),
                    snippet = snippet,
                    pageIndex = pageIdx,
                ))

                searchFrom = matchIndex + 1
            }
        }

        return results
    }

    /**
     * Search per-element text (COMPOSE mode). Also searches across element
     * boundaries by concatenating all text.
     */
    private fun searchComposeMode(query: String): List<DocxSearchMatch> {
        val results = mutableListOf<DocxSearchMatch>()
        val queryLower = query.lowercase()

        // Per-element search (existing logic)
        for ((index, text) in elementTexts.withIndex()) {
            val textLower = text.lowercase()
            var searchFrom = 0

            while (searchFrom < textLower.length) {
                val matchIndex = textLower.indexOf(queryLower, searchFrom)
                if (matchIndex == -1) break

                val snippetStart = (matchIndex - 30).coerceAtLeast(0)
                val snippetEnd = (matchIndex + query.length + 30).coerceAtMost(text.length)
                val prefix = if (snippetStart > 0) "..." else ""
                val suffix = if (snippetEnd < text.length) "..." else ""
                val snippet = prefix + text.substring(snippetStart, snippetEnd) + suffix

                results.add(DocxSearchMatch(
                    elementIndex = index,
                    range = matchIndex until (matchIndex + query.length),
                    snippet = snippet,
                ))

                searchFrom = matchIndex + 1
            }
        }

        // Cross-element search: concatenate all text and find matches that span elements
        if (elementTexts.size > 1) {
            val fullText = elementTexts.joinToString(" ").lowercase()
            var searchFrom = 0
            while (searchFrom < fullText.length) {
                val matchIndex = fullText.indexOf(queryLower, searchFrom)
                if (matchIndex == -1) break

                // Map back to element index
                var charOffset = 0
                var elementIdx = 0
                for ((idx, text) in elementTexts.withIndex()) {
                    if (matchIndex < charOffset + text.length) {
                        elementIdx = idx
                        break
                    }
                    charOffset += text.length + 1 // +1 for space separator
                }

                // Check if this match spans element boundaries (not already found)
                val localOffset = matchIndex - charOffset
                val elementText = elementTexts[elementIdx]
                val existsInElement = localOffset >= 0 && localOffset + query.length <= elementText.length

                if (!existsInElement) {
                    // Cross-element match — assign to the starting element
                    val snippetStart = (matchIndex - 20).coerceAtLeast(0)
                    val snippetEnd = (matchIndex + query.length + 20).coerceAtMost(fullText.length)
                    val snippet = "..." + fullText.substring(snippetStart, snippetEnd) + "..."

                    results.add(DocxSearchMatch(
                        elementIndex = elementIdx,
                        range = localOffset.coerceAtLeast(0) until (localOffset + query.length).coerceAtMost(elementText.length),
                        snippet = snippet,
                    ))
                }

                searchFrom = matchIndex + 1
            }
        }

        return results
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

    /**
     * Get search highlight ranges for a specific element index.
     * Returns pairs of (IntRange, globalMatchIndex) for highlight rendering.
     */
    fun getHighlightsForElement(elementIndex: Int): List<Pair<IntRange, Int>> {
        val search = _searchState.value
        if (!search.hasMatches) return emptyList()

        return search.matches
            .mapIndexedNotNull { globalIndex, match ->
                if (match.elementIndex == elementIndex) {
                    match.range to globalIndex
                } else null
            }
    }

    // --- Page Renderer ---

    fun getPageRenderer(): DocxPageRenderer? = pageRenderer

    fun updateCurrentPage(page: Int) {
        _state.value = _state.value.copy(currentPage = page)
    }

    /**
     * Get search highlight rectangles for a specific page (CANVAS mode).
     */
    fun getHighlightsForPage(pageIndex: Int): List<DocxPageHighlight> {
        val search = _searchState.value
        if (!search.hasMatches) return emptyList()

        val renderer = pageRenderer ?: return emptyList()
        val segments = renderer.getTextSegments(pageIndex)
        if (segments.isEmpty()) return emptyList()

        val pageText = renderer.getPageText(pageIndex).lowercase()
        val query = search.query.lowercase()
        if (query.length < 2) return emptyList()

        val highlights = mutableListOf<DocxPageHighlight>()
        var searchFrom = 0

        // Find all query occurrences in the page text
        while (searchFrom < pageText.length) {
            val matchIdx = pageText.indexOf(query, searchFrom)
            if (matchIdx == -1) break

            // Map character offset to segment positions
            var charOffset = 0
            for (seg in segments) {
                val segEnd = charOffset + seg.text.length
                if (matchIdx < segEnd && matchIdx + query.length > charOffset) {
                    // This segment overlaps with the match
                    val isCurrentMatch = isCurrentMatchOnPage(pageIndex, matchIdx, search)
                    highlights.add(DocxPageHighlight(
                        rect = DocxHighlightRect(
                            left = seg.x,
                            top = seg.y,
                            right = seg.x + seg.width,
                            bottom = seg.y + seg.height,
                        ),
                        isCurrent = isCurrentMatch,
                    ))
                }
                charOffset = segEnd
            }

            searchFrom = matchIdx + 1
        }

        return highlights
    }

    private fun isCurrentMatchOnPage(pageIndex: Int, charOffset: Int, search: DocxSearchState): Boolean {
        val currentMatch = search.currentMatch ?: return false
        return currentMatch.pageIndex == pageIndex
    }

    // --- Editing ---

    fun toggleEditMode() {
        val current = _state.value
        val enteringEdit = !current.isEditing

        if (enteringEdit) {
            _state.value = current.copy(
                isEditing = true,
                editingElementIndex = -1,
                viewMode = DocxViewMode.COMPOSE,
            )
        } else {
            // Leaving edit mode — switch back to CANVAS
            if (current.isDirty) {
                // Re-prepare page renderer if document was modified
                val doc = current.document
                if (doc != null) {
                    pageRenderer?.invalidateCache()
                    pageRenderer?.prepare(doc)
                    _autoPageBreaks.value = layoutCalculator.computeAutoPageBreaks(doc)
                }
            }
            _state.value = current.copy(
                isEditing = false,
                editingElementIndex = -1,
                viewMode = DocxViewMode.CANVAS,
                pageCount = pageRenderer?.pageCount ?: 0,
            )
        }
    }

    fun startEditingElement(index: Int) {
        if (!_state.value.isEditing) return
        editingSelectionStart = 0
        editingSelectionEnd = 0
        _state.value = _state.value.copy(editingElementIndex = index)
        updateFormattingAtCursor()
    }

    fun stopEditingElement() {
        editingSelectionStart = 0
        editingSelectionEnd = 0
        _state.value = _state.value.copy(editingElementIndex = -1)
    }

    fun updateEditingSelection(start: Int, end: Int) {
        editingSelectionStart = start
        editingSelectionEnd = end
        updateFormattingAtCursor()
    }

    fun updateParagraphText(elementIndex: Int, newText: String) {
        val document = _state.value.document ?: return
        val element = document.body.getOrNull(elementIndex) as? DocxParagraph ?: return

        // Skip if text hasn't changed (e.g., only cursor/selection moved)
        if (element.text == newText) return

        // Rebuild runs: keep first run's properties, replace text
        val baseProps = element.runs.firstOrNull()?.properties ?: RunProperties()
        val newRuns = listOf(DocxRun(newText, baseProps))

        document.body[elementIndex] = element.copy(runs = newRuns)
        if (!_state.value.isDirty) {
            _state.value = _state.value.copy(isDirty = true)
        }

        // Update search text cache
        if (elementIndex < elementTexts.size) {
            val mutableTexts = elementTexts.toMutableList()
            mutableTexts[elementIndex] = newText
            elementTexts = mutableTexts
        }
    }

    fun toggleBold(elementIndex: Int) {
        applyFormattingToSelection(elementIndex) { it.copy(bold = !it.bold) }
    }

    fun toggleItalic(elementIndex: Int) {
        applyFormattingToSelection(elementIndex) { it.copy(italic = !it.italic) }
    }

    fun toggleUnderline(elementIndex: Int) {
        applyFormattingToSelection(elementIndex) { it.copy(underline = !it.underline) }
    }

    private fun applyFormattingToSelection(
        elementIndex: Int,
        transform: (RunProperties) -> RunProperties,
    ) {
        val document = _state.value.document ?: return
        val element = document.body.getOrNull(elementIndex) as? DocxParagraph ?: return

        val selStart = editingSelectionStart
        val selEnd = editingSelectionEnd

        val newRuns = if (selStart == selEnd) {
            // No selection — apply to all runs
            element.runs.map { run -> run.copy(properties = transform(run.properties)) }
        } else {
            splitAndTransformRuns(element.runs, selStart, selEnd, transform)
        }

        val newBody = document.body.toMutableList()
        newBody[elementIndex] = element.copy(runs = newRuns)
        _state.value = _state.value.copy(
            isDirty = true,
            document = document.copy(body = newBody),
            formattingVersion = _state.value.formattingVersion + 1,
        )
        updateFormattingAtCursor()
    }

    private fun splitAndTransformRuns(
        runs: List<DocxRun>,
        selStart: Int,
        selEnd: Int,
        transform: (RunProperties) -> RunProperties,
    ): List<DocxRun> {
        val result = mutableListOf<DocxRun>()
        var textPos = 0

        for (run in runs) {
            if (run.isPageBreak) {
                result.add(run)
                continue
            }

            val runText = if (run.isBreak) "\n" else run.text
            val runStart = textPos
            val runEnd = textPos + runText.length
            textPos = runEnd

            when {
                runEnd <= selStart || runStart >= selEnd -> result.add(run)
                runStart >= selStart && runEnd <= selEnd ->
                    result.add(run.copy(properties = transform(run.properties)))
                else -> {
                    val splitStart = (selStart - runStart).coerceAtLeast(0)
                    val splitEnd = (selEnd - runStart).coerceAtMost(runText.length)
                    if (splitStart > 0) {
                        result.add(DocxRun(runText.substring(0, splitStart), run.properties))
                    }
                    result.add(DocxRun(
                        runText.substring(splitStart, splitEnd),
                        transform(run.properties),
                    ))
                    if (splitEnd < runText.length) {
                        result.add(DocxRun(runText.substring(splitEnd), run.properties))
                    }
                }
            }
        }

        return result
    }

    private fun updateFormattingAtCursor() {
        val document = _state.value.document ?: return
        val element = document.body.getOrNull(_state.value.editingElementIndex) as? DocxParagraph ?: return

        val pos = editingSelectionStart
        var textPos = 0
        for (run in element.runs) {
            if (run.isPageBreak) continue
            val runText = if (run.isBreak) "\n" else run.text
            val runEnd = textPos + runText.length
            if (pos < runEnd) {
                _formattingAtCursor.value = run.properties
                return
            }
            textPos = runEnd
        }
        _formattingAtCursor.value = element.runs.lastOrNull()?.properties ?: RunProperties()
    }

    fun saveDocument() {
        val document = _state.value.document ?: return
        val uri = documentUri ?: return

        viewModelScope.launch {
            _state.value = _state.value.copy(isSaving = true)
            try {
                DocxWriter.save(context, document, uri)
                _state.value = _state.value.copy(isDirty = false, isSaving = false)
                _events.emit(DocxEvent.SaveSuccess("Document saved"))
            } catch (e: Exception) {
                _state.value = _state.value.copy(isSaving = false)
                _events.emit(DocxEvent.SaveError("Save failed: ${e.message ?: "Unknown error"}"))
            }
        }
    }

    fun saveDocumentAs(outputUri: Uri) {
        val document = _state.value.document ?: return

        viewModelScope.launch {
            _state.value = _state.value.copy(isSaving = true)
            try {
                DocxWriter.save(context, document, outputUri)
                _state.value = _state.value.copy(isDirty = false, isSaving = false)
                _events.emit(DocxEvent.SaveSuccess("Document saved"))
            } catch (e: Exception) {
                _state.value = _state.value.copy(isSaving = false)
                _events.emit(DocxEvent.SaveError("Save failed: ${e.message ?: "Unknown error"}"))
            }
        }
    }

    // --- Helpers ---

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
    }
}
