package com.modocs.feature.xlsx

import android.content.Context
import android.net.Uri
import android.provider.OpenableColumns
import androidx.lifecycle.ViewModel
import androidx.lifecycle.viewModelScope
import dagger.hilt.android.lifecycle.HiltViewModel
import dagger.hilt.android.qualifiers.ApplicationContext
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.Job
import kotlinx.coroutines.flow.MutableSharedFlow
import kotlinx.coroutines.flow.MutableStateFlow
import kotlinx.coroutines.flow.StateFlow
import kotlinx.coroutines.flow.asSharedFlow
import kotlinx.coroutines.flow.asStateFlow
import kotlinx.coroutines.launch
import kotlinx.coroutines.withContext
import com.modocs.core.common.OoxmlDecryptor
import com.modocs.core.common.documentErrorMessage
import kotlinx.coroutines.CancellationException
import javax.inject.Inject

data class XlsxViewerState(
    val isLoading: Boolean = true,
    val fileName: String = "",
    val errorMessage: String? = null,
    val document: XlsxDocument? = null,
    val activeSheetIndex: Int = 0,
    val isEditing: Boolean = false,
    val isDirty: Boolean = false,
    val isSaving: Boolean = false,
    /** Currently editing cell as (rowIndex, colIndex), or null if not editing a cell. */
    val editingCell: Pair<Int, Int>? = null,
    val isPasswordRequired: Boolean = false,
    val passwordError: String? = null,
    /**
     * True once the workbook has been opened by decrypting it. The in-memory
     * document is plaintext, so writing it back over the original would quietly
     * strip the user's password protection.
     */
    val wasPasswordProtected: Boolean = false,
)

data class XlsxSearchState(
    val isSearchActive: Boolean = false,
    val query: String = "",
    val isSearching: Boolean = false,
    val matches: List<XlsxSearchMatch> = emptyList(),
    val currentMatchIndex: Int = -1,
) {
    val totalMatches: Int get() = matches.size
    val hasMatches: Boolean get() = matches.isNotEmpty()
    val currentMatch: XlsxSearchMatch? get() =
        if (currentMatchIndex in matches.indices) matches[currentMatchIndex] else null
}

data class XlsxSearchMatch(
    val sheetIndex: Int,
    val rowIndex: Int,
    val colIndex: Int,
    val snippet: String,
)

@HiltViewModel
class XlsxViewerViewModel @Inject constructor(
    @ApplicationContext private val context: Context,
) : ViewModel() {

    private val _state = MutableStateFlow(XlsxViewerState())
    val state: StateFlow<XlsxViewerState> = _state.asStateFlow()

    private val _searchState = MutableStateFlow(XlsxSearchState())
    val searchState: StateFlow<XlsxSearchState> = _searchState.asStateFlow()

    sealed interface XlsxEvent {
        data class Error(val message: String) : XlsxEvent
        data class SaveSuccess(val message: String) : XlsxEvent
        data class SaveError(val message: String) : XlsxEvent
    }

    private val _events = MutableSharedFlow<XlsxEvent>()
    val events = _events.asSharedFlow()

    private var searchJob: Job? = null
    private var documentUri: Uri? = null

    fun loadXlsx(uri: Uri, displayName: String?) {
        if (_state.value.document != null) return

        documentUri = uri

        viewModelScope.launch {
            _state.value = _state.value.copy(isLoading = true, errorMessage = null)

            val name = displayName ?: resolveFileName(uri) ?: "Spreadsheet"

            try {
                // Check if file is password-protected (OLE2 container)
                if (OoxmlDecryptor.isEncryptedOoxmlFile(context, uri)) {
                    _state.value = _state.value.copy(
                        isLoading = false,
                        isPasswordRequired = true,
                        fileName = name,
                    )
                    return@launch
                }

                val document = XlsxParser.parse(context, uri)
                _state.value = _state.value.copy(
                    isLoading = false,
                    fileName = name,
                    document = document,
                )
            } catch (e: CancellationException) {
                throw e
            } catch (t: Throwable) {
                _state.value = _state.value.copy(
                    isLoading = false,
                    fileName = name,
                    errorMessage = "Failed to open spreadsheet: ${documentErrorMessage(t)}",
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
                        val document = XlsxParser.parse(result.inputStream)
                        _state.value = _state.value.copy(
                            isLoading = false,
                            isPasswordRequired = false,
                            passwordError = null,
                            document = document,
                            wasPasswordProtected = true,
                        )
                    } catch (e: CancellationException) {
                        throw e
                    } catch (t: Throwable) {
                        _state.value = _state.value.copy(
                            isLoading = false,
                            isPasswordRequired = false,
                            errorMessage = "Failed to open spreadsheet: ${documentErrorMessage(t)}",
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

    fun selectSheet(index: Int) {
        val doc = _state.value.document ?: return
        if (index in doc.sheets.indices) {
            _state.value = _state.value.copy(activeSheetIndex = index)
            // Re-run search if active
            if (_searchState.value.isSearchActive && _searchState.value.query.isNotEmpty()) {
                performSearch(_searchState.value.query)
            }
        }
    }

    // --- Search ---

    fun toggleSearch() {
        val current = _searchState.value
        if (current.isSearchActive) {
            _searchState.value = XlsxSearchState()
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
                val doc = _state.value.document ?: return@withContext emptyList()
                val results = mutableListOf<XlsxSearchMatch>()
                val queryLower = query.lowercase()

                for ((sheetIdx, sheet) in doc.sheets.withIndex()) {
                    for (row in sheet.rows) {
                        for (cell in row.cells) {
                            if (cell.value.lowercase().contains(queryLower)) {
                                results.add(
                                    XlsxSearchMatch(
                                        sheetIndex = sheetIdx,
                                        rowIndex = row.rowIndex,
                                        colIndex = cell.columnIndex,
                                        snippet = cell.value,
                                    )
                                )
                            }
                        }
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

    fun nextMatch() {
        val current = _searchState.value
        if (current.matches.isEmpty()) return
        val nextIndex = (current.currentMatchIndex + 1) % current.matches.size
        _searchState.value = current.copy(currentMatchIndex = nextIndex)

        // Switch to the sheet containing the match
        val match = current.matches[nextIndex]
        if (match.sheetIndex != _state.value.activeSheetIndex) {
            _state.value = _state.value.copy(activeSheetIndex = match.sheetIndex)
        }
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

        val match = current.matches[prevIndex]
        if (match.sheetIndex != _state.value.activeSheetIndex) {
            _state.value = _state.value.copy(activeSheetIndex = match.sheetIndex)
        }
    }

    // --- Editing ---

    fun toggleEditMode() {
        val current = _state.value
        _state.value = current.copy(
            isEditing = !current.isEditing,
            editingCell = null,
        )
    }

    fun startEditingCell(rowIndex: Int, colIndex: Int) {
        if (!_state.value.isEditing) return
        _state.value = _state.value.copy(editingCell = rowIndex to colIndex)
    }

    fun stopEditingCell() {
        _state.value = _state.value.copy(editingCell = null)
    }

    fun updateCellValue(rowIndex: Int, colIndex: Int, newValue: String) {
        val document = _state.value.document ?: return
        val sheetIndex = _state.value.activeSheetIndex
        val sheet = document.sheets.getOrNull(sheetIndex) ?: return

        // Find or create the row
        val row = sheet.rows.find { it.rowIndex == rowIndex }
        if (row != null) {
            val cellIdx = row.cells.indexOfFirst { it.columnIndex == colIndex }
            if (cellIdx >= 0) {
                // Update existing cell
                val oldCell = row.cells[cellIdx]
                row.cells[cellIdx] = oldCell.copy(
                    value = newValue,
                    type = inferEditedCellType(newValue),
                    formula = null,
                    rawValue = null,
                )
            } else {
                // Add a new cell to this row
                row.cells.add(
                    XlsxCell(
                        columnIndex = colIndex,
                        value = newValue,
                        type = inferEditedCellType(newValue),
                    )
                )
            }
        } else {
            // Add a new row with the cell
            sheet.rows.add(
                XlsxRow(
                    rowIndex = rowIndex,
                    cells = mutableListOf(
                        XlsxCell(
                            columnIndex = colIndex,
                            value = newValue,
                            type = inferEditedCellType(newValue),
                        )
                    ),
                )
            )
        }

        document.modifiedSheets.add(sheetIndex)
        _state.value = _state.value.copy(isDirty = true)
    }

    fun saveDocument() {
        val document = _state.value.document ?: return
        val uri = documentUri ?: return

        viewModelScope.launch {
            // Writing the decrypted document back over the original would remove
            // its password with no warning — a confidentiality regression the
            // user never asked for. Make them choose a destination instead.
            if (_state.value.wasPasswordProtected) {
                _events.emit(
                    XlsxEvent.SaveError(
                        "This spreadsheet is password-protected. Saving over it would " +
                            "remove the password — use Save As to write an unprotected copy."
                    )
                )
                return@launch
            }

            _state.value = _state.value.copy(isSaving = true)
            try {
                XlsxWriter.save(context, document, uri)
                _state.value = _state.value.copy(isDirty = false, isSaving = false)
                _events.emit(XlsxEvent.SaveSuccess("Spreadsheet saved"))
            } catch (e: Exception) {
                _state.value = _state.value.copy(isSaving = false)
                _events.emit(XlsxEvent.SaveError("Save failed: ${e.message ?: "Unknown error"}"))
            }
        }
    }

    fun saveDocumentAs(outputUri: Uri) {
        val document = _state.value.document ?: return

        viewModelScope.launch {
            _state.value = _state.value.copy(isSaving = true)
            try {
                XlsxWriter.save(context, document, outputUri)
                _state.value = _state.value.copy(isDirty = false, isSaving = false)
                _events.emit(
                    XlsxEvent.SaveSuccess(
                        if (_state.value.wasPasswordProtected) {
                            "Saved — note this copy is not password-protected"
                        } else {
                            "Spreadsheet saved"
                        }
                    )
                )
            } catch (e: Exception) {
                _state.value = _state.value.copy(isSaving = false)
                _events.emit(XlsxEvent.SaveError("Save failed: ${e.message ?: "Unknown error"}"))
            }
        }
    }

    private fun inferEditedCellType(value: String): CellType {
        return if (value.toDoubleOrNull() != null) CellType.NUMBER else CellType.STRING
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
