package com.modocs.feature.home

import android.net.Uri
import androidx.lifecycle.ViewModel
import androidx.lifecycle.viewModelScope
import com.modocs.core.common.DocumentType
import com.modocs.core.common.MAX_FILE_SIZE_BYTES
import com.modocs.core.model.RecentFile
import com.modocs.core.storage.FileAccessManager
import com.modocs.core.storage.RecentFilesRepository
import dagger.hilt.android.lifecycle.HiltViewModel
import kotlinx.coroutines.flow.MutableSharedFlow
import kotlinx.coroutines.flow.MutableStateFlow
import kotlinx.coroutines.flow.SharingStarted
import kotlinx.coroutines.flow.StateFlow
import kotlinx.coroutines.flow.asSharedFlow
import kotlinx.coroutines.flow.stateIn
import java.io.IOException
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.withContext
import kotlinx.coroutines.launch
import javax.inject.Inject

sealed interface HomeEvent {
    data class OpenDocument(val uri: Uri, val documentType: DocumentType, val displayName: String?) : HomeEvent
    data class ShowError(val message: String) : HomeEvent
}

@HiltViewModel
class HomeViewModel @Inject constructor(
    private val recentFilesRepository: RecentFilesRepository,
    private val fileAccessManager: FileAccessManager,
) : ViewModel() {

    val recentFiles: StateFlow<List<RecentFile>> = recentFilesRepository
        .getRecentFiles()
        .stateIn(viewModelScope, SharingStarted.WhileSubscribed(5000), emptyList())

    private val _events = MutableSharedFlow<HomeEvent>()
    val events = _events.asSharedFlow()

    fun onFileSelected(uri: Uri, displayName: String?) {
        viewModelScope.launch {
            if (!fileAccessManager.validateFileSize(uri)) {
                val maxMb = MAX_FILE_SIZE_BYTES / (1024 * 1024)
                _events.emit(HomeEvent.ShowError("File exceeds the ${maxMb}MB size limit"))
                return@launch
            }

            fileAccessManager.persistPermission(uri)

            val mimeType = fileAccessManager.getMimeType(uri)
            val name = displayName ?: "Unknown"
            val docType = DocumentType.fromMimeType(mimeType)
                .takeIf { it != DocumentType.UNKNOWN }
                ?: DocumentType.fromFileName(name)

            if (docType == DocumentType.UNKNOWN) {
                _events.emit(HomeEvent.ShowError("Unsupported file format"))
                return@launch
            }

            _events.emit(HomeEvent.OpenDocument(uri, docType, name))
        }
    }

    fun onRecentFileClicked(recentFile: RecentFile) {
        val uri = Uri.parse(recentFile.uri)
        viewModelScope.launch {
            val event = withContext(Dispatchers.IO) {
                try {
                    if (!fileAccessManager.validateFileSize(uri)) {
                        HomeEvent.ShowError("File exceeds the ${MAX_FILE_SIZE_BYTES / (1024 * 1024)}MB size limit")
                    } else {
                        fileAccessManager.openInputStream(uri)?.use {
                            HomeEvent.OpenDocument(uri, recentFile.documentType, recentFile.displayName)
                        }
                    }
                } catch (_: SecurityException) {
                    null
                } catch (_: IOException) {
                    null
                }
            }
            _events.emit(event ?: HomeEvent.ShowError("File unavailable. Use Open File to select it again."))
        }
    }
}
