package com.modocs.feature.home

import androidx.lifecycle.ViewModel
import androidx.lifecycle.viewModelScope
import dagger.hilt.android.lifecycle.HiltViewModel
import kotlinx.coroutines.flow.MutableStateFlow
import kotlinx.coroutines.flow.StateFlow
import kotlinx.coroutines.flow.asStateFlow
import kotlinx.coroutines.flow.update
import kotlinx.coroutines.launch
import java.io.File
import javax.inject.Inject

enum class UpdateStatus {
    Idle,
    Checking,
    UpdateAvailable,
    UpToDate,
    Downloading,
    ReadyToInstall,
    Error,
}

data class UpdateState(
    val status: UpdateStatus = UpdateStatus.Idle,
    val currentVersion: String = "",
    val latestVersion: String? = null,
    val releaseNotes: String? = null,
    val downloadProgress: Float = 0f,
    val errorMessage: String? = null,
)

@HiltViewModel
class SettingsViewModel @Inject constructor(
    private val updateChecker: AppUpdateChecker,
) : ViewModel() {

    private val _updateState = MutableStateFlow(
        UpdateState(currentVersion = updateChecker.getCurrentVersion()),
    )
    val updateState: StateFlow<UpdateState> = _updateState.asStateFlow()

    private var latestRelease: AppUpdateChecker.ReleaseInfo? = null
    private var downloadedApk: File? = null

    fun checkForUpdate() {
        viewModelScope.launch {
            _updateState.update { it.copy(status = UpdateStatus.Checking, errorMessage = null) }

            updateChecker.checkForUpdate()
                .onSuccess { releaseInfo ->
                    if (releaseInfo != null) {
                        latestRelease = releaseInfo
                        _updateState.update {
                            it.copy(
                                status = UpdateStatus.UpdateAvailable,
                                latestVersion = releaseInfo.versionName,
                                releaseNotes = releaseInfo.releaseNotes.ifBlank { null },
                            )
                        }
                    } else {
                        _updateState.update { it.copy(status = UpdateStatus.UpToDate) }
                    }
                }
                .onFailure { e ->
                    _updateState.update {
                        it.copy(
                            status = UpdateStatus.Error,
                            errorMessage = e.message ?: "Failed to check for updates",
                        )
                    }
                }
        }
    }

    fun downloadUpdate() {
        val release = latestRelease ?: return

        viewModelScope.launch {
            _updateState.update {
                it.copy(status = UpdateStatus.Downloading, downloadProgress = 0f)
            }

            updateChecker.downloadApk(
                downloadUrl = release.downloadUrl,
                versionName = release.versionName,
                onProgress = { progress ->
                    _updateState.update { it.copy(downloadProgress = progress) }
                },
            ).onSuccess { apkFile ->
                downloadedApk = apkFile
                _updateState.update { it.copy(status = UpdateStatus.ReadyToInstall) }
            }.onFailure { e ->
                _updateState.update {
                    it.copy(
                        status = UpdateStatus.Error,
                        errorMessage = e.message ?: "Download failed",
                    )
                }
            }
        }
    }

    fun installUpdate() {
        downloadedApk?.let { updateChecker.installApk(it) }
    }
}
