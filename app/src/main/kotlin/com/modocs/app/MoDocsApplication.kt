package com.modocs.app

import android.app.Application
import android.util.Log
import com.modocs.core.common.DECRYPTED_TEMP_PREFIX
import dagger.hilt.android.HiltAndroidApp
import kotlinx.coroutines.CoroutineScope
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.SupervisorJob
import kotlinx.coroutines.launch

@HiltAndroidApp
class MoDocsApplication : Application() {

    override fun onCreate() {
        super.onCreate()
        sweepDecryptedTempFiles()
    }

    /**
     * Delete plaintext copies of password-protected PDFs left behind in the cache.
     *
     * The viewer deletes its temp file in `onCleared()`, but that never runs if
     * the process is killed, so decrypted copies of protected documents could
     * otherwise sit in the cache until Android decided to trim it.
     *
     * `Application.onCreate` runs before any Activity, so nothing can be holding
     * one of these files open at this point and deleting them all is safe.
     */
    private fun sweepDecryptedTempFiles() {
        CoroutineScope(SupervisorJob() + Dispatchers.IO).launch {
            try {
                val stale = cacheDir.listFiles { file ->
                    file.isFile && file.name.startsWith(DECRYPTED_TEMP_PREFIX)
                } ?: return@launch

                for (file in stale) {
                    if (!file.delete()) {
                        Log.w(TAG, "Could not delete stale decrypted file ${file.name}")
                    }
                }
            } catch (e: Exception) {
                Log.w(TAG, "Decrypted-cache sweep failed: ${e.message}")
            }
        }
    }

    private companion object {
        const val TAG = "MoDocsApplication"
    }
}
