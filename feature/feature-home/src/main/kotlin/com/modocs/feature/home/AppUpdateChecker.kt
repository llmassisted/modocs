package com.modocs.feature.home

import android.content.Context
import android.content.Intent
import androidx.core.content.FileProvider
import dagger.hilt.android.qualifiers.ApplicationContext
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.withContext
import org.json.JSONObject
import java.io.File
import java.net.HttpURLConnection
import java.net.URL
import javax.inject.Inject
import javax.inject.Singleton

@Singleton
class AppUpdateChecker @Inject constructor(
    @ApplicationContext private val context: Context,
) {
    data class ReleaseInfo(
        val versionName: String,
        val downloadUrl: String,
        val releaseNotes: String,
    )

    fun getCurrentVersion(): String {
        return try {
            context.packageManager.getPackageInfo(context.packageName, 0).versionName ?: ""
        } catch (_: Exception) {
            ""
        }
    }

    suspend fun checkForUpdate(): Result<ReleaseInfo?> = withContext(Dispatchers.IO) {
        try {
            val url = URL(RELEASES_URL)
            val connection = url.openConnection() as HttpURLConnection
            connection.requestMethod = "GET"
            connection.setRequestProperty("Accept", "application/vnd.github+json")
            connection.connectTimeout = 15_000
            connection.readTimeout = 15_000

            try {
                val responseCode = connection.responseCode
                if (responseCode != HttpURLConnection.HTTP_OK) {
                    return@withContext Result.failure(Exception("GitHub API returned $responseCode"))
                }

                val json = connection.inputStream.bufferedReader().use { it.readText() }
                val release = JSONObject(json)

                val tagName = release.getString("tag_name")
                val body = release.optString("body", "")
                val assets = release.getJSONArray("assets")

                var apkUrl: String? = null
                for (i in 0 until assets.length()) {
                    val asset = assets.getJSONObject(i)
                    if (asset.getString("name").endsWith(".apk")) {
                        apkUrl = asset.getString("browser_download_url")
                        break
                    }
                }

                if (apkUrl == null) {
                    return@withContext Result.failure(Exception("No APK found in latest release"))
                }

                val remoteVersion = tagName.removePrefix("v")
                val currentVersion = getCurrentVersion()

                if (isNewerVersion(remoteVersion, currentVersion)) {
                    Result.success(ReleaseInfo(remoteVersion, apkUrl, body))
                } else {
                    Result.success(null)
                }
            } finally {
                connection.disconnect()
            }
        } catch (e: Exception) {
            Result.failure(e)
        }
    }

    suspend fun downloadApk(
        downloadUrl: String,
        versionName: String,
        onProgress: (Float) -> Unit,
    ): Result<File> = withContext(Dispatchers.IO) {
        val updatesDir = File(context.cacheDir, "updates")
        updatesDir.mkdirs()
        val apkFile = File(updatesDir, "modocs-v$versionName.apk")

        try {
            val url = URL(downloadUrl)
            val connection = url.openConnection() as HttpURLConnection
            connection.connectTimeout = 15_000
            connection.readTimeout = 60_000

            try {
                val totalBytes = connection.contentLength.toLong()
                var downloadedBytes = 0L
                var lastReportedPercent = -1

                connection.inputStream.use { input ->
                    apkFile.outputStream().use { output ->
                        val buffer = ByteArray(8192)
                        var bytesRead: Int
                        while (input.read(buffer).also { bytesRead = it } != -1) {
                            output.write(buffer, 0, bytesRead)
                            downloadedBytes += bytesRead
                            if (totalBytes > 0) {
                                val percent = (downloadedBytes * 100 / totalBytes).toInt()
                                if (percent > lastReportedPercent) {
                                    lastReportedPercent = percent
                                    onProgress(downloadedBytes.toFloat() / totalBytes)
                                }
                            }
                        }
                    }
                }
            } finally {
                connection.disconnect()
            }

            Result.success(apkFile)
        } catch (e: Exception) {
            apkFile.delete()
            Result.failure(e)
        }
    }

    fun installApk(apkFile: File) {
        val uri = FileProvider.getUriForFile(
            context,
            "${context.packageName}.fileprovider",
            apkFile,
        )
        val intent = Intent(Intent.ACTION_VIEW).apply {
            setDataAndType(uri, "application/vnd.android.package-archive")
            addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
            addFlags(Intent.FLAG_ACTIVITY_NEW_TASK)
        }
        context.startActivity(intent)
    }

    companion object {
        private const val RELEASES_URL =
            "https://api.github.com/repos/llmassisted/modocs/releases/latest"

        fun isNewerVersion(remote: String, current: String): Boolean {
            val remoteNum = remote.toDoubleOrNull()
            val currentNum = current.toDoubleOrNull()
            if (remoteNum != null && currentNum != null) {
                return remoteNum > currentNum
            }
            val remoteParts = remote.split(".").mapNotNull { it.toIntOrNull() }
            val currentParts = current.split(".").mapNotNull { it.toIntOrNull() }
            val maxLen = maxOf(remoteParts.size, currentParts.size)
            for (i in 0 until maxLen) {
                val r = remoteParts.getOrElse(i) { 0 }
                val c = currentParts.getOrElse(i) { 0 }
                if (r > c) return true
                if (r < c) return false
            }
            return false
        }
    }
}
