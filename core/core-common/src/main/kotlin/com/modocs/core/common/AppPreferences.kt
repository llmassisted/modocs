package com.modocs.core.common

import android.content.Context
import kotlinx.coroutines.flow.MutableStateFlow
import kotlinx.coroutines.flow.asStateFlow
import java.security.MessageDigest

data class ReaderPreferences(
    val theme: String = "System", val dynamicColor: Boolean = true,
    val keepAwake: Boolean = false, val rememberPosition: Boolean = true,
    val spreadsheetZoom: Float = 1f, val keepRecents: Boolean = true,
)

class AppPreferences private constructor(context: Context) {
    private val prefs = context.getSharedPreferences("reader_settings", Context.MODE_PRIVATE)
    private val positions = context.getSharedPreferences("reading_positions", Context.MODE_PRIVATE)
    private fun read() = ReaderPreferences(prefs.getString("theme", "System") ?: "System",
        prefs.getBoolean("dynamicColor", true), prefs.getBoolean("keepAwake", false),
        prefs.getBoolean("rememberPosition", true), prefs.getFloat("spreadsheetZoom", 1f), prefs.getBoolean("keepRecents", true))
    private val _state = MutableStateFlow(read())
    val state = _state.asStateFlow()
    fun update(value: ReaderPreferences) {
        prefs.edit().putString("theme", value.theme).putBoolean("dynamicColor", value.dynamicColor)
            .putBoolean("keepAwake", value.keepAwake).putBoolean("rememberPosition", value.rememberPosition)
            .putFloat("spreadsheetZoom", value.spreadsheetZoom).putBoolean("keepRecents", value.keepRecents).apply()
        if (!value.rememberPosition) positions.edit().clear().apply()
        _state.value = value
    }
    private fun key(document: String) = MessageDigest.getInstance("SHA-256").digest(document.toByteArray()).joinToString("") { "%02x".format(it) }
    fun position(document: String): Int = if (state.value.rememberPosition) positions.getInt(key(document), 0) else 0
    fun savePosition(document: String, index: Int) {
        if (state.value.rememberPosition) positions.edit().putInt(key(document), index.coerceAtLeast(0)).apply()
    }
    fun clearPositions() = positions.edit().clear().apply()
    companion object {
        @Volatile private var instance: AppPreferences? = null
        fun get(context: Context): AppPreferences = instance ?: synchronized(this) {
            instance ?: AppPreferences(context.applicationContext).also { instance = it }
        }
    }
}
