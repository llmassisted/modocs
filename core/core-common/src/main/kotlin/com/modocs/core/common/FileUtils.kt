package com.modocs.core.common

import android.content.Context
import android.net.Uri
import android.provider.OpenableColumns

/**
 * Maximum file size allowed (100 MB). Applied to both the in-app picker and
 * documents opened from external apps. Note this is the *compressed* on-disk
 * size; OOXML expansion is separately bounded by MAX_UNCOMPRESSED_BYTES.
 */
const val MAX_FILE_SIZE_BYTES = 100L * 1024 * 1024

enum class DocumentType(
    val displayName: String,
    val mimeType: String,
    /** Other MIME types that apps and document providers report for this format. */
    val alternateMimeTypes: List<String> = emptyList(),
) {
    PDF("PDF", "application/pdf"),
    DOCX("Word Document", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"),
    XLSX("Excel Spreadsheet", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
    PPTX("PowerPoint Presentation", "application/vnd.openxmlformats-officedocument.presentationml.presentation"),
    // Older Android versions map .csv to text/comma-separated-values, and some apps send the others.
    CSV("CSV File", "text/csv", listOf("text/comma-separated-values", "application/csv", "text/x-csv")),
    UNKNOWN("Unknown", "application/octet-stream");

    companion object {
        fun fromMimeType(mimeType: String?): DocumentType {
            val type = mimeType?.substringBefore(';')?.trim()?.lowercase()
            return entries.firstOrNull { it != UNKNOWN && (it.mimeType == type || type in it.alternateMimeTypes) } ?: UNKNOWN
        }

        fun fromFileName(name: String): DocumentType {
            val ext = name.substringAfterLast('.', "").lowercase()
            return when (ext) {
                "pdf" -> PDF
                "docx" -> DOCX
                "xlsx" -> XLSX
                "pptx" -> PPTX
                "csv" -> CSV
                else -> UNKNOWN
            }
        }
    }
}

fun Uri.getFileName(context: Context): String? {
    val cursor = context.contentResolver.query(this, null, null, null, null) ?: return null
    return cursor.use {
        if (it.moveToFirst()) {
            val nameIndex = it.getColumnIndex(OpenableColumns.DISPLAY_NAME)
            if (nameIndex >= 0) it.getString(nameIndex) else null
        } else {
            null
        }
    }
}

fun Uri.getFileSize(context: Context): Long? {
    val cursor = context.contentResolver.query(this, null, null, null, null) ?: return null
    return cursor.use {
        if (it.moveToFirst()) {
            val sizeIndex = it.getColumnIndex(OpenableColumns.SIZE)
            if (sizeIndex >= 0) it.getLong(sizeIndex) else null
        } else {
            null
        }
    }
}
