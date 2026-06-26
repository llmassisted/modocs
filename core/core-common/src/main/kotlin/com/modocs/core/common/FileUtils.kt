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

enum class DocumentType(val displayName: String) {
    PDF("PDF"),
    DOCX("Word Document"),
    XLSX("Excel Spreadsheet"),
    PPTX("PowerPoint Presentation"),
    UNKNOWN("Unknown");

    companion object {
        fun fromMimeType(mimeType: String?): DocumentType = when (mimeType) {
            "application/pdf" -> PDF
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document" -> DOCX
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" -> XLSX
            "application/vnd.openxmlformats-officedocument.presentationml.presentation" -> PPTX
            else -> UNKNOWN
        }

        fun fromFileName(name: String): DocumentType {
            val ext = name.substringAfterLast('.', "").lowercase()
            return when (ext) {
                "pdf" -> PDF
                "docx" -> DOCX
                "xlsx" -> XLSX
                "pptx" -> PPTX
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
