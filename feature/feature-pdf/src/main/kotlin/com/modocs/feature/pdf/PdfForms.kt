package com.modocs.feature.pdf

import android.content.Context
import android.net.Uri
import com.tom_roush.pdfbox.android.PDFBoxResourceLoader
import com.tom_roush.pdfbox.pdmodel.PDDocument
import com.tom_roush.pdfbox.pdmodel.interactive.form.*
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.withContext

data class PdfFormField(val name: String, val label: String, val value: String,
    val checkbox: Boolean = false, val options: List<String> = emptyList())

object PdfForms {
    suspend fun read(context: Context, uri: Uri): List<PdfFormField> = withContext(Dispatchers.IO) {
        PDFBoxResourceLoader.init(context)
        context.contentResolver.openInputStream(uri)?.use { input ->
            PDDocument.load(input).use { doc ->
                doc.documentCatalog.acroForm?.fieldTree?.mapNotNull { field ->
                    if (field.isReadOnly || field !is PDTextField && field !is PDCheckBox && field !is PDChoice) null
                    else if (field is PDChoice && field.isMultiSelect) null
                    else PdfFormField(field.fullyQualifiedName, field.alternateFieldName ?: field.fullyQualifiedName,
                        if (field is PDCheckBox) field.isChecked.toString() else field.valueAsString,
                        checkbox = field is PDCheckBox,
                        options = if (field is PDChoice) field.optionsExportValues else emptyList())
                }.orEmpty()
            }
        } ?: error("Cannot read PDF form fields")
    }

    fun applyValues(document: PDDocument, values: Map<String, String>) {
        if (values.isEmpty()) return
        val form = requireNotNull(document.documentCatalog.acroForm) { "The PDF form is no longer available" }
        for ((name, value) in values) {
            val field = requireNotNull(form.getField(name)) { "Missing form field: $name" }
            require(!field.isReadOnly) { "Read-only form field: $name" }
            when (field) {
                is PDCheckBox -> if (value == "true") field.check() else field.unCheck()
                is PDTextField -> field.setValue(value)
                is PDChoice -> field.setValue(value)
                else -> error("Unsupported form field: $name")
            }
        }
        form.refreshAppearances()
    }
}
