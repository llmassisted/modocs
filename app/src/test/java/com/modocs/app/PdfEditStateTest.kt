package com.modocs.app

import com.modocs.feature.pdf.CheckmarkAnnotation
import com.modocs.feature.pdf.FillSignState
import org.junit.Assert.*
import org.junit.Test

class PdfEditStateTest {
    @Test fun editsAreComparedWithTheLastSavedCopyIncludingFormFields() {
        val mark = CheckmarkAnnotation("mark", 0, .2f, .4f, isCheck = true)
        val saved = FillSignState(annotations = listOf(mark), savedAnnotations = listOf(mark))
        assertFalse(saved.isDirty)
        assertTrue(saved.copy(annotations = emptyList()).isDirty)
        val changedField = saved.copy(formValues = mapOf("name" to "Alice"))
        assertTrue(changedField.isDirty)
        assertFalse(changedField.copy(savedFormValues = changedField.formValues).isDirty)
    }
}
