package com.modocs.feature.pptx

import android.content.Context
import android.graphics.Bitmap
import android.graphics.BitmapFactory
import android.graphics.Canvas
import android.graphics.Paint
import android.graphics.RectF
import android.graphics.Typeface
import android.net.Uri
import android.provider.OpenableColumns
import android.text.Layout
import android.text.SpannableStringBuilder
import android.text.Spanned
import android.text.StaticLayout
import android.text.TextPaint
import android.text.style.AbsoluteSizeSpan
import android.text.style.ForegroundColorSpan
import android.text.style.StyleSpan
import android.text.style.UnderlineSpan
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
import kotlinx.coroutines.sync.Mutex
import kotlinx.coroutines.sync.withLock
import kotlinx.coroutines.withContext
import com.modocs.core.common.OoxmlDecryptor
import javax.inject.Inject

data class PptxViewerState(
    val isLoading: Boolean = true,
    val fileName: String = "",
    val errorMessage: String? = null,
    val document: PptxDocument? = null,
    val currentSlide: Int = 0,
    val slideCount: Int = 0,
    val isPasswordRequired: Boolean = false,
    val passwordError: String? = null,
)

data class PptxSearchState(
    val isSearchActive: Boolean = false,
    val query: String = "",
    val isSearching: Boolean = false,
    val matches: List<PptxSearchMatch> = emptyList(),
    val currentMatchIndex: Int = -1,
) {
    val totalMatches: Int get() = matches.size
    val hasMatches: Boolean get() = matches.isNotEmpty()
    val currentMatch: PptxSearchMatch? get() =
        if (currentMatchIndex in matches.indices) matches[currentMatchIndex] else null
}

data class PptxSearchMatch(
    val slideIndex: Int,
    val shapeIndex: Int,
    val snippet: String,
)

@HiltViewModel
class PptxViewerViewModel @Inject constructor(
    @ApplicationContext private val context: Context,
) : ViewModel() {

    private val _state = MutableStateFlow(PptxViewerState())
    val state: StateFlow<PptxViewerState> = _state.asStateFlow()

    private val _searchState = MutableStateFlow(PptxSearchState())
    val searchState: StateFlow<PptxSearchState> = _searchState.asStateFlow()

    sealed interface PptxEvent {
        data class Error(val message: String) : PptxEvent
    }

    private val _events = MutableSharedFlow<PptxEvent>()
    val events = _events.asSharedFlow()

    private var searchJob: Job? = null
    private var documentUri: Uri? = null
    private data class SlideBitmapCacheKey(val slideIndex: Int, val targetWidth: Int)

    private val slideBitmapCache = LinkedHashMap<SlideBitmapCacheKey, Bitmap>(8, 0.75f, true)
    private val renderMutex = Mutex()

    fun loadPptx(uri: Uri, displayName: String?) {
        if (_state.value.document != null) return

        documentUri = uri

        viewModelScope.launch {
            _state.value = _state.value.copy(isLoading = true, errorMessage = null)

            val name = displayName ?: resolveFileName(uri) ?: "Presentation"

            // Check if file is password-protected (OLE2 container)
            if (OoxmlDecryptor.isEncryptedOoxmlFile(context, uri)) {
                _state.value = _state.value.copy(
                    isLoading = false,
                    isPasswordRequired = true,
                    fileName = name,
                )
                return@launch
            }

            try {
                val document = PptxParser.parse(context, uri)
                _state.value = _state.value.copy(
                    isLoading = false,
                    fileName = name,
                    document = document,
                    slideCount = document.slides.size,
                )
            } catch (e: Exception) {
                _state.value = _state.value.copy(
                    isLoading = false,
                    errorMessage = "Failed to open presentation: ${e.message ?: "Unknown error"}",
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
                        val document = PptxParser.parse(result.inputStream)
                        _state.value = _state.value.copy(
                            isLoading = false,
                            isPasswordRequired = false,
                            passwordError = null,
                            document = document,
                            slideCount = document.slides.size,
                        )
                    } catch (e: Exception) {
                        _state.value = _state.value.copy(
                            isLoading = false,
                            isPasswordRequired = false,
                            errorMessage = "Failed to open presentation: ${e.message ?: "Unknown error"}",
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

    fun goToSlide(index: Int) {
        val doc = _state.value.document ?: return
        if (index in doc.slides.indices) {
            _state.value = _state.value.copy(currentSlide = index)
        }
    }

    fun nextSlide() {
        val current = _state.value
        if (current.currentSlide < current.slideCount - 1) {
            _state.value = current.copy(currentSlide = current.currentSlide + 1)
        }
    }

    fun previousSlide() {
        val current = _state.value
        if (current.currentSlide > 0) {
            _state.value = current.copy(currentSlide = current.currentSlide - 1)
        }
    }

    suspend fun renderSlide(slideIndex: Int, targetWidth: Int): Bitmap? = renderMutex.withLock {
        withContext(Dispatchers.Default) {
            val doc = _state.value.document ?: return@withContext null
            val slide = doc.slides.getOrNull(slideIndex) ?: return@withContext null

            val cacheKey = SlideBitmapCacheKey(slideIndex, targetWidth)
            slideBitmapCache[cacheKey]?.let { return@withContext it }

            val aspectRatio = doc.slideHeight.toFloat() / doc.slideWidth.toFloat()
            val targetHeight = (targetWidth * aspectRatio).toInt()

            val bitmap = Bitmap.createBitmap(targetWidth, targetHeight, Bitmap.Config.ARGB_8888)
            val canvas = Canvas(bitmap)

            val scaleX = targetWidth.toFloat() / doc.slideWidth.toFloat()
            val scaleY = targetHeight.toFloat() / doc.slideHeight.toFloat()

            // Background
            val bgColor = slide.backgroundColor ?: android.graphics.Color.WHITE
            canvas.drawColor(bgColor)

            // Render background shapes (from layout/master) first
            for (shape in slide.backgroundShapes) {
                renderShape(canvas, shape, slide, scaleX, scaleY)
            }

            // Render slide shapes
            for (shape in slide.shapes) {
                renderShape(canvas, shape, slide, scaleX, scaleY)
            }

            slideBitmapCache[cacheKey] = bitmap
            trimSlideBitmapCache()
            bitmap
        }
    }

    private fun trimSlideBitmapCache() {
        while (slideBitmapCache.size > MAX_SLIDE_BITMAP_CACHE_SIZE) {
            val oldestKey = slideBitmapCache.keys.first()
            slideBitmapCache.remove(oldestKey)
        }
    }

    private fun renderShape(
        canvas: Canvas,
        shape: PptxShape,
        slide: PptxSlide,
        scaleX: Float,
        scaleY: Float,
    ) {
        when (shape) {
            is PptxTextBox -> renderTextBox(canvas, shape, scaleX, scaleY)
            is PptxImageShape -> renderImage(canvas, shape, slide.images, scaleX, scaleY)
            is PptxRectangle -> renderRectangle(canvas, shape, scaleX, scaleY)
            is PptxLine -> renderLine(canvas, shape, scaleX, scaleY)
        }
    }

    private fun renderTextBox(
        canvas: Canvas,
        textBox: PptxTextBox,
        scaleX: Float,
        scaleY: Float,
    ) {
        val left = textBox.x * scaleX
        val top = textBox.y * scaleY
        val width = textBox.width * scaleX
        val height = textBox.height * scaleY

        canvas.save()
        if (textBox.rotation != 0f) {
            canvas.rotate(textBox.rotation, left + width / 2, top + height / 2)
        }

        // Fill
        if (textBox.fillColor != null) {
            val paint = Paint().apply {
                color = textBox.fillColor
                style = Paint.Style.FILL
            }
            canvas.drawRect(left, top, left + width, top + height, paint)
        }

        // Border
        if (textBox.borderColor != null && textBox.borderWidthEmu > 0) {
            val paint = Paint().apply {
                color = textBox.borderColor
                style = Paint.Style.STROKE
                strokeWidth = (textBox.borderWidthEmu * scaleX).coerceIn(1f, 10f)
                isAntiAlias = true
            }
            canvas.drawRect(left, top, left + width, top + height, paint)
        }

        // Text — padding in EMU: 91440 = 0.1 inch
        val paddingX = 91440f * scaleX
        val paddingY = 45720f * scaleY
        val textLeft = left + paddingX
        val textWidth = (width - paddingX * 2).coerceAtLeast(1f)
        val textY = top + paddingY

        renderParagraphs(canvas, textBox.paragraphs, textLeft, textY, textWidth, scaleX, scaleY)

        canvas.restore()
    }

    /**
     * Render paragraphs using SpannableStringBuilder for proper multi-run text flow.
     * Each paragraph becomes one StaticLayout with spans for formatting.
     */
    private fun renderParagraphs(
        canvas: Canvas,
        paragraphs: List<PptxParagraph>,
        textLeft: Float,
        startY: Float,
        textWidth: Float,
        scaleX: Float,
        scaleY: Float,
    ) {
        var textY = startY

        for (paragraph in paragraphs) {
            // Spacing before
            textY += paragraph.spacingBeforePt * 12700f * scaleY

            if (paragraph.runs.isEmpty()) {
                // Empty paragraph — advance by roughly one line height
                textY += 14f * 12700f * scaleY
                continue
            }

            // Build spannable with all runs combined
            val ssb = SpannableStringBuilder()

            // Add bullet prefix
            val bulletPrefix = paragraph.bulletChar?.let { "$it " } ?: ""
            if (bulletPrefix.isNotEmpty()) {
                ssb.append(bulletPrefix)
            }

            var maxFontSizePx = 14f * 12700f * scaleY // reasonable default

            for (run in paragraph.runs) {
                val start = ssb.length
                ssb.append(run.text)
                val end = ssb.length

                if (start == end) continue

                val fontSizePt = run.fontSizePt ?: 18f
                val fontSizePx = (fontSizePt * 12700f * scaleY).coerceIn(6f, 500f)
                maxFontSizePx = maxOf(maxFontSizePx, fontSizePx)

                // Size span
                ssb.setSpan(
                    AbsoluteSizeSpan(fontSizePx.toInt(), false),
                    start, end, Spanned.SPAN_EXCLUSIVE_EXCLUSIVE,
                )

                // Color span
                val color = run.fontColor ?: android.graphics.Color.BLACK
                ssb.setSpan(
                    ForegroundColorSpan(color),
                    start, end, Spanned.SPAN_EXCLUSIVE_EXCLUSIVE,
                )

                // Style span (bold/italic)
                val style = when {
                    run.bold && run.italic -> Typeface.BOLD_ITALIC
                    run.bold -> Typeface.BOLD
                    run.italic -> Typeface.ITALIC
                    else -> Typeface.NORMAL
                }
                if (style != Typeface.NORMAL) {
                    ssb.setSpan(
                        StyleSpan(style),
                        start, end, Spanned.SPAN_EXCLUSIVE_EXCLUSIVE,
                    )
                }

                // Underline span
                if (run.underline) {
                    ssb.setSpan(
                        UnderlineSpan(),
                        start, end, Spanned.SPAN_EXCLUSIVE_EXCLUSIVE,
                    )
                }
            }

            if (ssb.isEmpty()) continue

            val indent = paragraph.level * 457200f * scaleX
            val availWidth = (textWidth - indent).coerceAtLeast(1f).toInt()

            val paint = TextPaint().apply {
                isAntiAlias = true
                textSize = maxFontSizePx
            }

            val alignment = when (paragraph.alignment) {
                PptxAlignment.CENTER -> Layout.Alignment.ALIGN_CENTER
                PptxAlignment.RIGHT -> Layout.Alignment.ALIGN_OPPOSITE
                else -> Layout.Alignment.ALIGN_NORMAL
            }

            @Suppress("DEPRECATION")
            val layout = StaticLayout(
                ssb, paint, availWidth, alignment, 1.15f, 0f, false,
            )

            canvas.save()
            canvas.translate(textLeft + indent, textY)
            layout.draw(canvas)
            canvas.restore()

            textY += layout.height

            // Spacing after paragraph
            textY += paragraph.spacingAfterPt * 12700f * scaleY
        }
    }

    private fun renderImage(
        canvas: Canvas,
        shape: PptxImageShape,
        images: Map<String, ByteArray>,
        scaleX: Float,
        scaleY: Float,
    ) {
        val imageData = images[shape.relationId] ?: return
        val bitmap = try {
            BitmapFactory.decodeByteArray(imageData, 0, imageData.size)
        } catch (_: Exception) {
            return
        } ?: return

        val left = shape.x * scaleX
        val top = shape.y * scaleY
        val width = shape.width * scaleX
        val height = shape.height * scaleY

        canvas.save()
        if (shape.rotation != 0f) {
            canvas.rotate(shape.rotation, left + width / 2, top + height / 2)
        }

        val destRect = RectF(left, top, left + width, top + height)
        canvas.drawBitmap(bitmap, null, destRect, null)
        bitmap.recycle()

        canvas.restore()
    }

    private fun renderRectangle(
        canvas: Canvas,
        rect: PptxRectangle,
        scaleX: Float,
        scaleY: Float,
    ) {
        val left = rect.x * scaleX
        val top = rect.y * scaleY
        val width = rect.width * scaleX
        val height = rect.height * scaleY

        canvas.save()
        if (rect.rotation != 0f) {
            canvas.rotate(rect.rotation, left + width / 2, top + height / 2)
        }

        if (rect.fillColor != null) {
            val paint = Paint().apply {
                color = rect.fillColor
                style = Paint.Style.FILL
            }
            canvas.drawRect(left, top, left + width, top + height, paint)
        }

        if (rect.borderColor != null && rect.borderWidthEmu > 0) {
            val paint = Paint().apply {
                color = rect.borderColor
                style = Paint.Style.STROKE
                strokeWidth = (rect.borderWidthEmu * scaleX).coerceAtLeast(1f)
            }
            canvas.drawRect(left, top, left + width, top + height, paint)
        }

        // Render text in rectangle if present
        if (rect.text.isNotEmpty()) {
            val paddingX = 91440f * scaleX
            val paddingY = 45720f * scaleY
            val textLeft = left + paddingX
            val textWidth = (width - paddingX * 2).coerceAtLeast(1f)
            val textY = top + paddingY
            renderParagraphs(canvas, rect.text, textLeft, textY, textWidth, scaleX, scaleY)
        }

        canvas.restore()
    }

    private fun renderLine(
        canvas: Canvas,
        line: PptxLine,
        scaleX: Float,
        scaleY: Float,
    ) {
        val x1 = line.x * scaleX
        val y1 = line.y * scaleY
        val x2 = (line.x + line.width) * scaleX
        val y2 = (line.y + line.height) * scaleY

        val paint = Paint().apply {
            color = line.color
            style = Paint.Style.STROKE
            strokeWidth = (line.lineWidthEmu * scaleX).coerceAtLeast(1f)
            isAntiAlias = true
        }
        canvas.drawLine(x1, y1, x2, y2, paint)
    }

    // --- Search ---

    fun toggleSearch() {
        val current = _searchState.value
        if (current.isSearchActive) {
            _searchState.value = PptxSearchState()
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
                val results = mutableListOf<PptxSearchMatch>()
                val queryLower = query.lowercase()

                for ((slideIdx, slide) in doc.slides.withIndex()) {
                    for ((shapeIdx, shape) in slide.shapes.withIndex()) {
                        val text = when (shape) {
                            is PptxTextBox -> shape.paragraphs.joinToString(" ") { it.text }
                            is PptxRectangle -> shape.text.joinToString(" ") { it.text }
                            else -> continue
                        }
                        if (text.lowercase().contains(queryLower)) {
                            val snippetStart = text.lowercase().indexOf(queryLower)
                            val start = (snippetStart - 20).coerceAtLeast(0)
                            val end = (snippetStart + query.length + 20).coerceAtMost(text.length)
                            val prefix = if (start > 0) "..." else ""
                            val suffix = if (end < text.length) "..." else ""
                            results.add(
                                PptxSearchMatch(
                                    slideIndex = slideIdx,
                                    shapeIndex = shapeIdx,
                                    snippet = prefix + text.substring(start, end) + suffix,
                                )
                            )
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

        val match = current.matches[nextIndex]
        _state.value = _state.value.copy(currentSlide = match.slideIndex)
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
        _state.value = _state.value.copy(currentSlide = match.slideIndex)
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
        slideBitmapCache.values.forEach { it.recycle() }
        slideBitmapCache.clear()
    }

    private companion object {
        const val MAX_SLIDE_BITMAP_CACHE_SIZE = 6
    }
}
