package com.modocs.feature.docx

import android.graphics.BitmapFactory
import androidx.compose.foundation.Image
import androidx.compose.foundation.background
import androidx.compose.foundation.border
import androidx.compose.foundation.clickable
import androidx.compose.foundation.horizontalScroll
import androidx.compose.foundation.layout.Arrangement
import androidx.compose.foundation.layout.Box
import androidx.compose.foundation.layout.Column
import androidx.compose.foundation.layout.IntrinsicSize
import androidx.compose.foundation.layout.Row
import androidx.compose.foundation.layout.Spacer
import androidx.compose.foundation.layout.fillMaxHeight
import androidx.compose.foundation.layout.fillMaxWidth
import androidx.compose.foundation.layout.height
import androidx.compose.foundation.layout.padding
import androidx.compose.foundation.layout.width
import androidx.compose.foundation.layout.widthIn
import androidx.compose.foundation.rememberScrollState
import androidx.compose.foundation.text.BasicTextField
import androidx.compose.material3.HorizontalDivider
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.Text
import androidx.compose.runtime.Composable
import androidx.compose.runtime.LaunchedEffect
import androidx.compose.runtime.getValue
import androidx.compose.runtime.mutableStateOf
import androidx.compose.runtime.remember
import androidx.compose.runtime.setValue
import androidx.compose.ui.Modifier
import androidx.compose.ui.focus.FocusRequester
import androidx.compose.ui.focus.focusRequester
import androidx.compose.ui.draw.drawBehind
import androidx.compose.ui.geometry.Offset
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.graphics.SolidColor
import androidx.compose.ui.graphics.asImageBitmap
import androidx.compose.ui.layout.ContentScale
import androidx.compose.ui.text.AnnotatedString
import androidx.compose.ui.text.SpanStyle
import androidx.compose.ui.text.TextStyle
import androidx.compose.ui.text.buildAnnotatedString
import androidx.compose.ui.text.font.FontStyle
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.text.style.BaselineShift
import androidx.compose.ui.text.style.TextAlign
import androidx.compose.ui.text.style.TextDecoration
import androidx.compose.ui.text.withStyle
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp
import com.modocs.fonts.FontResolver

/**
 * Renders a single [DocxElement] as a Compose composable.
 *
 * @param pageScale Scale factor to match CANVAS view proportions.
 *                  Computed as `screenWidthDp / pageWidthPt` so that margins,
 *                  font sizes, and spacing are proportional to the bitmap view.
 */
@Composable
fun DocxElementRenderer(
    element: DocxElement,
    document: DocxDocument,
    fontResolver: FontResolver,
    pageScale: Float = 1f,
    searchHighlights: List<IntRange> = emptyList(),
    currentHighlightIndex: Int = -1,
    textOffset: Int = 0,
    isEditing: Boolean = false,
    isActivelyEditing: Boolean = false,
    onTapToEdit: () -> Unit = {},
    onTextChanged: (String) -> Unit = {},
    modifier: Modifier = Modifier,
) {
    when (element) {
        is DocxParagraph -> ParagraphRenderer(
            paragraph = element,
            document = document,
            fontResolver = fontResolver,
            pageScale = pageScale,
            searchHighlights = searchHighlights,
            currentHighlightIndex = currentHighlightIndex,
            textOffset = textOffset,
            isEditing = isEditing,
            isActivelyEditing = isActivelyEditing,
            onTapToEdit = onTapToEdit,
            onTextChanged = onTextChanged,
            modifier = modifier,
        )
        is DocxTable -> TableRenderer(
            table = element,
            document = document,
            fontResolver = fontResolver,
            pageScale = pageScale,
            modifier = modifier,
        )
        is DocxImage -> ImageRenderer(
            image = element,
            document = document,
            modifier = modifier,
        )
    }
}

// --- Paragraph ---

@Composable
private fun ParagraphRenderer(
    paragraph: DocxParagraph,
    document: DocxDocument,
    fontResolver: FontResolver,
    pageScale: Float,
    searchHighlights: List<IntRange>,
    currentHighlightIndex: Int,
    textOffset: Int,
    isEditing: Boolean,
    isActivelyEditing: Boolean,
    onTapToEdit: () -> Unit,
    onTextChanged: (String) -> Unit,
    modifier: Modifier = Modifier,
) {
    val props = paragraph.properties

    // Handle page break before
    if (props.pageBreakBefore) {
        PageBreakIndicator()
    }

    // Handle list bullet/number prefix
    val listPrefix = remember(paragraph.listInfo, document.numbering) {
        buildListPrefix(paragraph.listInfo, document.numbering)
    }

    val textAlign = when (props.alignment) {
        ParagraphAlignment.CENTER -> TextAlign.Center
        ParagraphAlignment.RIGHT -> TextAlign.End
        ParagraphAlignment.JUSTIFY -> TextAlign.Justify
        ParagraphAlignment.LEFT -> TextAlign.Start
    }

    // Heading scale factors matching DocxPageRenderer/DocxToPdfConverter
    val headingScale = when (props.headingLevel) {
        1 -> 1.8f; 2 -> 1.5f; 3 -> 1.3f; 4 -> 1.15f; 5 -> 1.1f; 6 -> 1.05f
        else -> 1f
    }
    val isHeading = props.headingLevel in 1..6

    // Use document font sizes scaled proportionally to match CANVAS view
    val baseFontSizePt = paragraph.runs.firstOrNull()?.properties?.fontSizeSp ?: 11f
    val scaledFontSize = (baseFontSizePt * headingScale * pageScale).sp

    // Scale indents proportionally (twips → points → scaled dp)
    val leftIndentDp = (props.indentLeftTwips / 20f * pageScale).dp
    val rightIndentDp = (props.indentRightTwips / 20f * pageScale).dp

    val hasPageBreak = paragraph.runs.any { it.isPageBreak }

    // Scale paragraph spacing proportionally
    val contentModifier = modifier
        .fillMaxWidth()
        .padding(
            start = leftIndentDp,
            end = rightIndentDp,
            top = (props.spacingBeforePt * pageScale).dp,
            bottom = (props.spacingAfterPt * pageScale).dp,
        )

    val baseTextStyle = TextStyle(
        color = Color.Black,
        fontSize = scaledFontSize,
        fontWeight = if (isHeading) FontWeight.Bold else null,
    )

    // Line height: use effectiveLineHeight scaled proportionally
    val lineH = (props.effectiveLineHeight(baseFontSizePt * headingScale) * pageScale).sp

    if (isActivelyEditing) {
        // Editable mode: BasicTextField
        val plainText = paragraph.text
        var editText by remember(plainText) { mutableStateOf(plainText) }
        val focusRequester = remember { FocusRequester() }

        val rp = paragraph.runs.firstOrNull()?.properties ?: RunProperties()
        val runFontSize = (rp.fontSizeSp ?: 11f) * headingScale * pageScale
        val editStyle = baseTextStyle.merge(
            TextStyle(
                fontWeight = if (rp.bold || isHeading) FontWeight.Bold else null,
                fontStyle = if (rp.italic) FontStyle.Italic else null,
                textDecoration = buildTextDecoration(rp),
                fontSize = runFontSize.sp,
                color = rp.color?.let { Color(it) } ?: Color.Black,
                fontFamily = fontResolver.resolve(rp.fontName),
                lineHeight = lineH,
                textAlign = textAlign,
            )
        )

        LaunchedEffect(Unit) {
            focusRequester.requestFocus()
        }

        Box(
            modifier = contentModifier
                .border(1.dp, Color(0xFF1976D2), MaterialTheme.shapes.small)
                .background(
                    Color(0x0A1976D2),
                    MaterialTheme.shapes.small,
                )
                .padding(4.dp),
        ) {
            BasicTextField(
                value = editText,
                onValueChange = { newText ->
                    editText = newText
                    onTextChanged(newText)
                },
                modifier = Modifier
                    .fillMaxWidth()
                    .focusRequester(focusRequester),
                textStyle = editStyle,
                cursorBrush = SolidColor(Color(0xFF1976D2)),
            )
        }
    } else {
        // Read-only mode
        val annotatedText = remember(paragraph.runs, searchHighlights, currentHighlightIndex, pageScale) {
            buildParagraphText(
                runs = paragraph.runs,
                fontResolver = fontResolver,
                styles = document.styles,
                listPrefix = listPrefix,
                searchHighlights = searchHighlights,
                currentHighlightIndex = currentHighlightIndex,
                textOffset = textOffset,
                pageScale = pageScale,
                headingScale = headingScale,
            )
        }

        if (annotatedText.isEmpty() && paragraph.runs.isEmpty()) {
            if (isEditing) {
                // Show clickable empty paragraph placeholder in edit mode
                Text(
                    text = " ",
                    modifier = contentModifier.clickable { onTapToEdit() },
                    style = baseTextStyle,
                )
            } else {
                Spacer(modifier = modifier.height((props.spacingAfterPt.coerceAtLeast(4f) * pageScale).dp))
            }
            return
        }

        Text(
            text = annotatedText,
            modifier = if (isEditing) {
                contentModifier.clickable { onTapToEdit() }
            } else {
                contentModifier
            },
            textAlign = textAlign,
            style = baseTextStyle,
            lineHeight = lineH,
        )
    }

    if (hasPageBreak) {
        PageBreakIndicator()
    }
}

private fun buildParagraphText(
    runs: List<DocxRun>,
    fontResolver: FontResolver,
    styles: Map<String, DocxStyle>,
    listPrefix: String?,
    searchHighlights: List<IntRange>,
    currentHighlightIndex: Int,
    textOffset: Int,
    pageScale: Float = 1f,
    headingScale: Float = 1f,
): AnnotatedString {
    return buildAnnotatedString {
        // Add list prefix
        if (listPrefix != null) {
            append("$listPrefix ")
        }

        for (run in runs) {
            if (run.isPageBreak) continue // rendered separately
            if (run.isBreak) {
                append("\n")
                continue
            }

            val rp = run.properties
            val scaledFontSize = ((rp.fontSizeSp ?: 11f) * headingScale * pageScale).sp
            val spanStyle = SpanStyle(
                fontWeight = if (rp.bold) FontWeight.Bold else null,
                fontStyle = if (rp.italic) FontStyle.Italic else null,
                textDecoration = buildTextDecoration(rp),
                fontSize = scaledFontSize,
                color = rp.color?.let { Color(it) } ?: Color.Black,
                background = rp.highlight?.let { Color(it) } ?: Color.Unspecified,
                fontFamily = fontResolver.resolve(rp.fontName),
                baselineShift = when {
                    rp.superscript -> BaselineShift.Superscript
                    rp.subscript -> BaselineShift.Subscript
                    else -> null
                },
            )

            withStyle(spanStyle) {
                append(run.text)
            }
        }

        // Apply search highlights on top
        val fullText = toString()
        for ((index, range) in searchHighlights.withIndex()) {
            val adjustedStart = range.first - textOffset
            val adjustedEnd = range.last + 1 - textOffset
            if (adjustedStart < 0 || adjustedEnd > fullText.length) continue

            val isCurrent = index == currentHighlightIndex
            addStyle(
                SpanStyle(
                    background = if (isCurrent) Color(0x99FF9800) else Color(0x66FFEB3B),
                ),
                start = adjustedStart.coerceAtLeast(0),
                end = adjustedEnd.coerceAtMost(fullText.length),
            )
        }
    }
}

private fun buildTextDecoration(rp: RunProperties): TextDecoration? {
    val decorations = mutableListOf<TextDecoration>()
    if (rp.underline) decorations.add(TextDecoration.Underline)
    if (rp.strikethrough) decorations.add(TextDecoration.LineThrough)
    return if (decorations.isEmpty()) null else TextDecoration.combine(decorations)
}

private fun buildListPrefix(
    listInfo: ListInfo?,
    numbering: Map<String, NumberingDefinition>,
): String? {
    if (listInfo == null) return null
    val numDef = numbering[listInfo.numId]
    val level = numDef?.levels?.get(listInfo.level)
    return if (level != null) {
        if (level.format == NumberFormat.BULLET) "\u2022" else level.format.formatNumber(1)
    } else {
        "\u2022"
    }
}

@Composable
private fun PageBreakIndicator() {
    Row(
        modifier = Modifier
            .fillMaxWidth()
            .padding(vertical = 20.dp),
        verticalAlignment = androidx.compose.ui.Alignment.CenterVertically,
        horizontalArrangement = Arrangement.Center,
    ) {
        androidx.compose.foundation.Canvas(
            modifier = Modifier.weight(1f).height(2.dp),
        ) {
            drawLine(
                color = Color.Black,
                start = androidx.compose.ui.geometry.Offset(0f, size.height / 2),
                end = androidx.compose.ui.geometry.Offset(size.width, size.height / 2),
                strokeWidth = 2f,
                pathEffect = androidx.compose.ui.graphics.PathEffect.dashPathEffect(
                    floatArrayOf(10f, 6f),
                    0f,
                ),
            )
        }
        Text(
            text = "  PAGE BREAK  ",
            style = MaterialTheme.typography.labelSmall,
            color = Color.Black,
            fontWeight = FontWeight.Bold,
        )
        androidx.compose.foundation.Canvas(
            modifier = Modifier.weight(1f).height(2.dp),
        ) {
            drawLine(
                color = Color.Black,
                start = androidx.compose.ui.geometry.Offset(0f, size.height / 2),
                end = androidx.compose.ui.geometry.Offset(size.width, size.height / 2),
                strokeWidth = 2f,
                pathEffect = androidx.compose.ui.graphics.PathEffect.dashPathEffect(
                    floatArrayOf(10f, 6f),
                    0f,
                ),
            )
        }
    }
}

// --- Table ---

@Composable
private fun TableRenderer(
    table: DocxTable,
    document: DocxDocument,
    fontResolver: FontResolver,
    pageScale: Float,
    modifier: Modifier = Modifier,
) {
    val gridColWidths = table.gridColWidths
    val totalGridTwips = gridColWidths.sum().takeIf { it > 0 }
    val tableBorders = table.properties.borders

    Column(
        modifier = modifier
            .fillMaxWidth()
            .padding(vertical = (8f * pageScale).dp)
            .horizontalScroll(rememberScrollState()),
    ) {
        for ((rowIdx, row) in table.rows.withIndex()) {
            val isFirstRow = rowIdx == 0
            val isLastRow = rowIdx == table.rows.lastIndex

            Row(
                modifier = Modifier
                    .fillMaxWidth()
                    .height(IntrinsicSize.Min),
            ) {
                var gridIdx = 0
                for ((cellIdx, cell) in row.cells.withIndex()) {
                    val isFirstCol = cellIdx == 0
                    val isLastCol = cellIdx == row.cells.lastIndex

                    if (cell.vMerge == VMergeType.CONTINUE) {
                        gridIdx += cell.gridSpan.coerceAtLeast(1)
                        continue
                    }

                    // Compute cell width from grid columns + gridSpan
                    val span = cell.gridSpan.coerceAtLeast(1)
                    val cellWidth = if (totalGridTwips != null && gridIdx < gridColWidths.size) {
                        var twips = 0
                        for (i in 0 until span) {
                            if (gridIdx + i < gridColWidths.size) twips += gridColWidths[gridIdx + i]
                        }
                        (twips / 20f * pageScale).dp
                    } else {
                        cell.widthTwips?.let { (it / 20f * pageScale).dp }
                    }
                    gridIdx += span

                    // Resolve borders: cell overrides > table defaults
                    val cellBorders = cell.borders
                    val topBorder = cellBorders?.top ?: if (isFirstRow) tableBorders.top else tableBorders.insideH
                    val bottomBorder = cellBorders?.bottom ?: if (isLastRow) tableBorders.bottom else tableBorders.insideH
                    val leftBorder = cellBorders?.left ?: if (isFirstCol) tableBorders.left else tableBorders.insideV
                    val rightBorder = cellBorders?.right ?: if (isLastCol) tableBorders.right else tableBorders.insideV

                    Box(
                        modifier = Modifier
                            .then(
                                if (cellWidth != null) Modifier.width(cellWidth)
                                else Modifier.weight(1f)
                            )
                            .fillMaxHeight()
                            .drawCellBorders(topBorder, bottomBorder, leftBorder, rightBorder, pageScale)
                            .then(
                                if (cell.shading != null) Modifier.background(Color(cell.shading!!))
                                else Modifier
                            )
                            .padding((2f * pageScale).dp),
                    ) {
                        Column {
                            for (para in cell.paragraphs) {
                                ParagraphRenderer(
                                    paragraph = para,
                                    document = document,
                                    fontResolver = fontResolver,
                                    pageScale = pageScale,
                                    searchHighlights = emptyList(),
                                    currentHighlightIndex = -1,
                                    textOffset = 0,
                                    isEditing = false,
                                    isActivelyEditing = false,
                                    onTapToEdit = {},
                                    onTextChanged = {},
                                )
                            }
                        }
                    }
                }
            }
        }
    }
}

/** Draw individual cell borders based on resolved border settings. */
private fun Modifier.drawCellBorders(
    top: CellBorder, bottom: CellBorder, left: CellBorder, right: CellBorder,
    pageScale: Float,
): Modifier = this.drawBehind {
    fun borderWidth(border: CellBorder): Float =
        (border.widthEighthPt / 8f * pageScale).dp.toPx().coerceAtLeast(0.5f)
    fun borderColor(border: CellBorder): Color =
        if (border.color != null) Color(border.color) else Color.Black

    if (top.style != BorderStyle.NONE) {
        val w = borderWidth(top)
        drawLine(borderColor(top), Offset(0f, w / 2), Offset(size.width, w / 2), strokeWidth = w)
    }
    if (bottom.style != BorderStyle.NONE) {
        val w = borderWidth(bottom)
        drawLine(borderColor(bottom), Offset(0f, size.height - w / 2), Offset(size.width, size.height - w / 2), strokeWidth = w)
    }
    if (left.style != BorderStyle.NONE) {
        val w = borderWidth(left)
        drawLine(borderColor(left), Offset(w / 2, 0f), Offset(w / 2, size.height), strokeWidth = w)
    }
    if (right.style != BorderStyle.NONE) {
        val w = borderWidth(right)
        drawLine(borderColor(right), Offset(size.width - w / 2, 0f), Offset(size.width - w / 2, size.height), strokeWidth = w)
    }
}

// --- Image ---

@Composable
private fun ImageRenderer(
    image: DocxImage,
    document: DocxDocument,
    modifier: Modifier = Modifier,
) {
    val imageBytes = document.images[image.relationId]
    if (imageBytes == null) {
        Box(
            modifier = modifier
                .fillMaxWidth()
                .height(100.dp)
                .background(Color(0xFFE0E0E0))
                .padding(8.dp),
        ) {
            Text(
                text = image.altText.ifEmpty { "[Image]" },
                style = MaterialTheme.typography.bodySmall,
                color = Color.DarkGray,
            )
        }
        return
    }

    val bitmap = remember(image.relationId) {
        try {
            BitmapFactory.decodeByteArray(imageBytes, 0, imageBytes.size)
        } catch (_: Exception) {
            null
        }
    }

    if (bitmap != null) {
        val widthDp = if (image.widthEmu > 0) DocxUnits.emuToDp(image.widthEmu).dp else 400.dp

        Image(
            bitmap = bitmap.asImageBitmap(),
            contentDescription = image.altText.ifEmpty { "Document image" },
            modifier = modifier
                .widthIn(max = widthDp)
                .padding(vertical = 4.dp),
            contentScale = ContentScale.FillWidth,
        )
    }
}
