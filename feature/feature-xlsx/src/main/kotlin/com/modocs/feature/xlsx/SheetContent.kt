package com.modocs.feature.xlsx

import android.graphics.Paint
import androidx.compose.foundation.Canvas
import androidx.compose.foundation.gestures.awaitEachGesture
import androidx.compose.foundation.gestures.awaitFirstDown
import androidx.compose.foundation.gestures.calculateZoom
import androidx.compose.foundation.text.selection.SelectionContainer
import androidx.compose.ui.semantics.customActions
import androidx.compose.ui.semantics.CustomAccessibilityAction
import androidx.compose.foundation.gestures.Orientation
import androidx.compose.foundation.gestures.detectTapGestures
import androidx.compose.foundation.gestures.rememberScrollableState
import androidx.compose.foundation.gestures.scrollable
import androidx.compose.foundation.layout.*
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.lazy.rememberLazyListState
import androidx.compose.material3.*
import androidx.compose.runtime.*
import androidx.compose.runtime.saveable.rememberSaveable
import androidx.compose.ui.Modifier
import androidx.compose.ui.draw.clipToBounds
import androidx.compose.ui.geometry.Offset
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.graphics.nativeCanvas
import androidx.compose.ui.input.pointer.pointerInput
import androidx.compose.ui.platform.LocalDensity
import androidx.compose.ui.semantics.contentDescription
import androidx.compose.ui.semantics.semantics
import androidx.compose.ui.unit.dp
import kotlinx.coroutines.launch
import com.modocs.core.ui.components.RememberReadingPosition

/** Only visible columns are painted; rows remain lazy even for sparse million-row sheets. */
@Composable
internal fun SheetContent(
    sheet: XlsxSheet, documentKey: String, styles: List<XlsxCellStyle>, searchState: XlsxSearchState,
    activeSheetIndex: Int, zoomScale: Float, onZoomChange: (Float) -> Unit,
    isEditing: Boolean, editingCell: Pair<Int, Int>?, onCellClick: (Int, Int) -> Unit,
) {
    key(documentKey, activeSheetIndex) {
        val list = rememberLazyListState()
        val currentZoom by rememberUpdatedState(zoomScale)
        var inspected by remember { mutableStateOf<Pair<Int, Int>?>(null) }
        inspected?.let { (row, col) ->
            val cell = sheet.cellAt(row, col)
            AlertDialog(onDismissRequest = { inspected = null }, title = { Text(cellReference(row, col)) },
                text = { SelectionContainer { Text(cell?.formula?.let { "=$it\nCached result: ${cell.value}" } ?: cell?.value.orEmpty()) } },
                confirmButton = { TextButton(onClick = { inspected = null }) { Text("Done") } })
        }
        val scope = rememberCoroutineScope()
        val density = LocalDensity.current
        var horizontal by rememberSaveable { mutableFloatStateOf(0f) }
        var address by rememberSaveable { mutableStateOf("") }
        var invalidAddress by remember { mutableStateOf(false) }
        var jumpTarget by remember { mutableStateOf<Pair<Int, Int>?>(null) }
        val matchedCells = remember(searchState.matches, activeSheetIndex) { searchState.matches.filter { it.sheetIndex == activeSheetIndex }.map { it.rowIndex to it.colIndex }.toHashSet() }
        val rowMap = remember(sheet) { sheet.rows.associateBy { it.rowIndex } }
        val cells = remember(sheet) { sheet.rows.associate { it.rowIndex to it.cells.associateBy { c -> c.columnIndex } } }
        val colCount = maxOf(sheet.columnCount, sheet.mergedCells.maxOfOrNull { it.endCol + 1 } ?: 0, 1).coerceAtMost(16384)
        val rowCount = maxOf(sheet.rows.maxOfOrNull { it.rowIndex + 1 } ?: 1, sheet.mergedCells.maxOfOrNull { it.endRow + 1 } ?: 1).coerceAtMost(1048576)
        val unit = density.density * zoomScale
        val widths = remember(sheet, unit) { FloatArray(colCount) { ((sheet.columnWidths[it] ?: 10f) * 8f * unit).coerceIn(24f * unit, 800f * unit) } }
        val x = remember(widths) { FloatArray(colCount + 1).also { out -> for (i in widths.indices) out[i + 1] = out[i] + widths[i] } }
        fun rowHeight(row: Int): Float = ((rowMap[row]?.heightPt ?: sheet.defaultRowHeight) * 1.4f).coerceAtLeast(28f) * unit
        fun rowY(row: Int): Float {
            val default = (sheet.defaultRowHeight * 1.4f).coerceAtLeast(28f) * unit
            return row * default + sheet.rows.filter { it.rowIndex < row && it.heightPt != null }.sumOf { (rowHeight(it.rowIndex) - default).toDouble() }.toFloat()
        }
        val frozenRows = sheet.frozenRows.coerceIn(0, minOf(rowCount, 5))
        RememberReadingPosition("$documentKey/sheet/$activeSheetIndex", list, true)
        Column(Modifier.fillMaxSize()) {
            Row(Modifier.fillMaxWidth().padding(horizontal = 8.dp), verticalAlignment = androidx.compose.ui.Alignment.CenterVertically) {
                OutlinedTextField(address, { address = it; invalidAddress = false }, label = { Text("Go to cell") },
                    singleLine = true, isError = invalidAddress, modifier = Modifier.weight(1f))
                TextButton(onClick = {
                    val target = parseCellReference(address)
                    if (target == null || target.first >= rowCount || target.second >= colCount) invalidAddress = true
                    else scope.launch {
                        list.scrollToItem((target.first - frozenRows).coerceAtLeast(0))
                        jumpTarget = target
                        onCellClick(target.first, target.second)
                    }
                }) { Text("Go") }
                TextButton(onClick = { onZoomChange((zoomScale - .25f).coerceAtLeast(.5f)) }) { Text("−") }
                Text("${(zoomScale * 100).toInt()}%", style = MaterialTheme.typography.labelSmall)
                TextButton(onClick = { onZoomChange((zoomScale + .25f).coerceAtMost(3f)) }) { Text("+") }
            }
            if (sheet.rows.any { row -> row.cells.any { it.formula != null } }) Text(
                "Formula results are cached; input edits do not update totals here.",
                style = MaterialTheme.typography.bodySmall, modifier = Modifier.padding(horizontal = 12.dp, vertical = 4.dp))
            if (sheet.frozenRows > 5) Text("Showing the first 5 frozen rows to leave room to read.", style = MaterialTheme.typography.bodySmall)
            BoxWithConstraints(Modifier.weight(1f).fillMaxWidth().pointerInput(Unit) {
                awaitEachGesture {
                    awaitFirstDown(requireUnconsumed = false)
                    do {
                        val event = awaitPointerEvent()
                        if (event.changes.count { it.pressed } >= 2) {
                            val zoom = event.calculateZoom()
                            if (zoom != 1f) {
                                onZoomChange((currentZoom * zoom).coerceIn(.5f, 3f))
                                event.changes.forEach { it.consume() }
                            }
                        }
                    } while (event.changes.any { it.pressed })
                }
            }) {
                val viewport = with(density) { maxWidth.toPx() }
                val headerWidth = 42f * unit
                val frozenCols = (0..sheet.frozenColumns.coerceIn(0, colCount)).lastOrNull { x[it] <= viewport * .4f } ?: 0
                val pinnedWidth = headerWidth + x[frozenCols]
                val maxScroll = (x.last() + headerWidth - viewport).coerceAtLeast(0f)
                val scroll = rememberScrollableState { delta ->
                    val old = horizontal.coerceIn(0f, maxScroll)
                    horizontal = (old - delta).coerceIn(0f, maxScroll)
                    old - horizontal
                }
                LaunchedEffect(jumpTarget) {
                    jumpTarget?.let { horizontal = (x[it.second] - x[frozenCols]).coerceIn(0f, maxScroll) }
                }
                val match = searchState.currentMatch
                LaunchedEffect(match) {
                    if (match != null && match.sheetIndex == activeSheetIndex) {
                        list.animateScrollToItem((match.rowIndex - frozenRows).coerceIn(0, (rowCount - frozenRows - 1).coerceAtLeast(0)))
                        horizontal = (x[match.colIndex.coerceIn(0, colCount - 1)] - x[frozenCols]).coerceIn(0f, maxScroll)
                    }
                }
                @Composable fun GridRow(row: Int) {
                    val height = if (row < 0) 30f * unit else rowHeight(row)
                    Canvas(Modifier.fillMaxWidth().height(with(density) { height.toDp() }).clipToBounds()
                        .semantics { contentDescription = if (row < 0) "Column headings" else "Row ${row + 1}: " + rowMap[row]?.cells.orEmpty().take(30).joinToString { it.value }
                            customActions = rowMap[row]?.cells.orEmpty().take(30).map { cell ->
                                CustomAccessibilityAction("Read ${cellReference(row, cell.columnIndex)}") { inspected = row to cell.columnIndex; true }
                            } }
                        .pointerInput(row, horizontal, sheet, isEditing) {
                            detectTapGestures(onLongPress = { point ->
                                if (row >= 0 && point.x >= headerWidth) {
                                    val logical = point.x - headerWidth + if (point.x >= pinnedWidth) horizontal.coerceIn(0f, maxScroll) else 0f
                                    if (logical < x.last()) {
                                        val col = x.binarySearch(logical).let { if (it >= 0) it else -it - 2 }.coerceIn(0, colCount - 1)
                                        val merge = sheet.mergedCells.firstOrNull { row in it.startRow..it.endRow && col in it.startCol..it.endCol }
                                        inspected = (merge?.startRow ?: row) to (merge?.startCol ?: col)
                                    }
                                }
                            }, onTap = { point ->
                                if (row < 0 || point.x < headerWidth) return@detectTapGestures
                                val logicalX = point.x - headerWidth + if (point.x >= pinnedWidth) horizontal.coerceIn(0f, maxScroll) else 0f
                                if (logicalX >= x.last()) return@detectTapGestures
                                val col = x.binarySearch(logicalX).let { if (it >= 0) it else -it - 2 }.coerceIn(0, colCount - 1)
                                val merge = sheet.mergedCells.firstOrNull { row in it.startRow..it.endRow && col in it.startCol..it.endCol }
                                onCellClick(merge?.startRow ?: row, merge?.startCol ?: col)
                            })
                        }) {
                        val canvas = drawContext.canvas.nativeCanvas
                        canvas.drawColor(android.graphics.Color.WHITE)
                        val paint = Paint(Paint.ANTI_ALIAS_FLAG)
                        fun drawPartition(startCol: Int, endCol: Int, shift: Float, leftClip: Float, rightClip: Float) {
                            canvas.save(); canvas.clipRect(leftClip, 0f, rightClip, height)
                            var col = startCol
                            while (col < endCol) {
                                val merge = if (row < 0) null else sheet.mergedCells.firstOrNull { row in it.startRow..it.endRow && col in it.startCol..it.endCol }
                                val from = merge?.startCol ?: col
                                val to = (merge?.endCol?.plus(1) ?: (col + 1)).coerceAtMost(colCount)
                                val left = headerWidth + x[from] - shift
                                val right = headerWidth + x[to] - shift
                                if (right >= leftClip && left <= rightClip) {
                                    val originRow = merge?.startRow ?: row
                                    val top = if (merge == null) 0f else rowY(merge.startRow) - rowY(row)
                                    val bottom = if (merge == null) height else rowY(merge.endRow + 1) - rowY(row)
                                    val cell = cells[originRow]?.get(from)
                                    val style = cell?.let { styles.getOrNull(it.styleIndex) }
                                    val selected = editingCell == (originRow to from)
                                    val matched = (originRow to from) in matchedCells
                                    paint.style = Paint.Style.FILL
                                    paint.color = when { row < 0 -> 0xFFF0F3F8.toInt(); selected -> 0xFFD6E3FF.toInt(); matched -> 0xFFFFE082.toInt(); else -> style?.fillColor ?: android.graphics.Color.WHITE }
                                    canvas.drawRect(left, 0f, right, height, paint)
                                    paint.color = 0xFFCDD3DC.toInt(); paint.style = Paint.Style.STROKE; paint.strokeWidth = 1f
                                    canvas.drawRect(left, top, right, bottom, paint)
                                    paint.style = Paint.Style.FILL; paint.color = style?.fontColor ?: android.graphics.Color.BLACK
                                    paint.textSize = (style?.fontSize ?: 11f) * unit * density.fontScale
                                    paint.isFakeBoldText = row < 0 || style?.fontBold == true
                                    paint.isUnderlineText = style?.fontUnderline == true
                                    paint.textSkewX = if (style?.fontItalic == true) -.2f else 0f
                                    val text = if (row < 0) columnLetter(col) else cell?.value.orEmpty()
                                    canvas.save(); canvas.clipRect(left + 3f, top, right - 3f, bottom)
                                    val textPaint = android.text.TextPaint(paint)
                                    val align = when (style?.horizontalAlignment) {
                                        CellAlignment.CENTER -> android.text.Layout.Alignment.ALIGN_CENTER
                                        CellAlignment.RIGHT -> android.text.Layout.Alignment.ALIGN_OPPOSITE
                                        CellAlignment.GENERAL -> if (cell?.type == CellType.NUMBER || cell?.type == CellType.DATE) android.text.Layout.Alignment.ALIGN_OPPOSITE else android.text.Layout.Alignment.ALIGN_NORMAL
                                        else -> android.text.Layout.Alignment.ALIGN_NORMAL
                                    }
                                    val layout = android.text.StaticLayout.Builder.obtain(text, 0, text.length, textPaint,
                                        (right - left - 10f * unit).toInt().coerceAtLeast(1))
                                        .setAlignment(align).setIncludePad(false).setMaxLines(if (style?.wrapText == true) 100 else 1).build()
                                    val textY = when (style?.verticalAlignment) {
                                        CellVerticalAlignment.BOTTOM -> (bottom - layout.height - 4f * unit).coerceAtLeast(top + 4f * unit)
                                        CellVerticalAlignment.CENTER -> top + ((bottom - top - layout.height) / 2f).coerceAtLeast(0f)
                                        else -> top + 4f * unit
                                    }
                                    canvas.translate(left + 5f * unit, textY)
                                    layout.draw(canvas)
                                    canvas.restore()
                                }
                                col = maxOf(col + 1, to)
                            }
                            canvas.restore()
                        }
                        val shift = horizontal.coerceIn(0f, maxScroll)
                        val first = maxOf(frozenCols, x.binarySearch(shift + x[frozenCols]).let { if (it >= 0) it else -it - 2 })
                        val last = (x.binarySearch(shift + viewport).let { if (it >= 0) it + 1 else -it }).coerceIn(first, colCount)
                        drawPartition(first, last, shift, pinnedWidth, viewport)
                        drawPartition(0, frozenCols, 0f, headerWidth, pinnedWidth)
                        paint.style = Paint.Style.FILL; paint.color = 0xFFF0F3F8.toInt()
                        canvas.drawRect(0f, 0f, headerWidth, height, paint)
                        paint.color = android.graphics.Color.DKGRAY; paint.textSize = 10f * unit; paint.isFakeBoldText = true
                        if (row >= 0) canvas.drawText("${row + 1}", 4f * unit, 19f * unit, paint)
                    }
                }
                Column(Modifier.fillMaxSize().scrollable(scroll, Orientation.Horizontal)) {
                    GridRow(-1)
                    for (row in 0 until frozenRows) GridRow(row)
                    LazyColumn(state = list, modifier = Modifier.weight(1f).fillMaxWidth()) {
                        items((rowCount - frozenRows).coerceAtLeast(0)) { GridRow(it + frozenRows) }
                    }
                }
            }
        }
    }
}
