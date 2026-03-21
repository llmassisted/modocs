package com.modocs.core.ui.components

import androidx.compose.foundation.gestures.awaitEachGesture
import androidx.compose.foundation.gestures.awaitFirstDown
import androidx.compose.foundation.gestures.calculatePan
import androidx.compose.foundation.gestures.calculateZoom
import androidx.compose.foundation.layout.Box
import androidx.compose.foundation.layout.BoxScope
import androidx.compose.runtime.Composable
import androidx.compose.runtime.getValue
import androidx.compose.runtime.mutableFloatStateOf
import androidx.compose.runtime.mutableIntStateOf
import androidx.compose.runtime.remember
import androidx.compose.runtime.setValue
import androidx.compose.ui.Modifier
import androidx.compose.ui.geometry.Offset
import androidx.compose.ui.graphics.graphicsLayer
import androidx.compose.ui.input.nestedscroll.NestedScrollConnection
import androidx.compose.ui.input.nestedscroll.NestedScrollSource
import androidx.compose.ui.input.nestedscroll.nestedScroll
import androidx.compose.ui.input.pointer.pointerInput
import androidx.compose.ui.layout.onSizeChanged

/**
 * A container that supports pinch-to-zoom while preserving inner scroll behavior.
 *
 * Pinch (multi-touch) gestures control zoom and pan. Single-finger gestures pass
 * through to child scrollable content (e.g. LazyColumn). When zoomed in, single-finger
 * scrolls are first consumed for panning within the zoom bounds, with any overflow
 * delegated to the child scroll container.
 *
 * @param modifier Modifier for the outer container.
 * @param enabled Whether zoom gestures are active.
 * @param minScale Minimum zoom scale (default 1f).
 * @param maxScale Maximum zoom scale (default 4f).
 * @param contentModifier A modifier applied to the inner content wrapper that includes
 *   the graphicsLayer transform. Callers can chain additional modifiers onto this.
 * @param content The scrollable content (e.g. LazyColumn).
 */
@Composable
fun ZoomableContainer(
    modifier: Modifier = Modifier,
    enabled: Boolean = true,
    minScale: Float = 1f,
    maxScale: Float = 4f,
    contentModifier: Modifier = Modifier,
    content: @Composable BoxScope.() -> Unit,
) {
    var scale by remember { mutableFloatStateOf(minScale) }
    var offsetX by remember { mutableFloatStateOf(0f) }
    var offsetY by remember { mutableFloatStateOf(0f) }
    var containerWidth by remember { mutableIntStateOf(0) }
    var containerHeight by remember { mutableIntStateOf(0) }

    fun maxOffsetX(): Float = (containerWidth * (scale - 1f)) / 2f
    fun maxOffsetY(): Float = (containerHeight * (scale - 1f)) / 2f

    // When zoomed, intercept single-finger scroll for panning; overflow goes to child.
    val nestedScrollConnection = remember {
        object : NestedScrollConnection {
            override fun onPreScroll(available: Offset, source: NestedScrollSource): Offset {
                if (scale <= minScale) return Offset.Zero

                val maxX = maxOffsetX()
                val maxY = maxOffsetY()

                // Consume horizontal scroll for panning
                val newOffsetX = (offsetX + available.x).coerceIn(-maxX, maxX)
                val consumedX = newOffsetX - offsetX
                offsetX = newOffsetX

                // Consume vertical scroll for panning within bounds
                val newOffsetY = (offsetY + available.y).coerceIn(-maxY, maxY)
                val consumedY = newOffsetY - offsetY
                offsetY = newOffsetY

                return Offset(consumedX, consumedY)
            }
        }
    }

    Box(
        modifier = modifier
            .onSizeChanged {
                containerWidth = it.width
                containerHeight = it.height
            }
            .then(
                if (enabled) {
                    Modifier
                        .nestedScroll(nestedScrollConnection)
                        .pointerInput(Unit) {
                            awaitEachGesture {
                                awaitFirstDown(requireUnconsumed = false)
                                do {
                                    val event = awaitPointerEvent()

                                    if (event.changes.size >= 2) {
                                        // Multi-touch: handle zoom + pan in both axes
                                        val zoomChange = event.calculateZoom()
                                        val panChange = event.calculatePan()

                                        val newScale =
                                            (scale * zoomChange).coerceIn(minScale, maxScale)
                                        scale = newScale

                                        if (newScale > minScale) {
                                            offsetX += panChange.x
                                            offsetY += panChange.y
                                            offsetX =
                                                offsetX.coerceIn(-maxOffsetX(), maxOffsetX())
                                            offsetY =
                                                offsetY.coerceIn(-maxOffsetY(), maxOffsetY())
                                        } else {
                                            offsetX = 0f
                                            offsetY = 0f
                                        }

                                        event.changes.forEach { it.consume() }
                                    } else if (scale > minScale && event.changes.size == 1) {
                                        // Single-finger when zoomed: handle horizontal pan here
                                        // (LazyColumn only scrolls vertically, so horizontal
                                        // drag never reaches the NestedScrollConnection)
                                        val change = event.changes[0]
                                        val dragX = change.position.x - change.previousPosition.x
                                        if (dragX != 0f) {
                                            offsetX = (offsetX + dragX)
                                                .coerceIn(-maxOffsetX(), maxOffsetX())
                                        }
                                        // Don't consume — vertical component still flows
                                        // to LazyColumn via nestedScroll
                                    }
                                } while (event.changes.any { it.pressed })
                            }
                        }
                } else {
                    Modifier
                },
            ),
    ) {
        Box(
            modifier = contentModifier
                .graphicsLayer {
                    scaleX = scale
                    scaleY = scale
                    translationX = offsetX
                    translationY = offsetY
                },
            content = content,
        )
    }
}
