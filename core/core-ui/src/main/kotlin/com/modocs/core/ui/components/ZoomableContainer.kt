package com.modocs.core.ui.components

import androidx.compose.foundation.gestures.awaitEachGesture
import androidx.compose.foundation.gestures.awaitFirstDown
import androidx.compose.foundation.gestures.calculatePan
import androidx.compose.foundation.gestures.calculateZoom
import androidx.compose.foundation.layout.Box
import androidx.compose.foundation.layout.BoxScope
import androidx.compose.ui.draw.clipToBounds
import androidx.compose.runtime.Composable
import androidx.compose.runtime.LaunchedEffect
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
 * Operates in two modes controlled by [applyTransform]:
 *
 * - **true (default):** Zoom is applied via graphicsLayer (visual-only scaling).
 *   Best for bitmap/image content where the coordinate system doesn't matter.
 *   Panning is handled via translation offsets.
 *
 * - **false:** No graphicsLayer transform is applied. The current zoom scale is
 *   passed to [content] so callers can apply it to layout dimensions (font sizes,
 *   padding, etc.) for true layout-level zoom. This keeps the input coordinate
 *   system aligned with the visual rendering, which is required for interactive
 *   elements like BasicTextField. Panning is not needed since the child
 *   scroll container handles scrolling at the zoomed dimensions naturally.
 *
 * @param modifier Modifier for the outer container.
 * @param enabled Whether zoom gestures are active.
 * @param scrollLocked When true, all scroll and pan gestures are frozen.
 * @param applyTransform When true, zoom is applied via graphicsLayer. When false,
 *   the scale is passed to content for layout-level zoom.
 * @param minScale Minimum zoom scale (default 1f).
 * @param maxScale Maximum zoom scale (default 4f).
 * @param contentModifier A modifier applied to the inner content wrapper.
 * @param content The scrollable content. Receives the current zoom scale.
 */
@Composable
fun ZoomableContainer(
    modifier: Modifier = Modifier,
    enabled: Boolean = true,
    scrollLocked: Boolean = false,
    applyTransform: Boolean = true,
    resetKey: Any? = null,
    minScale: Float = 1f,
    maxScale: Float = 4f,
    contentModifier: Modifier = Modifier,
    content: @Composable BoxScope.(zoomScale: Float) -> Unit,
) {
    var scale by remember { mutableFloatStateOf(minScale) }
    var offsetX by remember { mutableFloatStateOf(0f) }
    var offsetY by remember { mutableFloatStateOf(0f) }

    // Reset zoom when resetKey changes (e.g. entering/exiting edit mode)
    LaunchedEffect(resetKey) {
        scale = minScale
        offsetX = 0f
        offsetY = 0f
    }
    var containerWidth by remember { mutableIntStateOf(0) }
    var containerHeight by remember { mutableIntStateOf(0) }

    fun maxOffsetX(): Float = (containerWidth * (scale - 1f)) / 2f
    fun maxOffsetY(): Float = (containerHeight * (scale - 1f)) / 2f

    val nestedScrollConnection = remember {
        object : NestedScrollConnection {
            override fun onPreScroll(available: Offset, source: NestedScrollSource): Offset {
                if (scrollLocked) return available

                // In layout-zoom mode, let the child scroll container handle everything
                if (!applyTransform) return Offset.Zero

                if (scale <= minScale) return Offset.Zero

                val maxX = maxOffsetX()
                val maxY = maxOffsetY()

                val newOffsetX = (offsetX + available.x).coerceIn(-maxX, maxX)
                val consumedX = newOffsetX - offsetX
                offsetX = newOffsetX

                val newOffsetY = (offsetY + available.y).coerceIn(-maxY, maxY)
                val consumedY = newOffsetY - offsetY
                offsetY = newOffsetY

                return Offset(consumedX, consumedY)
            }
        }
    }

    Box(
        modifier = modifier
            .clipToBounds()
            .onSizeChanged {
                containerWidth = it.width
                containerHeight = it.height
            }
            .then(
                if (enabled) {
                    Modifier
                        .nestedScroll(nestedScrollConnection)
                        .pointerInput(scrollLocked, applyTransform) {
                            awaitEachGesture {
                                awaitFirstDown(requireUnconsumed = false)
                                do {
                                    val event = awaitPointerEvent()

                                    if (scrollLocked) {
                                        // Don't handle any zoom/pan while dragging annotations
                                    } else if (event.changes.size >= 2) {
                                        val zoomChange = event.calculateZoom()
                                        val panChange = event.calculatePan()

                                        val newScale =
                                            (scale * zoomChange).coerceIn(minScale, maxScale)
                                        scale = newScale

                                        if (applyTransform) {
                                            // graphicsLayer mode: manage offsets
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
                                        }

                                        event.changes.forEach { it.consume() }
                                    } else if (applyTransform && scale > minScale &&
                                        event.changes.size == 1
                                    ) {
                                        // Single-finger horizontal pan (graphicsLayer mode only)
                                        val change = event.changes[0]
                                        val dragX = change.position.x - change.previousPosition.x
                                        if (dragX != 0f) {
                                            offsetX = (offsetX + dragX)
                                                .coerceIn(-maxOffsetX(), maxOffsetX())
                                        }
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
                .then(
                    if (applyTransform) {
                        Modifier.graphicsLayer {
                            scaleX = scale
                            scaleY = scale
                            translationX = offsetX
                            translationY = offsetY
                        }
                    } else {
                        Modifier
                    },
                ),
        ) {
            content(scale)
        }
    }
}
