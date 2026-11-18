package com.modocs.core.common

import android.graphics.Bitmap
import android.graphics.BitmapFactory

/** A picture's pixel size read from its header, or null if Android cannot decode the format. */
fun bitmapPixelSize(bytes: ByteArray): Pair<Int, Int>? {
    val options = BitmapFactory.Options().apply { inJustDecodeBounds = true }
    BitmapFactory.decodeByteArray(bytes, 0, bytes.size, options)
    return if (options.outWidth > 0 && options.outHeight > 0) options.outWidth to options.outHeight else null
}

/**
 * Decode [bytes] no wider than needed for [targetWidthPx]. A phone photo decoded in full
 * takes ~48 MB of heap; sampling keeps it to what is on screen. Null if Android cannot
 * decode the format (EMF, WMF, TIFF) or memory runs out.
 */
fun decodeSampledBitmap(bytes: ByteArray, targetWidthPx: Int): Bitmap? {
    val (width, _) = bitmapPixelSize(bytes) ?: return null
    var sample = 1
    while (width / (sample * 2) >= targetWidthPx.coerceAtLeast(1)) sample *= 2
    return try {
        BitmapFactory.decodeByteArray(bytes, 0, bytes.size, BitmapFactory.Options().apply { inSampleSize = sample })
    } catch (_: OutOfMemoryError) {
        null
    }
}
