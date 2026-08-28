package com.modocs.core.common

/**
 * Filename prefix for plaintext copies of decrypted documents written to the
 * cache directory.
 *
 * Shared so the code that creates these files and the startup sweep that
 * deletes them cannot drift apart — a rename on one side alone would silently
 * leave decrypted documents on disk.
 */
const val DECRYPTED_TEMP_PREFIX = "decrypted_"

/**
 * Turn a document-loading failure into one user-facing line.
 *
 * A viewer should say "I couldn't open this" rather than disappear, so callers
 * pair this with a `catch (t: Throwable)` that rethrows
 * [kotlinx.coroutines.CancellationException] first. That catch deliberately
 * includes [Error]: an enormous spreadsheet can exhaust the heap and a
 * pathologically nested document can exhaust the stack, and neither is worth
 * taking the process down for.
 */
fun documentErrorMessage(t: Throwable): String = when (t) {
    is OutOfMemoryError -> "This document is too large to open on this device"
    is StackOverflowError -> "This document is too deeply nested to open"
    else -> t.message?.takeIf { it.isNotBlank() } ?: "Unknown error"
}
