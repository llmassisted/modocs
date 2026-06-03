package com.modocs.core.common

import java.io.ByteArrayOutputStream
import java.io.IOException
import java.io.InputStream
import java.util.zip.ZipInputStream

/**
 * Maximum total uncompressed bytes extracted from a single OOXML container.
 * Guards against decompression ("zip") bombs where a small, highly compressible
 * file expands to gigabytes and exhausts memory.
 */
const val MAX_UNCOMPRESSED_BYTES = 200L * 1024 * 1024 // 200 MB

/** Maximum number of entries, to guard against entry-count explosions. */
private const val MAX_ZIP_ENTRIES = 10_000

/**
 * Read every non-directory entry of a ZIP stream into memory, keyed by entry
 * name, while enforcing total-size and entry-count caps.
 *
 * Entry names are used only as map keys (never as filesystem paths), so this is
 * not vulnerable to path traversal ("zip slip"); the caps below protect against
 * resource-exhaustion DoS from malicious documents.
 *
 * @throws IOException if the uncompressed size or entry count cap is exceeded.
 */
fun readZipEntriesCapped(
    inputStream: InputStream,
    maxTotalBytes: Long = MAX_UNCOMPRESSED_BYTES,
): Map<String, ByteArray> {
    val entries = LinkedHashMap<String, ByteArray>()
    var totalBytes = 0L
    var entryCount = 0

    ZipInputStream(inputStream.buffered()).use { zip ->
        var entry = zip.nextEntry
        val buffer = ByteArray(8192)
        while (entry != null) {
            if (!entry.isDirectory) {
                if (++entryCount > MAX_ZIP_ENTRIES) {
                    throw IOException("Archive has too many entries")
                }
                val baos = ByteArrayOutputStream()
                while (true) {
                    val read = zip.read(buffer)
                    if (read == -1) break
                    totalBytes += read
                    if (totalBytes > maxTotalBytes) {
                        throw IOException("Archive exceeds maximum uncompressed size")
                    }
                    baos.write(buffer, 0, read)
                }
                entries[entry.name] = baos.toByteArray()
            }
            zip.closeEntry()
            entry = zip.nextEntry
        }
    }

    return entries
}
