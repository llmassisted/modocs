package com.modocs.core.common

import android.content.Context
import android.net.Uri
import android.util.Log
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.withContext
import org.apache.poi.poifs.crypt.Decryptor
import org.apache.poi.poifs.crypt.EncryptionInfo
import org.apache.poi.poifs.filesystem.POIFSFileSystem
import java.io.ByteArrayInputStream
import java.io.DataInputStream
import java.io.InputStream
import java.security.GeneralSecurityException

/**
 * Detects and decrypts password-protected OOXML files (DOCX, XLSX, PPTX).
 *
 * Password-protected OOXML files are stored as OLE2 compound documents
 * containing an encrypted ZIP package. This utility uses Apache POI to
 * decrypt them back into a regular ZIP stream for the existing parsers.
 */
object OoxmlDecryptor {

    private const val TAG = "OoxmlDecryptor"

    /** OLE2 compound document magic bytes: D0 CF 11 E0 A1 B2 1C E1 */
    private val OLE2_MAGIC = byteArrayOf(
        0xD0.toByte(), 0xCF.toByte(), 0x11.toByte(), 0xE0.toByte(),
        0xA1.toByte(), 0xB2.toByte(), 0x1C.toByte(), 0xE1.toByte(),
    )

    sealed class DecryptResult {
        data class Success(val inputStream: InputStream) : DecryptResult()
        data object WrongPassword : DecryptResult()
        data class Failed(val message: String) : DecryptResult()
    }

    /**
     * Check if a file is an OLE2 compound document (i.e. password-protected OOXML)
     * by reading the first 8 magic bytes.
     */
    suspend fun isOle2File(context: Context, uri: Uri): Boolean = withContext(Dispatchers.IO) {
        try {
            context.contentResolver.openInputStream(uri)?.use { stream ->
                val header = ByteArray(OLE2_MAGIC.size)
                // readFully, not a single read(): read() may return fewer bytes
                // than asked for, and a content:// stream backed by a pipe rather
                // than a file (cloud providers do this) makes that a real
                // possibility. A short read here would mean a password-protected
                // file is treated as unencrypted and never prompts for its
                // password, because this call gates isEncryptedOoxmlFile.
                DataInputStream(stream).readFully(header)
                header.contentEquals(OLE2_MAGIC)
            } ?: false
        } catch (_: Exception) {
            // Includes EOFException: a file shorter than the magic isn't OLE2.
            false
        }
    }

    /**
     * Check whether an OLE2 file has the encrypted OOXML package streams.
     * Legacy binary Office files are also OLE2, but they do not contain these
     * entries and should not be presented as password-protected DOCX/XLSX/PPTX.
     */
    suspend fun isEncryptedOoxmlFile(context: Context, uri: Uri): Boolean =
        withContext(Dispatchers.IO) {
            // Cheap magic-byte gate first. An ordinary OOXML file is a ZIP, not an
            // OLE2 container, so the overwhelmingly common case returns here
            // without POI ever being touched. That matters: this method runs on
            // every document open, so loading a heavyweight library eagerly means
            // any problem inside it blocks opening perfectly normal files.
            if (!isOle2File(context, uri)) return@withContext false

            try {
                context.contentResolver.openInputStream(uri)?.use { inputStream ->
                    POIFSFileSystem(inputStream).use { poifs ->
                        val root = poifs.root
                        root.hasEntry("EncryptedPackage") && root.hasEntry("EncryptionInfo")
                    }
                } ?: false
            } catch (t: Throwable) {
                // Throwable, not Exception: a missing or unloadable POI dependency
                // surfaces as NoClassDefFoundError / ExceptionInInitializerError,
                // which are Errors. Treat a failed probe as "not encrypted" and let
                // the normal parser report something the user can act on.
                Log.w(TAG, "OLE2 encryption probe failed, treating as unencrypted: $t")
                false
            }
        }

    /**
     * Decrypt a password-protected OOXML file.
     * Returns the decrypted ZIP stream on success, which can be passed directly
     * to existing OOXML parsers (DocxParser, XlsxParser, PptxParser).
     */
    suspend fun decrypt(context: Context, uri: Uri, password: String): DecryptResult =
        withContext(Dispatchers.IO) {
            try {
                context.contentResolver.openInputStream(uri).use { inputStream ->
                    if (inputStream == null) {
                        return@withContext DecryptResult.Failed("Cannot open file")
                    }
                    POIFSFileSystem(inputStream).use { poifs ->
                        val info = EncryptionInfo(poifs)
                        val decryptor = Decryptor.getInstance(info)

                        if (!decryptor.verifyPassword(password)) {
                            return@withContext DecryptResult.WrongPassword
                        }

                        // Read the decrypted package fully into memory so the result is
                        // self-contained; the POIFS/source stream are closed on return.
                        val bytes = decryptor.getDataStream(poifs).use { it.readBytes() }
                        DecryptResult.Success(ByteArrayInputStream(bytes))
                    }
                }
            } catch (_: GeneralSecurityException) {
                DecryptResult.WrongPassword
            } catch (t: Throwable) {
                // Throwable for the same reason as the probe above: if POI cannot
                // load, the user gets a message rather than a dead app.
                DecryptResult.Failed(t.message ?: "Decryption failed")
            }
        }
}
