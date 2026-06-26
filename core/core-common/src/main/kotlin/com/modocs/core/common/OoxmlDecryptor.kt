package com.modocs.core.common

import android.content.Context
import android.net.Uri
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.withContext
import org.apache.poi.poifs.crypt.Decryptor
import org.apache.poi.poifs.crypt.EncryptionInfo
import org.apache.poi.poifs.filesystem.POIFSFileSystem
import java.io.ByteArrayInputStream
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
                val header = ByteArray(8)
                val bytesRead = stream.read(header)
                bytesRead == 8 && header.contentEquals(OLE2_MAGIC)
            } ?: false
        } catch (_: Exception) {
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
            try {
                context.contentResolver.openInputStream(uri)?.use { inputStream ->
                    POIFSFileSystem(inputStream).use { poifs ->
                        val root = poifs.root
                        root.hasEntry("EncryptedPackage") && root.hasEntry("EncryptionInfo")
                    }
                } ?: false
            } catch (_: Exception) {
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
            } catch (e: Exception) {
                DecryptResult.Failed(e.message ?: "Decryption failed")
            }
        }
}
