package com.modocs.feature.pdf

import android.content.Context
import android.net.Uri
import android.os.ParcelFileDescriptor
import com.modocs.core.common.DECRYPTED_TEMP_PREFIX
import com.tom_roush.pdfbox.android.PDFBoxResourceLoader
import com.tom_roush.pdfbox.pdmodel.PDDocument
import com.tom_roush.pdfbox.pdmodel.encryption.InvalidPasswordException
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.withContext
import java.io.File
import java.io.IOException

/**
 * Decrypts password-protected PDF files using PDFBox.
 *
 * Loads the encrypted PDF with the user's password, removes all security,
 * and saves a decrypted copy to a temp file. The temp file can then be
 * opened by Android's built-in PdfRenderer.
 */
object PdfDecryptor {

    sealed class DecryptResult {
        data class Success(val tempFile: File) : DecryptResult()
        data object WrongPassword : DecryptResult()
        data class Failed(val message: String) : DecryptResult()
    }

    suspend fun decrypt(context: Context, uri: Uri, password: String): DecryptResult =
        withContext(Dispatchers.IO) {
            PDFBoxResourceLoader.init(context)

            val inputStream = context.contentResolver.openInputStream(uri)
                ?: return@withContext DecryptResult.Failed("Cannot open file")

            try {
                val document = PDDocument.load(inputStream, password)

                val tempFile = File(context.cacheDir, "$DECRYPTED_TEMP_PREFIX${System.nanoTime()}.pdf")
                document.isAllSecurityToBeRemoved = true
                document.save(tempFile)
                document.close()

                DecryptResult.Success(tempFile)
            } catch (_: InvalidPasswordException) {
                DecryptResult.WrongPassword
            } catch (e: IOException) {
                // A truncated or otherwise corrupt file also surfaces as IOException.
                // Reporting it as a wrong password would trap the user on the
                // password dialog with no correct answer, so it is a real failure.
                DecryptResult.Failed(e.message ?: "This PDF appears to be damaged")
            } catch (e: Exception) {
                DecryptResult.Failed(e.message ?: "Decryption failed")
            } finally {
                inputStream.close()
            }
        }
}
