package com.modocs.core.storage

import com.modocs.core.common.DocumentType
import com.modocs.core.model.RecentFile
import com.modocs.core.storage.db.RecentFileDao
import com.modocs.core.storage.db.RecentFileEntity
import kotlinx.coroutines.flow.Flow
import kotlinx.coroutines.flow.map
import javax.inject.Inject
import javax.inject.Singleton

@Singleton
class RecentFilesRepository @Inject constructor(
    private val recentFileDao: RecentFileDao,
    @dagger.hilt.android.qualifiers.ApplicationContext private val context: android.content.Context,
) {

    fun getRecentFiles(): Flow<List<RecentFile>> {
        return recentFileDao.getRecentFiles().map { entities ->
            entities.map { it.toDomainModel() }
        }
    }

    suspend fun addRecentFile(
        uri: String,
        displayName: String,
        documentType: DocumentType,
        fileSizeBytes: Long?,
    ) {
        if (!com.modocs.core.common.AppPreferences.get(context).state.value.keepRecents) return
        val existing = recentFileDao.findByUri(uri)
        val entity = RecentFileEntity(
            id = existing?.id ?: 0,
            uri = uri,
            displayName = displayName,
            documentType = documentType.name,
            lastOpenedTimestamp = System.currentTimeMillis(),
            fileSizeBytes = fileSizeBytes,
        )
        recentFileDao.upsert(entity)
    }

    suspend fun clearHistory() = recentFileDao.deleteAll()

    suspend fun removeRecentFile(id: Long) {
        recentFileDao.deleteById(id)
    }

    private fun RecentFileEntity.toDomainModel(): RecentFile = RecentFile(
        id = id,
        uri = uri,
        displayName = displayName,
        documentType = try {
            DocumentType.valueOf(documentType)
        } catch (_: IllegalArgumentException) {
            DocumentType.UNKNOWN
        },
        lastOpenedTimestamp = lastOpenedTimestamp,
        fileSizeBytes = fileSizeBytes,
    )
}
