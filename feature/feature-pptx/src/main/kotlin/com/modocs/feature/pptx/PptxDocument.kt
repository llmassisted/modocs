package com.modocs.feature.pptx

import android.graphics.Color as AndroidColor

/**
 * In-memory representation of a parsed PPTX file.
 */
data class PptxDocument(
    val slides: List<PptxSlide>,
    val slideWidth: Long = 9144000,   // EMU, default 10 inches
    val slideHeight: Long = 6858000,  // EMU, default 7.5 inches
    val rawEntries: Map<String, ByteArray> = emptyMap(),
    val warnings: List<String> = emptyList(),
    /** Resolved theme color scheme from theme1.xml: name (dk1, lt1, accent1, etc.) -> ARGB int. */
    val themeColors: Map<String, Int> = emptyMap(),
)

data class PptxSlide(
    val slideNumber: Int,
    val shapes: List<PptxShape>,
    /** Background color, or null for default white. */
    val backgroundColor: Int? = null,
    /** Embedded images: relationship ID -> byte array. */
    val images: Map<String, ByteArray> = emptyMap(),
    /** Shapes from slide layout/master rendered behind the slide's own shapes. */
    val backgroundShapes: List<PptxShape> = emptyList(),
)

sealed interface PptxShape {
    val x: Long       // EMU
    val y: Long       // EMU
    val width: Long   // EMU
    val height: Long  // EMU
}

data class PptxTextBox(
    override val x: Long,
    override val y: Long,
    override val width: Long,
    override val height: Long,
    val paragraphs: List<PptxParagraph>,
    val fillColor: Int? = null,
    val borderColor: Int? = null,
    val borderWidthEmu: Long = 0,
    val rotation: Float = 0f,
) : PptxShape

data class PptxImageShape(
    override val x: Long,
    override val y: Long,
    override val width: Long,
    override val height: Long,
    val relationId: String,
    val rotation: Float = 0f,
) : PptxShape

data class PptxRectangle(
    override val x: Long,
    override val y: Long,
    override val width: Long,
    override val height: Long,
    val fillColor: Int? = null,
    val borderColor: Int? = null,
    val borderWidthEmu: Long = 0,
    val cornerRadiusEmu: Long = 0,
    val rotation: Float = 0f,
    val text: List<PptxParagraph> = emptyList(),
) : PptxShape

data class PptxLine(
    override val x: Long,
    override val y: Long,
    override val width: Long,
    override val height: Long,
    val color: Int = AndroidColor.BLACK,
    val lineWidthEmu: Long = 12700, // 1pt default
) : PptxShape

data class PptxParagraph(
    val runs: List<PptxRun>,
    val alignment: PptxAlignment = PptxAlignment.LEFT,
    val spacingBeforePt: Float = 0f,
    val spacingAfterPt: Float = 0f,
    val bulletChar: String? = null,
    val level: Int = 0,
) {
    val text: String get() = runs.joinToString("") { it.text }
}

data class PptxRun(
    val text: String,
    val bold: Boolean = false,
    val italic: Boolean = false,
    val underline: Boolean = false,
    val fontSizePt: Float? = null,
    val fontColor: Int? = null,
    val fontName: String? = null,
)

enum class PptxAlignment {
    LEFT, CENTER, RIGHT, JUSTIFY
}

object PptxColors {
    fun parseColor(colorStr: String?): Int? {
        if (colorStr.isNullOrBlank()) return null
        return try {
            val hex = when (colorStr.length) {
                6 -> "FF$colorStr"
                8 -> colorStr
                else -> return null
            }
            AndroidColor.parseColor("#$hex")
        } catch (_: Exception) {
            null
        }
    }

    fun themeColor(index: Int): Int? = when (index) {
        0 -> AndroidColor.parseColor("#FFFFFFFF")  // lt1
        1 -> AndroidColor.parseColor("#FF000000")  // dk1
        2 -> AndroidColor.parseColor("#FFE7E6E6")  // lt2
        3 -> AndroidColor.parseColor("#FF44546A")  // dk2
        4 -> AndroidColor.parseColor("#FF4472C4")  // accent1
        5 -> AndroidColor.parseColor("#FFED7D31")  // accent2
        6 -> AndroidColor.parseColor("#FFA5A5A5")  // accent3
        7 -> AndroidColor.parseColor("#FFFFC000")  // accent4
        8 -> AndroidColor.parseColor("#FF5B9BD5")  // accent5
        9 -> AndroidColor.parseColor("#FF70AD47")  // accent6
        else -> null
    }
}

object PptxUnits {
    /** EMU to dp (1 EMU = 1/914400 inch, 1dp ≈ 1/160 inch). */
    fun emuToDp(emu: Long): Float = emu / 914400f * 160f

    /** EMU to points (1 EMU = 1/12700 point). */
    fun emuToPt(emu: Long): Float = emu / 12700f
}
