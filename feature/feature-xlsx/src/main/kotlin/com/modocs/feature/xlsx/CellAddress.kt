package com.modocs.feature.xlsx

fun cellReference(row: Int, column: Int): String {
    var col = column + 1
    var letters = ""
    while (col > 0) {
        letters = ('A' + (col - 1) % 26) + letters
        col = (col - 1) / 26
    }
    return letters + (row + 1)
}

fun parseCellReference(value: String): Pair<Int, Int>? {
    val match = Regex("^([A-Za-z]{1,3})([1-9][0-9]{0,6})$").matchEntire(value.trim()) ?: return null
    val col = match.groupValues[1].uppercase().fold(0) { n, c -> n * 26 + (c - 'A' + 1) }
    val row = match.groupValues[2].toInt()
    return if (col in 1..16384 && row in 1..1048576) row - 1 to col - 1 else null
}
