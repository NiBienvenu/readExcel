package util
import java.io.File

fun main() {
    val file = File("/home/bienvenu/Downloads/readExcel/src/main/kotlin/poi/excel.xlsx")
    readExcelFile(file)
}

fun readExcelFile(file: File) {
    val workbook = readWorkbook(file)
    val sheet = workbook.getSheetAt(0)

    for (i in 0..sheet.lastRowNum) {
        val row = sheet.getRow(i)
        if (row != null) {
            for (j in 0 until row.lastCellNum) {
                val cell = row.getCell(j, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK)
                when (getCellType(cell)) {
                    CellType.STRING -> print("${getCellStringValue(cell)}\t")
                    CellType.NUMERIC -> print("${getCellNumericValue(cell)}\t")
                    CellType.BOOLEAN -> print("${getCellBooleanValue(cell)}\t")
                    else -> print("\t")
                }
            }
            println()
        }
    }
}

fun readWorkbook(file: File): Workbook {
    // Utiliser la bibliothèque standard de Kotlin pour lire le fichier Excel
    val inputStream = file.inputStream()
    return readXlsxWorkbook(inputStream)
}

fun readXlsxWorkbook(inputStream: java.io.InputStream): Workbook {
    // Implémentation de la lecture du fichier Excel en utilisant la bibliothèque standard de Kotlin
    // ...
    return Workbook()
}

class Workbook {
    fun getSheetAt(index: Int): Sheet {
        // Implémentation de la récupération de la feuille de calcul
        return Sheet()
    }
}

class Sheet {
    val lastRowNum: Int = 0

    fun getRow(index: Int): Row? {
        // Implémentation de la récupération de la ligne
        return Row(index)
    }
}

class Row(private val index: Int) {
    val lastCellNum: Int = 0

    fun getCell(index: Int, missingCellPolicy: Int): Cell {
        // Implémentation de la récupération de la cellule
        return Cell()
    }

    object MissingCellPolicy {
        const val CREATE_NULL_AS_BLANK = 1
    }
}

class Cell {
    fun getCellType(): CellType {
        // Implémentation de la récupération du type de la cellule
        return CellType.STRING
    }

    fun getStringCellValue(): String {
        // Implémentation de la récupération de la valeur de la cellule en tant que chaîne de caractères
        return ""
    }

    fun getNumericCellValue(): Double {
        // Implémentation de la récupération de la valeur de la cellule en tant que nombre
        return 0.0
    }

    fun getBooleanCellValue(): Boolean {
        // Implémentation de la récupération de la valeur de la cellule en tant que booléen
        return false
    }
}

enum class CellType {
    STRING, NUMERIC, BOOLEAN
}

fun getCellType(cell: Cell): CellType {
    return cell.getCellType()
}

fun getCellStringValue(cell: Cell): String {
    return cell.getStringCellValue()
}

fun getCellNumericValue(cell: Cell): Double {
    return cell.getNumericCellValue()
}

fun getCellBooleanValue(cell: Cell): Boolean {
    return cell.getBooleanCellValue()
}