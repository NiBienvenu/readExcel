package org.example.poi

import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.io.File

fun main() {
    println("Hello World!")
    try {
        val fileexcel = File("/home/bienvenu/Downloads/readExcel/src/main/kotlin/poi/excel.xlsx")
        readColmnExcelFile(fileexcel)
    }catch (e: Exception) {
        println("Error reading Excel file: ${e.message}")
        e.printStackTrace()
        return
    }

}

fun readExcelFile(file: File) {
    val workbook = WorkbookFactory.create(file)
    val sheet = workbook.getSheetAt(0) // Get the first sheet

    for (row in sheet) {
        for (cell in row) {
            when (cell.cellType) {
               CellType.STRING -> print("${cell.stringCellValue}\t")
               CellType.NUMERIC -> print("${cell.numericCellValue}\t")
                CellType.BOOLEAN -> print("${cell.booleanCellValue}\t")
                else -> print("\t")
            }
        }
        println()
    }
    workbook.close()
}
fun readColmnExcelFile(file: File) {
    val workbook = WorkbookFactory.create(file)
    val sheet = workbook.getSheetAt(0) // Get the first sheet

    val rowCount = sheet.lastRowNum + 1


    for (i in 0 until rowCount) {
        val row = sheet.getRow(i)
        if (row != null) {
            val columnCount = row.physicalNumberOfCells
            for (j in 0 until columnCount) {
                val cell = row.getCell(j, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK)
                when (cell.cellType) {
                    CellType.STRING -> print("${cell.stringCellValue}\t")
                    CellType.NUMERIC -> print("${cell.numericCellValue}\t")
                    CellType.BOOLEAN -> print("${cell.booleanCellValue}\t")
                    else -> print("\t")
                }
            }
            println()
        }

    }
    workbook.close()
}