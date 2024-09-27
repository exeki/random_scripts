package ru.kazantsev.random_scripts

import groovy.transform.Field
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.DataFormatter
import org.apache.poi.ss.usermodel.DateUtil
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory

//Наименование для разделенных файлов
@Field static String TARGET_DOC_TITLE = 'Разделенный файл'
//Путь до разделяемого файла
@Field static String PATH_TO_FILE = 'C:\\projects\\splitXlsx\\src\\main\\resources\\Products for import v2.xlsx'
//Путь до директории, куда будут сохраняться разделяемые файлы
@Field static String PATH_TO_OUT_DIR = 'C:\\projects\\splitXlsx\\src\\main\\resources'
//Максимальный размер разделенных файлов
@Field static Integer ROW_PER_DOC = 20000
//Индексы столбцов делимого файла, которые будут приведены к строке
@Field static List<Integer> FORCE_CELL_FORMAT_TO_STRING = []

/**
 * Получает данные из ячейки
 * нужен тк для разных типов данных испольщуются разные методы
 * @param cell ячейка
 * @return значение ячейки
 */
def static getValue(Cell cell) {
    switch (cell.getCellType()) {
        case CellType.STRING:
            return cell.getRichStringCellValue().getString()
        case CellType.NUMERIC:
            if (DateUtil.isCellDateFormatted(cell)) {
                return cell.getDateCellValue()
            } else {
                return cell.getNumericCellValue()
            }
        case CellType.BOOLEAN:
            return cell.getBooleanCellValue()
        case CellType.ERROR:
            return cell.getErrorCellValue()
        case CellType.BLANK:
            return ""
        case CellType._NONE:
            return cell.getStringCellValue()
        case CellType.FORMULA:
            return cell.getCellFormula()
        default:
            throw new Exception("Неизвестный тип данных \"${cell.getCellType().toString()}\" в ячейке: строка: ${cell.getRow().getRowNum().toString()}, столбец: ${cell.getColumnIndex().toString()}")
    }
}

static Sheet createWorkbook(Sheet bigSheet, Integer rowSize) {
    Workbook newWorkbook = WorkbookFactory.create(true)
    Sheet newSheet = newWorkbook.createSheet()
    copyRowToSheet(bigSheet.getRow(0), newSheet, 0, rowSize)
    return newSheet
}

static copyRowToSheet(Row rowToCopy, Sheet sheetToWrite, Integer newRowIndex, Integer rowSize) {
    println("ДОБАВЛЯЮ СТРОКУ ИСХОДНОГО ДОКУМЕНТА " + rowToCopy.getRowNum().toString())
    Row newRow = sheetToWrite.createRow(newRowIndex)
    DataFormatter dataFormatter = new DataFormatter();
    for (int i = 0; i <= rowSize; i++) {
        Cell cell = rowToCopy.getCell(i)
        Cell newCell
        if(cell != null) {
            newCell = newRow.createCell(i)
            if(i in FORCE_CELL_FORMAT_TO_STRING) {
                newCell.setCellValue(dataFormatter.formatCellValue(cell))
            } else {
                newCell.setCellValue(getValue(cell))
            }
        }

    }
}

static void saveWorkbook(Sheet sheet, Integer docNumber) {
    String name = TARGET_DOC_TITLE + "_" +
            docNumber.toString() +
            "_size_" + sheet.getLastRowNum().toString() + ".xlsx"
    String path = PATH_TO_OUT_DIR + "\\" + name
    FileOutputStream outputStream = new FileOutputStream(path);
    sheet.getWorkbook().write(outputStream);
    outputStream.close();
    println("СОХРАНЯЮ " + path)
}

println("ОПЯТЬ РАБОТАТЬ")
Integer docNumber = 1
File file = new File(PATH_TO_FILE)
Workbook bigWorkbook = WorkbookFactory.create(file)
Sheet bigSheet = bigWorkbook.getSheetAt(0)
println("СТРОК В ДОКУМЕНТЕ " + bigSheet.getLastRowNum())

Integer rowSize = bigSheet.getRow(0).getLastCellNum()
Integer currentSize = ROW_PER_DOC
Sheet newSheet = createWorkbook(bigSheet, rowSize)
Integer newSheetIndex = 1
bigSheet.rowIterator().eachWithIndex { Row item, Integer index ->
    if (index == 0) return
    copyRowToSheet(item, newSheet, newSheetIndex, rowSize)
    newSheetIndex++
    if (index == currentSize) {
        saveWorkbook(newSheet, docNumber)
        docNumber++
        currentSize += ROW_PER_DOC
        newSheet = createWorkbook(bigSheet, rowSize)
        newSheetIndex = 1
    }
}
saveWorkbook(newSheet, docNumber)
println("ГОТОВО ХОЗЯИН")
