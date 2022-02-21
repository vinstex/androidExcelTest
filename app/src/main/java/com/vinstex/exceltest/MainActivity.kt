package com.vinstex.exceltest

import android.os.Bundle
import android.widget.TextView
import androidx.appcompat.app.AppCompatActivity
import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.*


class MainActivity : AppCompatActivity() {
    private lateinit var textView: TextView

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)
        textView = findViewById(R.id.textView)
        val ourWB = createWorkbook()
        createExcelFile(ourWB)
        updateCell()
        // deleteSheet()
        deleteRow()
        //compute()
        //deleteColumn()

    }

    private fun createWorkbook(): Workbook {
        // Creating a workbook object from the XSSFWorkbook() class
        val ourWorkbook = XSSFWorkbook()

        //Creating a sheet called "statSheet" inside the workbook and then add data to it
        val sheet: Sheet = ourWorkbook.createSheet("statSheet")
        ourWorkbook.createSheet("testSheet")
        addData(sheet)

        return ourWorkbook
    }

    private fun addData(sheet: Sheet) {

        //Creating rows at passed in indices
        val row1 = sheet.createRow(0)
        val row2 = sheet.createRow(1)
        val row3 = sheet.createRow(2)
        val row4 = sheet.createRow(3)
        val row5 = sheet.createRow(4)
        val row6 = sheet.createRow(5)
        val row7 = sheet.createRow(6)
        val row8 = sheet.createRow(7)


        //Adding data to each  cell
        createCell(row1, 0, "Name")
        createCell(row1, 1, "Score")

        createCell(row2, 0, "Mike")
        createCell(row2, 1, "470")

        createCell(row3, 0, "Montessori")
        createCell(row3, 1, "460")

        createCell(row4, 0, "Sandra")
        createCell(row4, 1, "380")

        createCell(row5, 0, "Moringa")
        createCell(row5, 1, "300")

        createCell(row6, 0, "Torres")
        createCell(row6, 1, "270")

        createCell(row7, 0, "McGee")
        createCell(row7, 1, "420")

        createCell(row8, 0, "Gibbs")
        createCell(row8, 1, "510")

    }

    //function for creating a cell.
    private fun createCell(sheetRow: Row, columnIndex: Int, cellValue: String?) {
        //create a cell at a passed in index
        val ourCell = sheetRow.createCell(columnIndex)
        //add the value to it
        ourCell?.setCellValue(cellValue)
    }

    private fun createExcelFile(ourWorkbook: Workbook) {

        //get our app file directory
        val ourAppFileDirectory = filesDir
        //Check whether it exists or not, and create if does not exist.
        if (ourAppFileDirectory != null && !ourAppFileDirectory.exists()) {
            ourAppFileDirectory.mkdirs()
        }

        //Create an excel file called test.xlsx
        val excelFile = File(ourAppFileDirectory, "test.xlsx")

        //Write a workbook to the file using a file outputstream
        try {
            val fileOut = FileOutputStream(excelFile)
            ourWorkbook.write(fileOut)
            fileOut.close()
        } catch (e: FileNotFoundException) {
            e.printStackTrace()
        } catch (e: IOException) {
            e.printStackTrace()
        }
    }

    private fun getExcelFile(): File? {

        val ourAppFileDirectory = filesDir
        ourAppFileDirectory?.let {

            //Check if file exists or not
            if (it.exists()) {
                //check the file in the directory called "test.xlsx"
                val retrievedExcel = File(ourAppFileDirectory, "test.xlsx")
                //return the file
                return retrievedExcel
            }
        }
        return null
    }

    //function for reading the workbook from the loaded spreadsheet file
    private fun retrieveWorkbook(): Workbook? {

        //Reading the workbook from the loaded spreadsheet file
        getExcelFile()?.let {
            try {

                //Reading it as stream
                val workbookStream = FileInputStream(it)

                //Return the loaded workbook
                return WorkbookFactory.create(workbookStream)
            } catch (e: Exception) {
                e.printStackTrace()
            }
        }

        return null
    }

    //function for selecting the sheet
    private fun selectSheet(): Sheet? {

        //choosing the workbook
        retrieveWorkbook()?.let { workbook ->

            //Checking the existence of a sheet
            if (workbook.numberOfSheets > 0) {

                //Return the first sheet
                return workbook.getSheetAt(0)
            }
        }

        return null
    }

    //function to compute the statistical functions
    private fun compute() {
        //get sheet
        selectSheet()?.let { sheet ->
            //finding the total number of rows
            val totalRows = sheet.physicalNumberOfRows
            val scoreArray = Array<Int>(totalRows) { 0 }
            for (i in 0 until totalRows) {
                scoreArray[i] = (sheet.getRow(i).getCell(1)).toString().toInt()
            }

            var mean = findMean(scoreArray)
            var variance = findVariance(scoreArray, mean)
            var stdDeviation: Double = Math.sqrt(variance)

            //formatting to 2 decimal places
            var meanTo2dp: String = String.format("%.2f", mean)
            var stdDeviationTo2dp = String.format("%.2f", stdDeviation)
            var varianceTo2dp: String = String.format("%.2f", variance)

            //displaying the text to the textview
            textView.setText("From the Spreadsheet, we get these:\n\n MEAN: " + meanTo2dp + "\n VARIANCE: " + varianceTo2dp + "\n STD DEVIATION: " + stdDeviationTo2dp)
        }

    }

    private fun findMean(arrayArg: Array<Int>): Double {
        var total = 0.0
        var i = 0
        for (a in arrayArg) {
            total += arrayArg[i]
            i++
        }
        var avg = total / arrayArg.size
        return avg
    }

    private fun findVariance(arrayArg: Array<Int>, mean: Double): Double {
        var indexVariance = 0.0
        var i = 0
        for (a in arrayArg) {
            indexVariance += Math.pow(((arrayArg[i].toDouble()) - mean), 2.0)
            i++
        }
        var avgVariance = indexVariance / arrayArg.size
        return avgVariance
    }

    private fun updateCell() {
        getExcelFile()?.let {
            try {

                //Reading it as stream
                val workbookStream = FileInputStream(it)

                //Return the loaded workbook
                val workbook = WorkbookFactory.create(workbookStream)
                if (workbook.numberOfSheets > 0) {

                    //Return the first sheet
                    val sheet = workbook.getSheetAt(0)
                    //choosing the first row as the headers
                    var nameHeaderCell = sheet.getRow(0).getCell(0)
                    var scoreHeaderCell = sheet.getRow(0).getCell(1)

                    //selecting cells to be editted and formatted
                    var targetCellLabel = sheet.getRow(1).getCell(0)
                    var targetCellValue = sheet.getRow(1).getCell(1)

                    val font: Font = workbook.createFont()
                    val headerCellStyle = workbook.createCellStyle()
                    val targetCellDataStyle = workbook.createCellStyle()

                    //choosing white color and a bold formatting
                    font.color = IndexedColors.WHITE.index
                    font.bold = true

                    //applying formatting styles to the cells
                    headerCellStyle.setAlignment(HorizontalAlignment.LEFT)
                    headerCellStyle.fillForegroundColor = IndexedColors.RED.index
                    headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND)
                    headerCellStyle.setFont(font);
                    nameHeaderCell.cellStyle = headerCellStyle
                    scoreHeaderCell.cellStyle = headerCellStyle

                    targetCellDataStyle.setAlignment(HorizontalAlignment.LEFT)
                    targetCellValue.cellStyle = targetCellDataStyle
                    targetCellLabel.cellStyle = targetCellDataStyle

                    //feeding in new values to the selected cells
                    targetCellLabel.setCellValue("Mitchelle")
                    targetCellValue.setCellValue(140.0)

                    workbookStream.close()

                    //saving the changes
                    try {
                        val fileOut = FileOutputStream(it)
                        workbook.write(fileOut)
                        fileOut.close()
                    } catch (e: FileNotFoundException) {
                        e.printStackTrace()
                    } catch (e: IOException) {
                        e.printStackTrace()
                    }
                }
            } catch (e: Exception) {
                e.printStackTrace()
            }
        }
    }

    private fun deleteSheet() {
        getExcelFile()?.let {
            try {

                //Reading it as stream
                val workbookStream = FileInputStream(it)

                //Return the loaded workbook
                val workbook = WorkbookFactory.create(workbookStream)
                if (workbook.numberOfSheets > 0) {
                    //removing the second sheet
                    workbook.removeSheetAt(1)
                    workbookStream.close()

                    try {
                        val fileOut = FileOutputStream(it)
                        workbook.write(fileOut)
                        fileOut.close()
                    } catch (e: FileNotFoundException) {
                        e.printStackTrace()
                    } catch (e: IOException) {
                        e.printStackTrace()
                    }
                }
            } catch (e: Exception) {
                e.printStackTrace()
            }
        }
    }

    private fun deleteRow() {
        getExcelFile()?.let {
            try {
                val rowNo = 1
                //Reading it as stream
                val workbookStream = FileInputStream(it)

                //Return the loaded workbook
                val workbook = WorkbookFactory.create(workbookStream)
                if (workbook.numberOfSheets > 0) {

                    //Return the first sheet
                    val sheet = workbook.getSheetAt(0)

                    //getting the total number of rows available
                    val totalNoOfRows = sheet.lastRowNum

                    val targetRow = sheet.getRow(rowNo)
                    if (targetRow != null) {
                        sheet.removeRow(targetRow)
                    }

                    /*excluding the last row, move the cells that come
                    after the deleted row one step behind*/
                    if (rowNo >= 0 && rowNo < totalNoOfRows) {
                        sheet.shiftRows(rowNo + 1, totalNoOfRows, -1)
                    }
                    workbookStream.close()

                    try {
                        val fileOut = FileOutputStream(it)
                        workbook.write(fileOut)
                        fileOut.close()
                    } catch (e: FileNotFoundException) {
                        e.printStackTrace()
                    } catch (e: IOException) {
                        e.printStackTrace()
                    }
                }
            } catch (e: Exception) {
                e.printStackTrace()
            }
        }
    }

    private fun deleteColumn() {
        getExcelFile()?.let {
            try {
                val colNo = 0
                //Reading it as stream
                val workbookStream = FileInputStream(it)

                //Return the loaded workbook
                val workbook = WorkbookFactory.create(workbookStream)
                if (workbook.numberOfSheets > 0) {

                    //Return the first sheet
                    val sheet = workbook.getSheetAt(0)
                    val totalRows = sheet.lastRowNum
                    val row = sheet.getRow(colNo)
                    val maxCell = row.lastCellNum.toInt()
                    if (colNo >= 0 && colNo <= maxCell) {
                        for (rowNo in 0..totalRows) {
                            val targetCol: Cell = sheet.getRow(rowNo).getCell(colNo)
                            if (targetCol != null) {
                                sheet.getRow(rowNo).removeCell(targetCol);
                            }
                        }
                    }
                    workbookStream.close()

                    try {
                        val fileOut = FileOutputStream(it)
                        workbook.write(fileOut)
                        fileOut.close()
                    } catch (e: FileNotFoundException) {
                        e.printStackTrace()
                    } catch (e: IOException) {
                        e.printStackTrace()
                    }
                }
            } catch (e: Exception) {
                e.printStackTrace()
            }
        }
    }

}