package com.vinstex.exceltest

import androidx.appcompat.app.AppCompatActivity
import android.os.Bundle
import android.widget.TextView
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
        compute()

    }

    private fun createWorkbook(): Workbook {
        // Creating a workbook object from the XSSFWorkbook() class
        val ourWorkbook = XSSFWorkbook()

        //Creating a sheet called "statSheet" inside the workbook and then add data to it
        val sheet: Sheet = ourWorkbook.createSheet("statSheet")
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


        //Adding data to each  cell
        createCell(row1, 0, "Mike")
        createCell(row1, 1, "470")

        createCell(row2, 0, "Montessori")
        createCell(row2, 1, "460")

        createCell(row3, 0, "Sandra")
        createCell(row3, 1, "380")

        createCell(row4, 0, "Moringa")
        createCell(row4, 1, "300")

        createCell(row5, 0, "Torres")
        createCell(row5, 1, "270")

        createCell(row6, 0, "McGee")
        createCell(row6, 1, "420")

        createCell(row7, 0, "Gibbs")
        createCell(row7, 1, "510")
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
                var createdExcel = File(ourAppFileDirectory, "test.xlsx")
                //return the file
                return createdExcel
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

            //setting text to the textview
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
}