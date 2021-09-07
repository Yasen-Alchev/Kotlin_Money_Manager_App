package com.example.nexttry

import android.content.Context
import android.util.Log
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream
import android.os.Environment
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.Workbook
import java.text.SimpleDateFormat
import java.util.*
import java.util.Locale





class ExcelFunctions(private val context: Context) {

    private fun Sheet.createItem(item: Item, count: Int){
        // Evaluate all formulas in the sheet, to update their value
        val formulaEvaluator = workbook.creationHelper.createFormulaEvaluator()

        //Init the empty sheet
        if(this.getRow(0)?.getCell(0) == null){
            this.createRow(0).apply {
                createCell(0).setCellValue("Item")
                createCell(1).setCellValue(item.name)
            }
            this.createRow(1).apply {
                createCell(0).setCellValue("Price")
                createCell(1).setCellValue(item.price.toString().toDouble())
            }
            this.createRow(2).apply {
                createCell(0).setCellValue(getTime())
                createCell(1).setCellValue(count.toDouble())
            }
            this.createRow(3).apply {
                createCell(0).setCellValue("Total")
                createCell(1).cellFormula = getRow(1).getCell(1).address.toString() + "* SUM(" + getRow(2).getCell(1).address.toString() + ":" + getRow(2).getCell(1).address.toString() + ")"
                createCell(2).cellFormula = "SUM(" + getCell(1).address.toString() + ":" + getCell(1).address.toString() + ")"
            }
        }
        else //if the sheet exists
        {
            val cellIndex = this.getRow(0).lastCellNum.toInt()
            val lastRowIndex = this.lastRowNum  // the "Total"'s row index
            val row = findTodaysRow(this)
            val containsItemName = findItemCellIndex(this, item.name.toString())

            if(containsItemName != -1 && this.getRow(1).getCell(containsItemName).toString().toDouble() == item.price)
            {
                this.shiftRows(lastRowIndex, lastRowIndex, 1) //shift total's row down by one row and leave empty row for the new date row

                //crete new row to the now free total's old row and set the item count to 1
                this.createRow(lastRowIndex).apply {
                    for(i in 0..cellIndex)
                    {
                        createCell(i)
                        if(i == containsItemName){
                            getCell(i).setCellValue(count.toDouble())
                        }
                    }
                    getCell(0).setCellValue(getTime())
                }
            }
            else
            {
                //write the new item right-side to the last one
                this.getRow(0).createCell(cellIndex).setCellValue(item.name)
                this.getRow(1).createCell(cellIndex).setCellValue(item.price.toString().toDouble())

                //if today is a new day, shift the total's row down and insert the new date' row over the total's row new one
                if(row == -1)
                {
                    this.shiftRows(lastRowIndex, lastRowIndex, 1) //shift total's row down by one row and leave empty row for the new date row

                    //crete new row to the now free total's old row and set the item count to 1
                    this.createRow(lastRowIndex).apply {
                        createCell(0).setCellValue(getTime())
                        createCell(cellIndex).setCellValue(count.toDouble())
                    }

                    // now the formulas are different because the columns and rows of total's row changed
                    // we now write all of them from the scratch for each cell of the new total's row
                    for((index, cell) in this.getRow(lastRowIndex + 1).cellIterator().withIndex())
                    {
                        if(index == 0) continue //skip the first cell where it contains the text "Total" that cannot be converted to numeric

                        val priceBox = this.getRow(1).getCell(index) ?: this.getRow(1).createCell(index)
                        val startBox = this.getRow(2).getCell(index) ?: this.getRow(2).createCell(index)
                        val endBox = this.getRow(lastRowIndex).getCell(index) ?: this.getRow(lastRowIndex).createCell(index)

                        cell.cellFormula = priceBox.address.toString() + "* SUM(" + startBox.address.toString() + ":" + endBox.address.toString() + ")"
                    }
                    this.getRow(lastRowIndex + 1).createCell(cellIndex + 1).cellFormula = "SUM(" + this.getRow(lastRowIndex + 1).getCell(1).address.toString() + ":" + this.getRow(lastRowIndex + 1).getCell(cellIndex).address.toString() + ")"
                }
                else
                {
                    val lastRow = this.lastRowNum

                    this.getRow(row).createCell(cellIndex)
                    //shift the SUM cell
                    this.getRow(lastRow).shiftCellsRight(cellIndex, cellIndex, 1)
                    //init the new cell where SUM was
                    this.getRow(lastRow).createCell(cellIndex)

                    for(i in 0..lastRow)
                    {
                        if(this.getRow(i).getCell(cellIndex) == null)
                        {
                            this.getRow(i).createCell(cellIndex)
                        }
                    }

                    //the new total formula to the new cell where SUM was
                    this.getRow(lastRow).createCell(cellIndex).cellFormula =
                        this.getRow(1).getCell(cellIndex).address.toString() +
                                "* SUM(" +  this.getRow(2).getCell(cellIndex).address.toString() +
                                ":" + this.getRow(row).createCell(cellIndex).address.toString() + ")"
                    //calculate the SUM
                    this.getRow(lastRow).createCell(cellIndex + 1).cellFormula = "SUM(" + this.getRow(lastRow).getCell(1).address.toString() + ":" + this.getRow(lastRow).getCell(cellIndex).address.toString() + ")"
                    //set the cell to 1
                    this.getRow(row).createCell(cellIndex).setCellValue(count.toDouble())
                }
            }
        }
        formulaEvaluator.evaluateAll()
    }

    private fun findTodaysRow(sheet: Sheet): Int
    {
        for ((index, row) in sheet.rowIterator().withIndex()){
            for(cell in row.cellIterator()){
                if(cell.toString().compareTo(getTime()) == 0)
                {
                    return index
                }
            }
        }
        return -1
    }

    private fun findItemCellIndex(sheet: Sheet, item: String): Int
    {
        for((index, cell) in sheet.getRow(0).cellIterator().withIndex()){
            if (cell.toString().compareTo(item) == 0){
                return index
            }
        }
        return -1
    }

    private fun getTime(): String
    {

        val currentDate = SimpleDateFormat("dd-MM-yyyy", Locale.getDefault()).format(Date())
        return currentDate.toString()
        //return LocalDateTime.now().toLocalDate().toString()
    }

    fun writeToFile(filename: String, item: Item, count: Int)
    {
        if (!isExternalStorageAvailable() || isExternalStorageReadOnly()) {
            Log.w("FileUtils", "Storage not available or read only")
            return
        }

        val file = File(context.getExternalFilesDir(null),  filename)
        if(file.exists())
        {
            val input = FileInputStream(file)
            val wb: Workbook = HSSFWorkbook(input)
            val sheet = wb.getSheet("Razhodi") ?: wb.getSheetAt(0)
            val row = findTodaysRow(sheet)
            val cell = findItemCellIndex(wb.getSheet("Razhodi"), item.name.toString())

            //create the new item
            if(cell == -1) {
                sheet.createItem(item, count)
            }
            else //add +1 to the current item count
            {
                val itemPrice:Double = sheet.getRow(1).getCell(cell).toString().toDouble()

                if(itemPrice == item.price) {
                    if(row != -1) {
                        sheet.getRow(row).apply {
                            if(getCell(cell).toString() == "")
                            {
                                getCell(cell).setCellValue("1".toDouble())
                            }else{
                                getCell(cell).setCellValue((getCell(cell).toString().toDouble() + count.toDouble()).toString().toDouble())
                            }
                        }
                    }
                    else
                    {
                        sheet.createItem(item, count)
                    }
                }
                else //the item price is not the same. The item is different
                {
                    println("Item: ${item.name} with price: $itemPrice Already Exists!")
                }
            }
            val outputStream = FileOutputStream(file)
            wb.write(outputStream)
            outputStream.close()
            wb.close()
        }
        else
        {
            val xlWb: Workbook = HSSFWorkbook()
            val xlWs = xlWb.createSheet()

            file.createNewFile()

            xlWb.setSheetName(0, "Razhodi")
            xlWs.createItem(item, count)

            val outputStream = FileOutputStream(file)

            xlWb.write(outputStream)
            outputStream.close()
            xlWb.close()
        }
    }

    private fun Sheet.printSheet(){
        println("-----${this.sheetName}-----")
        for (row in this.rowIterator()){
            for(cell in row.cellIterator()){
                print("\t$cell  ")
            }
            println()
        }
    }

    fun readFromFile(filename: String)
    {
        // check if available and not read only
        if (!isExternalStorageAvailable() || isExternalStorageReadOnly()) {
            Log.w("FileUtils", "Storage not available or read only")
            return
        }

        val file = File(filename)
        if(file.exists()) {
            val input = FileInputStream(filename)
            val xlWb = WorkbookFactory.create(input)

            val iter = xlWb.sheetIterator()

            while (iter.hasNext()) {
                iter.next().printSheet()
            }
        }
        else {
            println("File is empty")
        }
    }

    private fun isExternalStorageReadOnly(): Boolean {
        val extStorageState = Environment.getExternalStorageState()
        return Environment.MEDIA_MOUNTED_READ_ONLY == extStorageState
    }

    private fun isExternalStorageAvailable(): Boolean {
        val extStorageState = Environment.getExternalStorageState()
        return Environment.MEDIA_MOUNTED == extStorageState
    }

}
