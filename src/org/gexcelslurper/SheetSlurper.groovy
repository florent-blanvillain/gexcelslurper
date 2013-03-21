package org.gexcelslurper

import org.apache.poi.ss.usermodel.*

class SheetSlurper {
    Sheet sheet
    ExcelSlurper slurper
    String sheetName
    int sheetIndex
    def eachLine

    SheetSlurper(Sheet sheet, ExcelSlurper slurper) {
        this.sheet = sheet
        this.slurper = slurper
        this.sheetName = sheet.sheetName
        this.sheetIndex = sheet.workbook.getSheetIndex(sheetName)
        this.eachLine = slurper.&eachLine
    }
}
