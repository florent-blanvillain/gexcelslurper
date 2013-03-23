package org.gexcelslurper

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.DataFormatter
import org.apache.poi.ss.usermodel.DateUtil
import org.apache.poi.ss.usermodel.Row

class RowSlurper {
    Row row
    SheetSlurper sheetSlurper
    List<String> labels
    int rowIndex

    RowSlurper(Row row, SheetSlurper sheetSlurper) {
        this.row = row
        this.sheetSlurper = sheetSlurper
        this.labels = sheetSlurper.labels
        this.rowIndex = row.rowNum
    }

    def cell(def idx) {
        if (idx instanceof String) {
            idx = labels.indexOf(idx.toLowerCase())
        }
        def cell = row.getCell(idx)
        if (!cell) {
            return null
        }

        def value
        switch (cell.cellType) {
            case Cell.CELL_TYPE_NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    value = cell.dateCellValue
                } else {
                    value = new DataFormatter().formatCellValue(cell)
                }
                break
            case Cell.CELL_TYPE_BOOLEAN:
                value = cell.booleanCellValue
                break
            default:
                value = cell.stringCellValue
                break
        }
        return value
    }

    def propertyMissing(String name) {
        cell(name)
    }

    def getAt(def index) {
        cell(index)
    }
}
