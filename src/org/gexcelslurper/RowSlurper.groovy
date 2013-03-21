package org.gexcelslurper

import org.apache.poi.ss.usermodel.*

class RowSlurper {
    Row row
    List<String> labels
    int rowIndex

    RowSlurper(Row row, List<String> labels) {
        this.row = row
        this.labels = labels
        this.rowIndex = row.rowNum
    }

    def cell(def idx) {
        if (labels && (idx instanceof String)) {
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
}
