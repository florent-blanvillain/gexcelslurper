package org.gexcelslurper

import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet

class SheetSlurper {
    Sheet sheet
    XlsWorkbookSlurper slurper
    String sheetName
    int sheetIndex
    List<String> labels

    SheetSlurper(Sheet sheet, XlsWorkbookSlurper slurper) {
        this.sheet = sheet
        this.slurper = slurper
        this.sheetName = sheet.sheetName
        this.sheetIndex = sheet.workbook.getSheetIndex(sheetName)
        this.labels = this[0]?.row*.toString()
    }

    /**
     * @param params parameters:
     * <ul>
     * <li>offset</li>
     * <li>labels: "labels:true" same as "offset:1"</li>
     * <li>max</li>
     * </ul>
     * @param closure
     */
    def eachRow(Map params = [:], Closure closure) {
        int offset = params.offset ?: 0
        int max = params.max ?: 9999999
        Iterator<Row> rowIterator = sheet.rowIterator()
        def linesRead = 0

        if (params.labels) {
            rowIterator.next()
        }

        offset.times { rowIterator.next() }

        while (rowIterator.hasNext() && linesRead++ < max) {
            Row row = rowIterator.next()

            closure.setDelegate(new RowSlurper(row, this))
            closure.call(linesRead)
        }
    }

    def getAt(int index) {
        new RowSlurper(sheet.getRow(index), this)
    }


    List toList(Map params = [:]) {
        List list = []

        eachRow(params) {
            list << delegate.toList()
        }

        list
    }

    List findAll(Map params = [:], Closure closure) {
        List list = []
        def linesRead = 0

        eachRow(params) {
            closure.setDelegate(delegate)
            if (closure.call(linesRead++)) {
                list << delegate.toList()
            }
        }

        list
    }

    RowSlurper findRowSlurper(Map params = [:], Closure closure) {
        RowSlurper rowSlurperResult = null

        int offset = params.offset ?: 0
        int max = params.max ?: 9999999
        Iterator<Row> rowIterator = sheet.rowIterator()
        def linesRead = 0

        if (params.labels) {
            rowIterator.next()
        }

        offset.times { rowIterator.next() }

        while (rowIterator.hasNext() && linesRead++ < max && !rowSlurperResult) {
            Row row = rowIterator.next()
            def rs = new RowSlurper(row, this)
            closure.setDelegate(rs)
            if (closure.call()) {
                rowSlurperResult = rs
            }
        }
        rowSlurperResult
    }

    List find(Map params = [:], Closure closure) {
        findRowSlurper(params, closure)?.toList()
    }
}
