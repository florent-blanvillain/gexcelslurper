package org.gexcelslurper

import org.apache.poi.ss.usermodel.*

class ExcelSlurper {
    Workbook workbook
    Sheet sheetUnderIteration = null

    ExcelSlurper(String filePath) {
        this(new File(filePath))
    }

    ExcelSlurper(File file) {
        this(file.newInputStream())
    }

    ExcelSlurper(URL url) {
        this(url.openStream())
    }

    ExcelSlurper(InputStream is) {
        workbook = WorkbookFactory.create(is)
    }

    def eachLine(Map params = [:], Closure closure) {
        int offset = params.offset ?: 0
        int max = params.max ?: 9999999
        Sheet sheet = sheetUnderIteration ?: getSheet(params.sheet)
        Iterator<Row> rowIterator = sheet.rowIterator()
        def linesRead = 0

        List<String> labels = null
        if (params.labels) {
            labels = rowIterator.next().collect { it.toString().toLowerCase() }
        }

        offset.times { rowIterator.next() }

        while (rowIterator.hasNext() && linesRead++ < max) {
            Row row = rowIterator.next()

            closure.setDelegate(new RowSlurper(row, labels))
            closure.call(linesRead)
        }
    }

    def eachSheet(Map params = [:], Closure closure) {
        int numberOfSheets = workbook.numberOfSheets
        int offset = params.offset ?: 0
        int max = (params.max ?: numberOfSheets) - 1

        (offset..max).each { int i ->
            Sheet sheet = getSheet(i)
            sheetUnderIteration = sheet
            closure.setDelegate(new SheetSlurper(sheet, this))
            closure.call(i)
        }

        sheetUnderIteration = null
    }

    List<String> getSheetNames() {
        def names = []
        assert (0..<workbook.numberOfSheets).asList() == [0,1]
        for (i in 0..<workbook.numberOfSheets) {
            names << workbook.getSheetAt(i).sheetName
        }
        names
    }

    protected Sheet getSheet(def idx) {
        Sheet sheet
        if (!idx) idx = 0
        if (idx instanceof Number) {
            sheet = workbook.getSheetAt(idx)
        } else if (idx ==~ /^\d+$/) {
            sheet = workbook.getSheetAt(Integer.valueOf(idx))
        } else {
            sheet = workbook.getSheet(idx)
        }
        return sheet
    }
}
