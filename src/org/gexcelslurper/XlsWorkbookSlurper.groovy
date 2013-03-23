package org.gexcelslurper

import org.apache.poi.ss.usermodel.*

class XlsWorkbookSlurper {
    Workbook workbook

    XlsWorkbookSlurper(String filePath) {
        this(new File(filePath))
    }

    XlsWorkbookSlurper(File file) {
        this(file.newInputStream())
    }

    XlsWorkbookSlurper(URL url) {
        this(url.openStream())
    }

    XlsWorkbookSlurper(InputStream is) {
        workbook = WorkbookFactory.create(is)
    }

    def eachSheet(Map params = [:], Closure closure) {
        int numberOfSheets = workbook.numberOfSheets
        int offset = params.offset ?: 0
        int max = (params.max ?: numberOfSheets) - 1

        (offset..max).each { int i ->
            Sheet sheet = getSheet(i)
            closure.setDelegate(new SheetSlurper(sheet, this))
            closure.call(i)
        }
    }

    /**
     * @param params  parameters:
     * <ul>
     * <li>offset</li>
     * <li>labels: "labels:true" same as "offset:1"</li>
     * <li>max</li>
     * <li>sheet</li>
     * </ul>
     * @param closure
     */
    def eachRow(Map params = [:], Closure closure) {
            new SheetSlurper(getSheet(params.sheet) ,this).eachRow(params, closure)
    }

    def getAt(def index) {
        new SheetSlurper(getSheet(index) ,this)
    }

    def propertyMissing(String name) {
        new SheetSlurper(getSheet(name) ,this)
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
