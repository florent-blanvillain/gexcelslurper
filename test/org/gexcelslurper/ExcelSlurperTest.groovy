package org.gexcelslurper

import org.junit.Test

import static org.junit.Assert.assertEquals

public class ExcelSlurperTest {
    static final String XLSX = 'gexcelslurper.xlsx'
    static final URL XLSX_URL = ClassLoader.systemClassLoader.getResource(XLSX)
    static final URL XLS_URL = ClassLoader.systemClassLoader.getResource('gexcelslurper.xls')
    static final XlsWorkbookSlurper xlsWorkbookSlurper = new XlsWorkbookSlurper(XLSX_URL)

    @Test
    void "load xls or xlsx files from File or URL"() {
        assert new XlsWorkbookSlurper(new File(this.class.classLoader.getResource(XLSX).file)).breaking[2].character == 'jesse'

        new XlsWorkbookSlurper(XLSX_URL).eachRow([max: 1]) {
            assert rowIndex == 0
        }

        assert new XlsWorkbookSlurper(XLS_URL)[0][0][0] == 'I can read old xls'
    }

    @Test
    void "direct access to cells by indexes of names if any"() {
        // access to cells : first index is the sheet, second is the row, third is the column
        // indexes are always 0-based
        assertEquals 'jesse', xlsWorkbookSlurper[0][2][1]

        // 'breaking' is the name of the first sheet and 'character' is the value of the first cell of the second row.
        // first row cells are automatically taken as labels for columns
        assertEquals 'jesse', xlsWorkbookSlurper.breaking[2].character
    }

    @Test
    void sheetNames() {
        assert xlsWorkbookSlurper.sheetNames == ['breaking', 'bad']
    }

    @Test
    void "list values of a row"() {
        assertEquals(['aaron', 'jesse'], xlsWorkbookSlurper[0][2].toList())
    }

    @Test
    void "list rows values of a sheet"() {
        assertEquals([["season 1"], ["Pilot"], ["The Cat's in the Bag"]], xlsWorkbookSlurper.bad.toList())
        assertEquals([["Pilot"], ["The Cat's in the Bag"]], xlsWorkbookSlurper.bad.toList(labels: true))
        assertEquals([["bryan", "walter"], ["aaron", "jesse"], ["anna", "skyler"]], xlsWorkbookSlurper.breaking.toList(labels: true))
    }

    @Test
    void "list rows values of a workbook"() {
        assertEquals([
                [["actor", "character"], ["bryan", "walter"], ["aaron", "jesse"], ["anna", "skyler"]],
                [["season 1"], ["Pilot"], ["The Cat's in the Bag"]]
        ],  xlsWorkbookSlurper.toList())
    }

    @Test
    void findAll() {
        assertEquals([['aaron', 'jesse']], xlsWorkbookSlurper.breaking.findAll { cell(0) == 'aaron' })
        assertEquals([['aaron', 'jesse'], ["anna", "skyler"]], xlsWorkbookSlurper.findAll(labels:true) { cell(0).startsWith('a') })
    }

    @Test
    void find() {
        assertEquals(['aaron', 'jesse'], xlsWorkbookSlurper.breaking.find(offset:1) { cell(0).startsWith('a') })
        assertEquals(['aaron', 'jesse'], xlsWorkbookSlurper.find(offset:1) { cell(0).startsWith('a') })
    }

    @Test
    void findRowSlurper() {
        assertEquals(3, xlsWorkbookSlurper.breaking.findRowSlurper { cell(1) == "skyler" }.rowIndex)
        assertEquals(3, xlsWorkbookSlurper.findRowSlurper { cell(1) == "skyler" }.rowIndex)
    }

    @Test
    void iterateOverRows() {
        // by default, eachRow iterates over the first sheet's rows
        // if labels is set to true, the first line of the sheet is ignored
        xlsWorkbookSlurper.eachRow(max: 2, labels: true) {
            switch (actor) {
                case 'bryan':
                    assert character == "walter"
                    break
                case 'aaron':
                    assert character == "jesse"
            }
        }

        // but you may specify another sheet by its name or index (always 0-based)
        // you may set an offset almost everywhere
        xlsWorkbookSlurper.eachRow(sheet: 1, max: 1, labels: true, offset: 1) {
            assert cell(0) == "The Cat's in the Bag"
        }
    }

    @Test
    void iterateOverSheets() {
        // eachSheet and eachRow can be nested
        // sheetIndex, sheetName and rowIndex are available in case you need it
        xlsWorkbookSlurper.eachSheet {
            assert sheetIndex in [0, 1]
            eachRow([offset: 2]) {
                if (sheetName == 'bad') {
                    assert cell(0) == 'The Cat\'s in the Bag'
                    assert sheetIndex == 1
                    assert rowIndex == 2
                } else {
                    assert sheetName == 'breaking'
                }
            }
        }
    }
}
