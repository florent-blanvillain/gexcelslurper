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
        assertEquals xlsWorkbookSlurper[0][2][1], 'jesse'

        assertEquals "'breaking' is the name of the first sheet and 'character' is the first cell value of the second row",
                xlsWorkbookSlurper.breaking[2].character, 'jesse'
    }

    @Test
    void sheetNames() {
        assert xlsWorkbookSlurper.sheetNames == ['breaking', 'bad']
    }

    @Test
    void iterateOverRows() {
        xlsWorkbookSlurper.eachRow([max: 2, labels: true]) {
            assert rowIndex in [1, 2]
            switch (actor) {
                case 'bryan':
                    assert character == "walter"
                    break
                case 'aaron':
                    assert character == "jesse"
            }
        }
    }

    @Test
    void iterateOverSheets() {
        xlsWorkbookSlurper.eachSheet {
            assert sheetIndex in [0, 1]
            eachRow([offset: 2]) {
                if (sheetName == 'bad') {
                    assert cell(0) == 'The Cat\'s in the Bag'
                    assert sheetIndex == 1
                }
            }
        }
    }
}
