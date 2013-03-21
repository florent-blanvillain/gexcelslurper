package org.gexcelslurper

import org.junit.Test

public class ExcelSlurperTest {
    static final String XLSX = 'gexcelslurper.xlsx'
    static final URL XLSX_URL = ClassLoader.systemClassLoader.getResource(XLSX)
    static final ExcelSlurper slurper = new ExcelSlurper(XLSX_URL)

    @Test
    void load() {
        new ExcelSlurper(new File(this.class.classLoader.getResource(XLSX).file)).eachLine([max: 1]) {
            assert rowIndex == 0
        }

        new ExcelSlurper(XLSX_URL).eachLine([max: 1]) {
            assert rowIndex == 0
        }
    }

    @Test
    void iterateOverLines() {
        slurper.eachLine([max: 2, labels: true]) {
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
        slurper.eachSheet {
            assert sheetIndex in [0, 1]
            eachLine([offset: 2]) {
                if (sheetName == 'bad') {
                    assert cell(0) == 'The Cat\'s in the Bag'
                    assert sheetIndex == 1
                }
            }
        }
    }

    @Test
    void sheetNames() {
        assert slurper.sheetNames == ['breaking', 'bad']
    }
}
