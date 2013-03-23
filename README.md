gexcelslurper
=============

Groovy Excel Slurper, reads excel with groovy and POI

It requires Apache POI from version 3.5.

Needed dependencies for Apache POI 3.5 to 3.8:
* 'org.apache.poi:poi:3.x'
* 'org.apache.poi:poi-ooxml:3.x'

For Apache POI 3.9 the following is also needed:
* 'org.apache.poi:poi-ooxml-schemas:3.9'

This small lib is based on the source code found on this [blog post](http://www.technipelago.se/content/technipelago/blog/44) but goes beyond.
It can be used as a drop-in replacement for the latter, except that you have `XlsWorkbookSlurper` class instead of `ExcelBuilder` and `eachRow` method instead of `eachLine`.

See [tests](https://github.com/florent-blanvillain/gexcelslurper/blob/master/test/org/gexcelslurper/ExcelSlurperTest.groovy) for more code examples.

