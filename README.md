gexcelslurper
=============

Groovy Excel Slurper, easily reads Excel files thanks to Groovy and Apache POI.

What it can do for you in a groovy way:
* read xls and xlsx files
* direct access to cells values
* iterate over sheets and rows
* access sheets and rows by name
* get all values from a row, a sheet or even a whole workbook with toList() methods
* offset, max, labels options
* findAll methods on workbooks and sheets

See [tests](https://github.com/florent-blanvillain/gexcelslurper/blob/master/test/org/gexcelslurper/ExcelSlurperTest.groovy) for code and explanations.

This small lib is based on the source code found on this [blog post](http://www.technipelago.se/content/technipelago/blog/44) but goes beyond.
It can be used as a drop-in replacement for the latter, except that you have to use `XlsWorkbookSlurper` class instead of `ExcelBuilder` and `eachRow` method instead of `eachLine`.

## Usage

Build with gradle and gather ./build/gexcelslurper-x.jar or just integrate `XlsWorkbookSlurper`, `SheetSlurper` and `RowSlurper` classes in your project.

It requires Apache POI from version 3.5.

Needed dependencies for Apache POI 3.5 to 3.8:
* `org.apache.poi:poi:3.x`
* `org.apache.poi:poi-ooxml:3.x`

For Apache POI 3.9 the following is also needed:
* `org.apache.poi:poi-ooxml-schemas:3.9`

## Implementation

Three wrappers arround org.apache.poi.ss.usermodel dot `Workbook`, `Sheet` and `Row` : `XlsWorkbookSlurper`, `SheetSlurper` and `RowSlurper`.
One can easily access wrapped objects by accessing `workbook`, `sheet` and `row` properties.

## WTFPL

Following a huge amount of messages asking for a licence (and by huge I mean zero), Here it is:

<quote>Copyright © 2013 Florent Blanvillain florent.blanvillain@gmail.com
This work is free. You can redistribute it and/or modify it under the
terms of the Do What The Fuck You Want To Public License, Version 2,
as published by Sam Hocevar. See the COPYING file for more details.</quote>