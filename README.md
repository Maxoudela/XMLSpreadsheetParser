XMLSpreadsheetParser
======

This project allow you to parse what Excel is putting inside a ClipBoard when you copy a selection. (XML Spreadsheet format)

You will be able to decode what Excel has put in it, and paste into your grid or application.

1) Extract the content from Clipboard or elsewhere, and materialize  a [Document](https://docs.oracle.com/javase/8/docs/api/org/w3c/dom/Document.html).
An example is given in the ClipBoardUtil class.

2) Instantiate a ClipBoardXML class and give it the [Document](https://docs.oracle.com/javase/8/docs/api/org/w3c/dom/Document.html). You need to give the position in your Grid object where you want to paste the value, and the grid bounds.
Also overrides two methods in order to deal with the values extracted from Excel.

3) Call parse() method :)

