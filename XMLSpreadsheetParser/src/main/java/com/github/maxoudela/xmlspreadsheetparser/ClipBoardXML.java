/*The MIT License (MIT)

 Copyright (c) 2015 Samir Hadzic

 Permission is hereby granted, free of charge, to any person obtaining a copy
 of this software and associated documentation files (the "Software"), to deal
 in the Software without restriction, including without limitation the rights
 to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 copies of the Software, and to permit persons to whom the Software is
 furnished to do so, subject to the following conditions:

 The above copyright notice and this permission notice shall be included in all
 copies or substantial portions of the Software.

 THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 SOFTWARE.
 */
package com.github.maxoudela.xmlspreadsheetparser;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

/**
 * This class is abstract, all you need to do is to override eraseValue and
 * handleValue.
 *
 * When we are dealing with Excel content, here is what we do:
 *
 * We browse each row and each cell. When we found a cell, we call handleValue
 * with the indexes set relatively to what you give in the constructor. The
 * value given can be a String or a Date.
 *
 * If the cell has some span, we will not call handleValue repeatedly.
 *
 * If we come across some blank cell in Excel selection, we will call on each
 * blank position the eraseValue method. When a cell is spanning, we do not call
 * earaseValue the spanned area.
 *
 *
 * Documentation is taken from :
 * https://msdn.microsoft.com/en-us/library/aa140066%28v=office.10%29.aspx#odc_xmlss_ss:cell
 */
/**
 *
 * @author Samir Hadzic
 */
public abstract class ClipBoardXML {

    private static final String EXCEL_DATE_FORMAT = "yyyy-MM-dd'T'HH:mm:ss.SSS";
    private static final String ROW_TAG = "Row";
    private static final String CELL_TAG = "Cell";
    private static final String DATA_TAG = "Data";
    private static final String TYPE_TAG = "ss:Type";
    private static final String DATE_TYPE_TAG = "DateTime";
    private static final String INDEX_TAG = "ss:Index";
    private static final String ROW_COUNT_TAG = "ss:ExpandedRowCount";
    private static final String COLUMN_COUNT_TAG = "ss:ExpandedColumnCount";
    private static final String TABLE_TAG = "Table";
    private static final String COLUMN_SPAN_TAG = "ss:MergeAcross";
    private static final String ROW_SPAN_TAG = "ss:MergeDown";

    private final Document doc;
    private int currentRow;
    private final int baseColumn;
    private final int baseRow;
    private int currentColumn;
    private int oldRow = 0;
    private int newRow = 0;

    private int oldCol = 0;
    private int newCol = 0;

    private int selectionRowCount = 0;
    private int selectionColumnCount = 0;

    private int gridRowCount = 0;
    private int gridColCount = 0;
    /**
     * We will flag to true each cell we have considered. All others will be
     * deleted.
     */
    private boolean[][] cellsUsed;

    /**
     * Construct your object by giving the {@link Document} that parsed the XML
     * Spreadsheet.
     *
     * Since you want to paste the Excel selection into your Grid. Gives the
     * starting position on your grid to begin the paste operation.
     *
     * Also give the edge of your grid in order not to override the bounds.
     *
     * @param doc
     * @param currentRow
     * @param currentColumn
     * @param gridRowCount
     * @param gridColCount
     * @throws ParseException
     */
    public ClipBoardXML(Document doc, int currentRow, int currentColumn, int gridRowCount, int gridColCount) throws ParseException {
        this.doc = doc;
        this.currentRow = currentRow;
        this.baseRow = currentRow;
        this.baseColumn = currentColumn;
        this.currentColumn = currentColumn;
        this.gridRowCount = gridRowCount;
        this.gridColCount = gridColCount;
    }

    /**
     * Parse the document. This method will call {@link #eraseValue(int, int) }
     * and {@link #handleValue(int, int, java.lang.Object) }.
     *
     * @throws ParseException
     */
    public void parse() throws ParseException {
        retrieveTableInfo();

        NodeList listRows = doc.getElementsByTagName(ROW_TAG);
        for (int temp = 0; temp < listRows.getLength(); temp++) {
            if (currentRow >= gridRowCount) {
                break;
            }
            Node row = listRows.item(temp);
            if (row.getNodeType() == Node.ELEMENT_NODE) {
                Element rowEl = (Element) row;
                handleRow(rowEl);
            }
        }
        erase();
    }

    protected abstract void eraseValue(int row, int column);

    protected abstract void handleValue(int row, int column, Object value);

    private void erase() {
        if (cellsUsed == null) {
            return;
        }
        int row = 0;
        for (boolean[] bitSet : cellsUsed) {
            int column = 0;
            for (boolean value : bitSet) {
                if (!value && row + baseRow < gridRowCount && column + baseColumn < gridColCount) {
                    eraseValue(row + baseRow, column + baseColumn);
                }
                ++column;
            }
            ++row;
        }
    }

    private void set(int row, int column) {
        if (cellsUsed == null) {
            return;
        }
        if (cellsUsed[row] == null) {
            cellsUsed[row] = new boolean[selectionColumnCount];
        }
        cellsUsed[row][column] = true;
    }

    private void handleRow(Element rowEl) throws ParseException {
        newRow = getIndex(rowEl, newRow);

        //We adjust the currentRow by the difference between the old and the new row.
        currentRow += newRow - oldRow - 1;
        rowEl.hasAttribute(INDEX_TAG);
        NodeList cells = rowEl.getElementsByTagName(CELL_TAG);
        handleColumns(cells);

        currentColumn = baseColumn;
        currentRow++;
        oldRow = newRow;
        //Reset columns
        newCol = 0;
        oldCol = 0;
    }

    private void handleColumns(NodeList cells) throws ParseException {
        for (int column = 0; column < cells.getLength(); column++) {
            Node cell = cells.item(column);
            if (cell.getNodeType() == Node.ELEMENT_NODE) {
                Element cellEl = (Element) cell;
                newCol = getIndex(cellEl, newCol);

                currentColumn += newCol - oldCol - 1;
                if (currentColumn >= gridColCount) {
                    break;
                }

                handleCell(currentRow, currentColumn, cellEl);
                set(newRow - 1, newCol - 1);
                currentColumn++;
            }

            oldCol = newCol;
        }
    }

    private void handleSpan(Element cellEl) {
        int columnSpan = 0;
        int rowSpan = 0;
        if (cellEl.hasAttribute(COLUMN_SPAN_TAG)) {
            columnSpan = Integer.valueOf(cellEl.getAttribute(COLUMN_SPAN_TAG));
            for (int i = 0; i < columnSpan; ++i) {
                set(newRow - 1, newCol + i);
            }

        }

        if (cellEl.hasAttribute(ROW_SPAN_TAG)) {
            rowSpan = Integer.valueOf(cellEl.getAttribute(ROW_SPAN_TAG));
            for (int i = 0; i < rowSpan; ++i) {
                set(newRow + i, newCol - 1);
            }
        }
        currentColumn += columnSpan;
        newCol += columnSpan;
    }

    private void retrieveTableInfo() {
        NodeList tableList = doc.getElementsByTagName(TABLE_TAG);

        //Retrieve possible row and column count of Excel selection.
        if (tableList.getLength() == 1 && tableList.item(0).getNodeType() == Node.ELEMENT_NODE) {
            Element tableEl = (Element) tableList.item(0);
            if (tableEl.hasAttribute(ROW_COUNT_TAG)) {
                selectionRowCount = Integer.valueOf(tableEl.getAttribute(ROW_COUNT_TAG));
            }
            if (tableEl.hasAttribute(COLUMN_COUNT_TAG)) {
                selectionColumnCount = Integer.valueOf(tableEl.getAttribute(COLUMN_COUNT_TAG));
            }
            //Initialize array if both values are respected.
            if (selectionRowCount != 0 && selectionColumnCount != 0) {
                cellsUsed = new boolean[selectionRowCount][selectionColumnCount];
            }
        }
    }

    /**
     * If the Element has not index specified, it means it's juste one after the
     * previous. If it has an index, it's the absolute index relative to the
     * first index of the first row.
     *
     * @param rowEl
     * @param index
     * @return
     */
    private int getIndex(Element rowEl, int index) {
        return rowEl.hasAttribute(INDEX_TAG) ? Integer.valueOf(rowEl.getAttribute(INDEX_TAG)) : index + 1;
    }

    private void handleCell(int currentRow, int currentColumn, Element cellEl) throws ParseException {
        NodeList datas = cellEl.getElementsByTagName(DATA_TAG);
        if (datas.getLength() == 1 && datas.item(0).getNodeType() == Node.ELEMENT_NODE) {
            Element data = (Element) datas.item(0);
            String cellType = data.getAttribute(TYPE_TAG);
            Object object;
            switch (cellType) {
                case DATE_TYPE_TAG:
                    //If the SpreadsheetCell is a date, we can give the double value. Otherwise, we just give the date.
                    Date date = new SimpleDateFormat(EXCEL_DATE_FORMAT).parse(data.getTextContent());
                    object = Long.valueOf(date.getTime()).doubleValue();
                    break;
                default:
                    object = data.getTextContent();
                    break;
            }
            if (object != null) {
                handleValue(currentRow, currentColumn, object);
            }
        }
    }
}
