package utils;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class RowColumnCountComparator {
    public static void compareRowAndColumnCount(XSSFSheet sheet1, XSSFSheet sheet2, String file1, String file2, HtmlLogger myWriter) throws Exception {
        // Row count
        int rowCount1 = sheet1.getPhysicalNumberOfRows();
        int rowCount2 = sheet2.getPhysicalNumberOfRows();
        if (rowCount1 != rowCount2) {
            myWriter.write("<h2 style='color:red;'>Row Count Mismatch</h2>\n");
            myWriter.write("<p>Row count mismatch in sheet '<b>" + sheet1.getSheetName() + "</b>':</p>\n");
            myWriter.write("<ul><li>" + file1 + ": <b>" + rowCount1 + "</b> rows</li>\n");
            myWriter.write("<li>" + file2 + ": <b>" + rowCount2 + "</b> rows</li></ul>\n");
        }
        // Column count (header row)
        Row header1 = sheet1.getRow(0);
        Row header2 = sheet2.getRow(0);
        int colCount1 = (header1 != null) ? header1.getPhysicalNumberOfCells() : 0;
        int colCount2 = (header2 != null) ? header2.getPhysicalNumberOfCells() : 0;
        if (colCount1 != colCount2) {
            myWriter.write("<h2 style='color:red;'>Column Count Mismatch</h2>\n");
            myWriter.write("<p>Column count mismatch in sheet '<b>" + sheet1.getSheetName() + "</b>':</p>\n");
            myWriter.write("<ul><li>" + file1 + ": <b>" + colCount1 + "</b> columns</li>\n");
            myWriter.write("<li>" + file2 + ": <b>" + colCount2 + "</b> columns</li></ul>\n");
        }
    }
}
