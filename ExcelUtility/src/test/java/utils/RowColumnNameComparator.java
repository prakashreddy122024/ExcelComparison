package utils;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class RowColumnNameComparator {
    public static void compareRowColumnNames(XSSFSheet sheet1, XSSFSheet sheet2, String file1, String file2, HtmlLogger myWriter) throws Exception {
        int maxRows = Math.max(sheet1.getLastRowNum(), sheet2.getLastRowNum());

        for (int r = 0; r <= maxRows; r++) {
            if (r == 0) {
                myWriter.write("<h2>Row and Column Name Comparison</h2>\n");
            }
            Row row1 = sheet1.getRow(r);
            Row row2 = sheet2.getRow(r);
            if (row1 == null && row2 == null) continue;
            int colCount1 = (row1 != null) ? row1.getLastCellNum() : 0;
            int colCount2 = (row2 != null) ? row2.getLastCellNum() : 0;
            int maxCols = Math.max(colCount1, colCount2);
            for (int c = 0; c < maxCols; c++) {
                String colName1 = getCellValue(row1, c);
                String colName2 = getCellValue(row2, c);
                if (colName1.isEmpty() && !colName2.isEmpty()) {
                    myWriter.write("<p style='color:orange;'>Extra column in <b>" + file2 + "</b> at row <b>" + (r + 1) + "</b>, column <b>" + (c + 1) + "</b> in sheet '<b>" + sheet2.getSheetName() + "</b>': '<b>" + colName2 + "</b>'</p>\n");
                } else if (!colName1.isEmpty() && colName2.isEmpty()) {
                    myWriter.write("<p style='color:orange;'>Extra column in <b>" + file1 + "</b> at row <b>" + (r + 1) + "</b>, column <b>" + (c + 1) + "</b> in sheet '<b>" + sheet1.getSheetName() + "</b>': '<b>" + colName1 + "</b>'</p>\n");
                }
            }
        }

    }

    private static String getCellValue(Row row, int colIdx) {
        if (row == null) return "";
        Cell cell = row.getCell(colIdx);
        if (cell == null) return "";
        return cell.toString().trim();
    }
}
