package utils;

import java.io.FileWriter;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class RowColumnNameComparator {
    public static void compareRowColumnNames(XSSFSheet sheet1, XSSFSheet sheet2, String file1, String file2, FileWriter myWriter) throws Exception {
        int maxRows = Math.max(sheet1.getLastRowNum(), sheet2.getLastRowNum());

        for (int r = 0; r <= maxRows; r++) {
            if (r == 0) {
                myWriter.write("================================================Row and Column Name Comparison======================================\n");
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
                    myWriter.write("Extra column in " + file2 + " at row " + (r + 1) + ", column " + (c + 1) + " in sheet '" + sheet2.getSheetName() + "': '" + colName2 + "'\n");
                } else if (!colName1.isEmpty() && colName2.isEmpty()) {
                    myWriter.write("Extra column in " + file1 + " at row " + (r + 1) + ", column " + (c + 1) + " in sheet '" + sheet1.getSheetName() + "': '" + colName1 + "'\n");
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
