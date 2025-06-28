package utils;

import java.io.FileWriter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelCompare {
    static String file1 = "BEFORE DATA 5.0.xlsx";
    static String file2 = "AFTER DATA 5.0.xlsx";
    static String logFile = "ExcelUtility/data/Log.txt";

    public static void main(String[] args) throws Exception {
        java.nio.file.Files.createDirectories(java.nio.file.Paths.get("data"));
        try (FileWriter myWriter = new FileWriter(logFile, false)) {
            ExcelUtils excel1 = new ExcelUtils("ExcelUtility/data/" + file1);
            ExcelUtils excel2 = new ExcelUtils("ExcelUtility/data/" + file2);

            // Write system date and time in the log
            java.time.LocalDateTime now = java.time.LocalDateTime.now();
            java.time.format.DateTimeFormatter formatter = java.time.format.DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
            myWriter.write("                                          " + now.format(formatter) + "\n");
            myWriter.write("================================================================================================================================\n");
            // 1. Compare Workbook Sheets
            SheetComparator.compareSheets(excel1.workbook, excel2.workbook, file1, file2, myWriter);
            compareExcel(excel1.workbook, excel2.workbook, myWriter);
        }
    }


    public static void compareExcel(XSSFWorkbook beforeWb, XSSFWorkbook afterWb, FileWriter myWriter) throws Exception {
        int beforeSheets = beforeWb.getNumberOfSheets();
        for (int i = 0; i < beforeSheets; i++) {
            String sheetName = beforeWb.getSheetName(i);
            XSSFSheet beforeSheet = beforeWb.getSheetAt(i);
            XSSFSheet afterSheet = afterWb.getSheet(sheetName);
            if (afterSheet == null) {
//                myWriter.write("Skipping comparison for missing sheet in AFTER: '" + sheetName + "'\n");
                continue;
            }
            compareSheet(beforeSheet, afterSheet, myWriter);
        }
    }

    public static void compareSheet(XSSFSheet sheet1, XSSFSheet sheet2, FileWriter myWriter) throws Exception {
        RowColumnCountComparator.compareRowAndColumnCount(sheet1, sheet2, file1, file2, myWriter);
        HeaderComparator.compareHeaderCountAndNames(sheet1, sheet2, file1, file2, myWriter);
        RowColumnNameComparator.compareRowColumnNames(sheet1, sheet2, file1, file2, myWriter);

        myWriter.write("\n");
        // Cell-level comparison
        int lastRow1 = sheet1.getLastRowNum();
        int lastRow2 = sheet2.getLastRowNum();
        int maxRows = Math.max(lastRow1, lastRow2);
        for (int r = 0; r <= maxRows; r++) {
            Row row1 = sheet1.getRow(r);
            Row row2 = sheet2.getRow(r);
            if (row1 == null && row2 == null) continue;
            int lastCol1 = (row1 != null) ? row1.getLastCellNum() : 0;
            int lastCol2 = (row2 != null) ? row2.getLastCellNum() : 0;
            int maxCols = Math.max(lastCol1, lastCol2);
            for (int c = 0; c < maxCols; c++) {
                Cell cell1 = (row1 != null) ? row1.getCell(c) : null;
                Cell cell2 = (row2 != null) ? row2.getCell(c) : null;
                DataFormatter df = new DataFormatter();
                String value1 = (cell1 == null) ? "" : df.formatCellValue(cell1);
                String value2 = (cell2 == null) ? "" : df.formatCellValue(cell2);

                // If header is mismatched, it will generate all column names
//                if (!value1.equals(value2)) {
//                    myWriter.write("Mismatch at Sheet: " + sheet1.getSheetName() + ", Row: " + (r+1) + ", Col: " + (c+1) + " => '" + value1 + "' vs '" + value2 + "'\n");
//                }
            }
        }
        myWriter.write("\n");
    }
}