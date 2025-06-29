package utils;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class HeaderComparator {
    public static void compareHeaderCountAndNames(XSSFSheet sheet1, XSSFSheet sheet2, String file1, String file2, HtmlLogger myWriter) throws Exception {
        Row header1 = sheet1.getRow(0);
        Row header2 = sheet2.getRow(0);
        int count1 = (header1 != null) ? header1.getPhysicalNumberOfCells() : 0;
        int count2 = (header2 != null) ? header2.getPhysicalNumberOfCells() : 0;

        int maxCols = Math.max(count1, count2);
        boolean anyNameMismatch = false;
        StringBuilder nameMismatchBuilder = new StringBuilder();
        // Print extra columns in BEFORE
        if (count1 > count2) {
            for (int c = count2; c < count1; c++) {
                String name1 = getCellValue(header1, c);
                if (!name1.isEmpty()) {
                    nameMismatchBuilder.append("Extra column in "+file1+" at column ").append(c + 1)
                        .append(" in sheet '").append(sheet1.getSheetName())
                        .append("': '").append(name1).append("'\n");
                }
            }
        }
        // Print extra columns in AFTER
        if (count2 > count1) {
            for (int c = count1; c < count2; c++) {
                String name2 = getCellValue(header2, c);
                if (!name2.isEmpty()) {
                    nameMismatchBuilder.append("Extra column in "+file2+" at column ").append(c + 1)
                        .append(" in sheet '").append(sheet2.getSheetName())
                        .append("': '").append(name2).append("'\n");
                }
            }
        }
        // Print header name mismatches
        for (int c = 0; c < Math.min(count1, count2); c++) {
            String name1 = getCellValue(header1, c);
            String name2 = getCellValue(header2, c);
            if (!name1.equals(name2)) {
                anyNameMismatch = true;
                nameMismatchBuilder.append("Header name mismatch at column ").append(c + 1)
                    .append(" in sheet '").append(sheet1.getSheetName())
                    .append("': '").append(name1).append("' vs '").append(name2).append("'\n");
            }
        }
        if (nameMismatchBuilder.length() > 0) {
            myWriter.write("<h2 style='color:red;'>Header Name Mismatch</h2>\n");
            myWriter.write("<pre>" + nameMismatchBuilder.toString() + "</pre>\n");
        }
    }

    private static String getCellValue(Row row, int colIdx) {
        if (row == null) return "";
        Cell cell = row.getCell(colIdx);
        if (cell == null) return "";
        return cell.toString().trim();
    }
}
