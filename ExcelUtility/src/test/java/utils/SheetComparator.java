package utils;

import java.util.HashSet;
import java.util.Set;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SheetComparator {
    public static void compareSheets(XSSFWorkbook wb1, XSSFWorkbook wb2, String file1, String file2, HtmlLogger myWriter) throws Exception {
        myWriter.write("<h2>Sheet Comparison</h2>\n");
        myWriter.write("<p>Comparing sheets between '<b>" + file1 + "</b>' and '<b>" + file2 + "</b>'.</p>\n");

        Set<String> wb1Sheets = new HashSet<>();
        Set<String> wb2Sheets = new HashSet<>();
        for (int i = 0; i < wb1.getNumberOfSheets(); i++) wb1Sheets.add(wb1.getSheetName(i));
        for (int i = 0; i < wb2.getNumberOfSheets(); i++) wb2Sheets.add(wb2.getSheetName(i));

        myWriter.write("<ul><li>" + file1 + " sheet count: <b>" + wb1Sheets.size() + "</b></li>\n");
        myWriter.write("<li>" + file2 + " sheet count: <b>" + wb2Sheets.size() + "</b></li></ul>\n");

        Set<String> extraInWb1 = new HashSet<>(wb1Sheets);
        extraInWb1.removeAll(wb2Sheets);

        Set<String> extraInWb2 = new HashSet<>(wb2Sheets);
        extraInWb2.removeAll(wb1Sheets);

        // Only print extra sheets in a user-friendly, one-per-line format, no set output
        if (!extraInWb1.isEmpty() || !extraInWb2.isEmpty()) {
            myWriter.write("<h3 style='color:orange;'>Extra Sheet Names Comparison</h3>\n");
            if (!extraInWb1.isEmpty()) {
                myWriter.write("<p>" + file1 + ": <b>" + extraInWb1 + "</b></p>\n");
            }
            if (!extraInWb2.isEmpty()) {
                myWriter.write("<p>" + file2 + " : <b>" + extraInWb2 + "</b></p>\n");
            }
        }

        // Find sheets with the same index but different names
        Set<String> mismatchedNames = new HashSet<>();
        int minSheets = Math.min(wb1.getNumberOfSheets(), wb2.getNumberOfSheets());
        for (int i = 0; i < minSheets; i++) {
            String name1 = wb1.getSheetName(i);
            String name2 = wb2.getSheetName(i);
            if (!name1.equals(name2)) {
                mismatchedNames.add(name1 + " (in " + file1 + ") vs " + name2 + " (in " + file2 + ")");
            }
        }

        if (extraInWb1.isEmpty() && extraInWb2.isEmpty() && mismatchedNames.isEmpty()) {
            myWriter.write("<p style='color:green;'>All sheet names match in both workbooks.</p>\n");
        }
    }
}