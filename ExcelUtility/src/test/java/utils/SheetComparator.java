package utils;

import java.io.FileWriter;
import java.util.HashSet;
import java.util.Set;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SheetComparator {
    public static void compareSheets(XSSFWorkbook wb1, XSSFWorkbook wb2, String file1, String file2, FileWriter myWriter) throws Exception {
        myWriter.write("================================================ Sheet Comparison =================================================='\n");
        myWriter.write("Comparing sheets between '" + file1 + "' and '" + file2 + "'.\n");

        Set<String> wb1Sheets = new HashSet<>();
        Set<String> wb2Sheets = new HashSet<>();
        for (int i = 0; i < wb1.getNumberOfSheets(); i++) wb1Sheets.add(wb1.getSheetName(i));
        for (int i = 0; i < wb2.getNumberOfSheets(); i++) wb2Sheets.add(wb2.getSheetName(i));

        myWriter.write(file1 + " sheet count: " + wb1Sheets.size() + "\n");
        myWriter.write(file2 + " sheet count: " + wb2Sheets.size() + "\n");

        Set<String> extraInWb1 = new HashSet<>(wb1Sheets);
        extraInWb1.removeAll(wb2Sheets);

        Set<String> extraInWb2 = new HashSet<>(wb2Sheets);
        extraInWb2.removeAll(wb1Sheets);

        // Only print extra sheets in a user-friendly, one-per-line format, no set output
        if (!extraInWb1.isEmpty() || !extraInWb2.isEmpty()) {
            myWriter.write("================================================Extra sheet Names comparation==============================================\n");
            if (!extraInWb1.isEmpty()) {
                myWriter.write(file1 + ": " + extraInWb1 + "\n");
            }
            if (!extraInWb2.isEmpty()) {
                myWriter.write(file2 + " : " + extraInWb2 + "\n");
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
            myWriter.write("All sheet names match in both workbooks.\n");
        }
    }
}