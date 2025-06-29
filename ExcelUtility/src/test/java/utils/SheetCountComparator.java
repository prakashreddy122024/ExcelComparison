package utils;

import java.io.FileWriter;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SheetCountComparator {
    public static void compareSheetCount(XSSFWorkbook wb1, XSSFWorkbook wb2, String file1, String file2, FileWriter myWriter) throws Exception {
        myWriter.write("<h2>Sheet Count Validation</h2>\n");
        int count1 = wb1.getNumberOfSheets();
        int count2 = wb2.getNumberOfSheets();
        if (count1 != count2) {
            myWriter.write("<p style='color:red;'>Sheet count mismatch: <b>" + file1 + "</b> has <b>" + count1 + "</b> sheets, <b>" + file2 + "</b> has <b>" + count2 + "</b> sheets.</p>\n");
            if (count1 > count2) {
                myWriter.write("<ul><li>" + file1 + " has <b>" + (count1 - count2) + "</b> extra sheet(s).</li></ul>\n");
            } else {
                myWriter.write("<ul><li>" + file2 + " has <b>" + (count2 - count1) + "</b> extra sheet(s).</li></ul>\n");
            }
        } else {
            myWriter.write("<p>Both workbooks have the same number of sheets: <b>" + count1 + "</b></p>\n");
        }
        myWriter.write("<hr>\n");
    }
}
