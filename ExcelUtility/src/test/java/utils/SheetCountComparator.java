package utils;

import java.io.FileWriter;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SheetCountComparator {
    public static void compareSheetCount(XSSFWorkbook wb1, XSSFWorkbook wb2, String file1, String file2, FileWriter myWriter) throws Exception {
        myWriter.write("======================================================================Sheet Count Validation====================================\n");
        int count1 = wb1.getNumberOfSheets();
        int count2 = wb2.getNumberOfSheets();
        if (count1 != count2) {
            myWriter.write("Sheet count mismatch: " + file1 + " has " + count1 + " sheets, " + file2 + " has " + count2 + " sheets.\n");
            if (count1 > count2) {
                myWriter.write(file1 + " has " + (count1 - count2) + " extra sheet(s).\n");
            } else {
                myWriter.write(file2 + " has " + (count2 - count1) + " extra sheet(s).\n");
            }
        } else {
            myWriter.write("Both workbooks have the same number of sheets: " + count1 + "\n");
        }
        myWriter.write("//==========================\n");
    }
}

