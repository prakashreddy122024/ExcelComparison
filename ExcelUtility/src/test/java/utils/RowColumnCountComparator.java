package utils;

import java.io.FileWriter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class RowColumnCountComparator {
    public static void compareRowAndColumnCount(XSSFSheet sheet1, XSSFSheet sheet2,String file1, String file2, FileWriter myWriter) throws Exception {
        // Row count
        int rowCount1 = sheet1.getPhysicalNumberOfRows();
        int rowCount2 = sheet2.getPhysicalNumberOfRows();
        if (rowCount1 != rowCount2) {
            myWriter.write("================================================Row Count Mismatch========================================\n");
            myWriter.write("Row count mismatch in sheet '" + sheet1.getSheetName() + "':\n");
            myWriter.write(file1 +" :"+ rowCount1 + " rows\n");
            myWriter.write(file2 +" :"+ rowCount2 + " rows\n\n");
        }
        // Column count (header row)
        Row header1 = sheet1.getRow(0);
        Row header2 = sheet2.getRow(0);
        int colCount1 = (header1 != null) ? header1.getPhysicalNumberOfCells() : 0;
        int colCount2 = (header2 != null) ? header2.getPhysicalNumberOfCells() : 0;
        if (colCount1 != colCount2) {
            myWriter.write("================================================Column Count Mismatch======================================\n");
            myWriter.write("Column count mismatch in sheet '" + sheet1.getSheetName() + "':\n");
            myWriter.write(file1 +" :"+ colCount1 + " columns\n");
            myWriter.write(file2 +" :"+ colCount2 + " columns\n\n");
        }
    }
}
