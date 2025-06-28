package utils;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtils {
    public XSSFWorkbook workbook;
    public ExcelUtils(String filePath) {
        try {
            workbook = new XSSFWorkbook(filePath);
        } catch(Exception exp) {
            exp.printStackTrace();
        }
    }

    public Object getCellData(XSSFSheet sheet, int row, int column) {
        DataFormatter formatter = new DataFormatter();
        Object value = formatter.formatCellValue(sheet.getRow(row).getCell(column));
        return value;
    }

}
