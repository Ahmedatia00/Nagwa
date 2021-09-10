import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelOperations {

    public static String importDataFromExcel() throws IOException {
        FileInputStream fis = new FileInputStream("F:\\Test data.xls"); // insert the path of the sheet at your PC
        Workbook wb = WorkbookFactory.create(fis);
        Sheet sh = wb.getSheet("Sheet1");
        Row row = sh.getRow(0);
        String data = row.getCell(1).getStringCellValue();
        return data;
    }

    public static void writeToExcel(int cellValue) throws IOException {
        FileInputStream fis = new FileInputStream("F:\\Test data.xls");
        Workbook wb = WorkbookFactory.create(fis);
        Sheet sh = wb.getSheet("Sheet1");
        Row row2 = sh.getRow(1);
        Cell cel = row2.createCell(1);
        cel.setCellValue(cellValue);
        FileOutputStream fos = new FileOutputStream("F:\\Test data.xls");
        wb.write(fos);
    }
}
