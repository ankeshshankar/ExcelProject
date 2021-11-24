package Application;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

public class Reading {

    public static void main(String[] args) {
        String excelFilePath = ".//.//DataFiles/Churn.xlsx";
        try {
            FileInputStream inputStream = new FileInputStream(excelFilePath);
            XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
            XSSFSheet sheet = workbook.getSheet("DataSet");
            // Using For Loop
/*            int rows = sheet.getLastRowNum();
            int cols = sheet.getRow(1).getLastCellNum();
            for (int i = 0; i < rows; i++) {
                XSSFRow row = sheet.getRow(i);
                for (int j = 0; j < cols; j++) {
                    XSSFCell cell = row.getCell(j);
                    switch (cell.getCellType()) {
                        case STRING:
                            System.out.print(cell.getStringCellValue()+"  ");
                            break;
                        case NUMERIC:
                            System.out.print(cell.getNumericCellValue()+"  ");
                            break;
                    }
                }
                System.out.println();
            }*/

            // Using Iterator

            Iterator<Row> iterator = sheet.iterator();
            while (iterator.hasNext()){
                XSSFRow row = (XSSFRow) iterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()){
                    XSSFCell cell = (XSSFCell) cellIterator.next();
                    switch (cell.getCellType()) {
                        case STRING:
                            System.out.print(cell.getStringCellValue()+" | ");
                            break;
                        case NUMERIC:
                            System.out.print(cell.getNumericCellValue()+" | ");
                            break;
                    }
                }
                System.out.println();
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
