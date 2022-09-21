package Domaci_22;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class ApachePoiUtil {
    public static void readExcel(String path){
        try {
            FileInputStream inputStream = new FileInputStream(new File("domaci22.xlsx"));

            XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
            XSSFSheet sheet = workbook.getSheet("Sheet1");

            for(int j = 0; j<3; j++) {
                XSSFRow row = sheet.getRow(j);

                for (int i = 0; i < 3; i++) {
                    XSSFCell cell = row.getCell(i);
                    String celija = cell.getStringCellValue();
                    System.out.print(celija + " ");
                }
                System.out.println();
            }
        }catch (FileNotFoundException ex){
            System.out.println("FileNotFound.class");
        } catch (IOException e) {
            e.printStackTrace();
        }catch (NullPointerException e) {
            //e.printStackTrace();
        }
    }
}