package Domaci_22;

import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

public class Main {
    public static void main(String[] args) throws IOException {
        ApachePoiUtil.readExcel("domaci22.xlsx");






        try {
            writeExcel("NewDoc.xlsx");
        } catch (FileNotFoundException e) {
            System.out.println("File not found!");
        } catch (IOException e) {
        }

    }

    public static void writeExcel(String fileName) throws FileNotFoundException, IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("NewDoc");

        XSSFRow row00 = sheet.createRow(0);
        XSSFCell cell00 = row00.createCell(0);
        cell00.setCellValue("Sanja");
        XSSFCell cell01 = row00.createCell(1);
        cell01.setCellValue("Stanic");
        XSSFCell cell02 = row00.createCell(2);
        cell02.setCellValue("sanja.stanic@gmail.com");

        XSSFRow row10 = sheet.createRow(1);
        XSSFCell cell10 = row10.createCell(0);
        cell10.setCellValue("Goran");
        XSSFCell cell11 = row10.createCell(1);
        cell11.setCellValue("Stojanac");
        XSSFCell cell12 = row10.createCell(2);
        cell12.setCellValue("goran.stojanac@gmail.com");

        FileOutputStream fileOutputStream = new FileOutputStream(new File(fileName));
        workbook.write(fileOutputStream);
        fileOutputStream.close();

    }


}










