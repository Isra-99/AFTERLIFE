package afterLife.actions;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class ReadFromExcel {

    public ReadFromExcel(WebDriver d){}
    public static String searchContent() throws IOException {
        //new File
        File source  =  new File("src/main/resources/QA.xlsx");
        // In order to read the above file we use fileInputstream class and create object of it
        FileInputStream input = new FileInputStream(source);
        //calling for particular work book
        XSSFWorkbook wb = new XSSFWorkbook(input);
        /// since data is written in first sheet so sheet number is one
        XSSFSheet sheet =        wb.getSheetAt(0);
        String text= sheet.getRow(1).getCell(1).getStringCellValue();
        XSSFCell searchLink = sheet.getRow(2).getCell(1);
        return  text;

    }
    public static String url() throws IOException {
        //new File
        File source  =  new File("src/main/resources/QA.xlsx");
        // In order to read the above file we use fileInputstream class and create object of it
        FileInputStream input = new FileInputStream(source);
        //calling for particular work book
        XSSFWorkbook wb = new XSSFWorkbook(input);
        /// since data is written in first sheet so sheet number is one
        XSSFSheet sheet =        wb.getSheetAt(0);
        XSSFHyperlink URl= sheet.getRow(0).getCell(1).getHyperlink();
        String url = URl.getAddress();
        String text= sheet.getRow(1).getCell(1).getStringCellValue();
        System.out.println(url);
        return  url;

    }
    public static XSSFCell searchLink() throws IOException {
        //new File
        File source  =  new File("src/main/resources/QA.xlsx");
        // In order to read the above file we use fileInputstream class and create object of it
        FileInputStream input = new FileInputStream(source);
        //calling for particular work book
        XSSFWorkbook wb = new XSSFWorkbook(input);
        /// since data is written in first sheet so sheet number is one
        XSSFSheet sheet =        wb.getSheetAt(0);
        XSSFCell text = sheet.getRow(2).getCell(1);
        System.out.println(text);
        return  text;

    }
}
