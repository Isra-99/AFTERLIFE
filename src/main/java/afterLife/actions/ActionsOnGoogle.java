package afterLife.actions;

import net.sf.jxls.reader.XLSReader;
import net.sf.jxls.reader.XLSReaderImpl;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;

import java.io.*;
import java.util.ArrayList;

import static afterLife.actions.Locators.*;

public class ActionsOnGoogle {
    /// Constructor of class ActionsOnGoogle
    public ActionsOnGoogle  (WebDriver d){}
    /// Method to enter text in google search bar
    public void enterSomething(String text, WebDriver driver){
        Locators.searchBar(driver).sendKeys(text);
    }
    /// Method to click search icon
    public void clickSearchButton(WebDriver driver){
        Locators.button(driver).submit();
    }
    /// Method to click After Life IMDb link
    public void clickAfterLifeIMDB(WebDriver driver) throws IOException {
        Locators.afterLifeLink(driver).click();
    }
    /// Method to click all cast
    public void clickAllCast(WebDriver driver){
        Locators.allCast(driver).click();
    }
    /// Method to write all cast names, screen names and appearances
    public void writeCastToExcel(WebDriver driver) throws IOException {
        /// Creating new list1, list2 and list 3 for cast names , screen names and appearances
        ArrayList<String> list1 = new ArrayList<>();
        ArrayList<String> list2 = new ArrayList<>();
        ArrayList<String> list3 = new ArrayList<>();
        for (int i = 2; i <= 174; i = i + 2) {

            String str = driver.findElement(By.xpath(name_xpathBefore + i + name_xpathafter)).getText();
            list1.add(str);
            String str1 = driver.findElement(By.xpath(name_xpathBefore1 + i + name_xpathAfter1)).getText();
            list2.add(str1);
            String str2 = driver.findElement(By.xpath(name_xpathBefore2 + i + name_xpathAfter2)).getText();
            list3.add(str2);
        }
        // Creating new file on the basis of path of xlsx file
        File source  =  new File("src/main/resources/QA.xlsx");
        // In order to read the above file we use fileInputstream class and create object of it
        FileInputStream input = new FileInputStream(source);
        //calling for particular work book
        XSSFWorkbook wb = new XSSFWorkbook(input);
        /// since data is written in first sheet so sheet number is one
        XSSFSheet sheet =        wb.getSheetAt(3);
        for (int j =0 ; j<=86 ; j++){
            Row r = sheet.createRow(j+2);
            r.createCell(0).setCellValue(list1.get(j));
            r.createCell(1).setCellValue(list2.get(j));
            r.createCell(2).setCellValue(list3.get(j));

        }
        for (int i = 177; i <= 193; i = i + 2) {

            String str = driver.findElement(By.xpath(name_xpathBefore + i + name_xpathafter)).getText();
            list1.add(str);
            String str1 = driver.findElement(By.xpath(name_xpathBefore1 + i + name_xpathAfter1)).getText();
            list2.add(str1);
            String str2 = driver.findElement(By.xpath(name_xpathBefore2 + i + name_xpathAfter2)).getText();
            list3.add(str2);
        }
        for (int j =86 ; j<=95 ; j++){
            Row r = sheet.createRow(j);
            r.createCell(0).setCellValue(list1.get(j));
            r.createCell(1).setCellValue(list2.get(j));
            r.createCell(2).setCellValue(list3.get(j));

        }

        FileOutputStream output = new FileOutputStream(source);
        wb.write(output);



    }

}
