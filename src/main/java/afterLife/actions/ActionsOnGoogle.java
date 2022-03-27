package afterLife.actions;

import com.google.common.base.Strings;
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
import java.util.List;

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
        ArrayList<WebElement> l = (ArrayList<WebElement>) driver.findElements(tableOfCast);
        ArrayList<String> ll= new ArrayList<>();
        for (WebElement e : l){
            String s  = (e.getText());
            ll.add(s);

        }
        ArrayList<String> names = new ArrayList<>();
        for (int i=0 ; i < ll.size() ; i=i+3)
        {

            System.out.println(ll.get(i));
            names.add(ll.get(i));
        }
        System.out.println(names.size());
        ArrayList<String> names1 = new ArrayList<>();
        for (int i=1 ; i < ll.size() ; i=i+3)
        {

            System.out.println(ll.get(i));
            names1.add(ll.get(i));
        }
        System.out.println(names1.size());
        ArrayList<String> names2 = new ArrayList<>();
        for (int i=2 ; i <= ll.size() ; i=i+3)
        {

            System.out.println(ll.get(i));
            names2.add(ll.get(i));
        }
        System.out.println(names2.size());
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
            r.createCell(0).setCellValue(names.get(j));
            r.createCell(1).setCellValue(names1.get(j));
            r.createCell(2).setCellValue(names2.get(j));

        }


        FileOutputStream output = new FileOutputStream(source);
        wb.write(output);



    }
}
