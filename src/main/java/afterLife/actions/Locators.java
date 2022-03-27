package afterLife.actions;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;

import java.io.IOException;
import java.util.ArrayList;

public class Locators {
    /// Constructor of locators
    public  Locators (WebDriver d){}
    /// WebElement of searchBar
    static public WebElement searchBar(WebDriver driver){
    return driver.findElement(By.xpath("//input[ @class='gLFyf gsfi']"));
    }
    /// WebElement of Button
    static public WebElement button(WebDriver driver){
        return driver.findElement(By.xpath("//div[@class='FPdoLc lJ9FBc']//input[@class='gNO89b']"));
    }
    static public WebElement afterLifeLink(WebDriver driver) throws IOException {
        ReadFromExcel readFromExcel = new ReadFromExcel(driver);
        String text = "//h3[contains(text(),'" + readFromExcel.searchLink() + "')]";
        //return driver.findElement(By.xpath("//h3[contains(text(),'After Life (TV Series 2019–2022) - IMDb')]"));
        return driver.findElement(By.xpath(text));
    }
    static public WebElement allCast(WebDriver driver){
        return driver.findElement(By.linkText("All cast & crew"));
    }
    //// Locator for list cast names, screen names and their appearances
    static public By tableOfCast = By.xpath("//table[@class='cast_list']//tr//td//a[not(.//img)]");
//// other option for xpath is //table[@class='cast_list']//tr//a[not(descendant::img)]




}
