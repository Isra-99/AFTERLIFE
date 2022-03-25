import afterLife.actions.ActionsOnGoogle;
import afterLife.actions.ReadFromExcel;
import net.sf.jxls.reader.XLSReader;
import net.sf.jxls.reader.XLSReaderImpl;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import java.io.IOException;

public class BaseClass {
    static public WebDriver driver;
    @BeforeTest
    public void search() throws IOException {
       ///creating path to open chrome
       System.setProperty("webdriver.chrome.driver", "src/main/resources/chromedriver.exe");
       ///creating object of chrome
        driver = new ChromeDriver();
        ReadFromExcel readFromExcel= new ReadFromExcel(driver);
        String text = readFromExcel.url();
       driver.get(text);
   }

   @AfterTest
    public void closeBrowser(){
        driver.quit();
   }


}


