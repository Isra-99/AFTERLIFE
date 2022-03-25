import afterLife.actions.ActionsOnGoogle;
import afterLife.actions.ReadFromExcel;
import org.testng.annotations.Test;

import java.io.IOException;

public class AfterLifeBasicTesting extends BaseClass{
    @Test
    public void searchAfterLife() throws IOException {
        ActionsOnGoogle actionsOnGoogle = new ActionsOnGoogle(driver);
        ReadFromExcel readFromExcel = new ReadFromExcel(driver);
        actionsOnGoogle.enterSomething(readFromExcel.searchContent(), driver);
        actionsOnGoogle.clickSearchButton(driver);
        actionsOnGoogle.clickAfterLifeIMDB(driver);
        actionsOnGoogle.clickAllCast(driver);
       actionsOnGoogle.writeCastToExcel(driver);



    }
}
