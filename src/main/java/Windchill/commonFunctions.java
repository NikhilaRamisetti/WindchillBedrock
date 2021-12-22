package Windchill;

import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;
import org.apache.commons.io.FileUtils;
import org.openqa.selenium.*;

import java.io.File;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Set;
import java.util.concurrent.TimeUnit;

public class commonFunctions {


    public static void viewProductMenuSection(String section, WebDriver driver, String ProductName){
        driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);
        try{
            driver.findElement(By.xpath("//a[@id='object_main_navigation_nav']")).click();
//        implicitWait(5);
//        driver.findElement(By.xpath("//li[@id='navigatorTabPanel__object_main_navigation']")).click();
            driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
            System.out.println(driver.findElements(By.xpath("//img[@class='x-tree-ec-icon x-tree-elbow-end-plus']")).size());
            if(driver.findElements(By.xpath("//img[@class='x-tree-ec-icon x-tree-elbow-end-plus']")).size() !=0) {
                driver.findElement(By.xpath("//span[text()='"+ProductName+"']")).click();
                driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);
            }
            driver.findElement(By.xpath("//span[text()='"+section+"']")).click();

        }
        catch(Exception e){
            System.out.println(e.getLocalizedMessage());
            //logger.log(LogStatus.ERROR, e.getLocalizedMessage());
        }
    }

    public static void createaNewACL(WebDriver driver, ExcelReader excelReader, ExtentTest logger) throws InterruptedException {
        driver.findElement(By.xpath("//a[@accesskey='W']")).click();
        Thread.sleep(2000);
        driver.findElement(By.xpath("//span[@class='x-tab-strip-text orgNavigation-icon']")).click();
        driver.findElement(By.xpath("//div[@id='orgNavigation']//div[@class='x-tree-node-el x-unselectable file x-tree-node-collapsed']")).click();
        driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
        Thread.sleep(2000);
        driver.findElement(By.xpath("//span[text()='Utilities']")).click();
        Thread.sleep(2000);
        driver.findElement(By.xpath("//a[contains(text(),'Policy Administration')]")).click();
        Thread.sleep(2000);
        ArrayList<String> tabs = new ArrayList<String> (driver.getWindowHandles());
        driver.switchTo().window(tabs.get(1));
        driver.findElement(By.xpath("//a[text()='Default']")).click();
        for (int i = 1; i <= excelReader.getRowCount(); i++) {
            List<String> ACLDetails = excelReader.getRowData(i, 0);
            driver.findElement(By.xpath("//i[@class='icon create-icon']")).click();
            Thread.sleep(2000);
            driver.findElement(By.xpath("//button[@id='newAccessTypeButton']")).click();
            Thread.sleep(2000);
            String parent = driver.getWindowHandle();
            //switch to query builder window
            if (commonFunctions.openNewWindowHandles("Find Type", driver)) {
                driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);

                driver.findElement(By.xpath("//input[@id='typePicker.tree.fitTextField']")).sendKeys(ACLDetails.get(0));//Passing Object type name as input
                driver.findElement(By.xpath("//input[@id='typePicker.tree.fitTextField']")).sendKeys(Keys.ENTER);
                driver.findElement(By.xpath("//div[@class='x-grid3-row ptc-grid-row-focus ptc-grid-row-checker-focus']//input[@type='radio']")).click();
                driver.findElement(By.xpath("//button[@accesskey='o']")).click();
                commonFunctions.closeWindowHandle(driver, parent);
            } else {
                logger.log(LogStatus.ERROR, "Find Type Window Not Found");
            }
            Thread.sleep(2000);
            driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
            driver.findElement(By.xpath("//span[text()='SEARCH']")).click();
            driver.findElement(By.xpath("//div[@class='ui-select-dropdown select2-drop select2-with-searchbox select2-drop-active']")).click();
            driver.findElement(By.xpath("//div[@class='ui-select-dropdown select2-drop select2-with-searchbox select2-drop-active']//input[@type='search']")).sendKeys(ACLDetails.get(1));//passing Participant name as input
            try{
                driver.findElement(By.xpath("//div[@class='select2-result-label ui-select-choices-row-inner']")).click();
            }
            catch(Exception e){
                System.out.println(e.getLocalizedMessage());
                logger.log(LogStatus.ERROR, e.getLocalizedMessage());
                logger.log(LogStatus.INFO, "The user doesn't exist");
            }
            Thread.sleep(2000);
            String grant = ACLDetails.get(2);
            String[] arrGrant = grant.split(",", 5);

            for (String grantProperty : arrGrant) {
                driver.findElement(By.xpath("//tr[@class='tableRow ng-pristine ng-untouched ng-valid ng-scope ng-not-empty']//td[text()='"+grantProperty+"']/following-sibling::td[1]")).click();
                System.out.println(grantProperty);
            }
            String deny = ACLDetails.get(3);
            String[] arrDeny = deny.split(",", 5);

            for (String denyProperty : arrDeny) {
                driver.findElement(By.xpath("//tr[@class='tableRow ng-pristine ng-untouched ng-valid ng-scope ng-not-empty']//td[text()='"+denyProperty+"']/following-sibling::td[2]")).click();
                System.out.println(denyProperty);
            }
            driver.findElement(By.xpath("//button[text()='OK']")).click();
            if(driver.findElements(By.xpath("//span[text()='An access control rule is already defined with the following attributes']")).size() !=0){
                logger.log(LogStatus.INFO, "An access rule for the object type and user already exists");
                if(ACLDetails.get(4).equals("Yes"))
                driver.findElement(By.xpath("//button[text()='Yes']")).click();
                else{
                    driver.findElement(By.xpath("//button[text()='No']")).click();
                }
            }
            Thread.sleep(2000);
        }
        driver.close();
        driver.switchTo().window(tabs.get(0));


    }

    public static boolean openNewWindowHandles(String WindowName,WebDriver driver) {
        String parent = driver.getWindowHandle();
        Set<String> windows = driver.getWindowHandles();
        for (String window : windows) {
            driver.switchTo().window(window);
            if (driver.getTitle().contains(WindowName)) {
                return true;
            }
        }
        return false;
    }

    public static void closeWindowHandle(WebDriver driver, String parent) throws InterruptedException {
        driver.switchTo().window(parent);
        Thread.sleep(2000);
    }
    public static String getScreenshot(WebDriver driver, String screenshotName, String FolderName, String RootFolder) throws Exception {
        String dateName = new SimpleDateFormat("yyyyMMddhhmmss").format(new Date());
        TakesScreenshot ts = (TakesScreenshot) driver;
        File source = ts.getScreenshotAs(OutputType.FILE);
        String destination = RootFolder + FolderName + screenshotName + dateName + ".png";
        File finalDestination = new File(destination);
        FileUtils.copyFile(source, finalDestination);
        return destination;
    }
    /*
    public static void runTest(Runnable code, ExtentTest logger) {
        try {
            code.run();
        } catch(Exception e) {
                System.out.println(e.getLocalizedMessage());
                logger.log(LogStatus.ERROR, e.getLocalizedMessage());
        }
    }
    */
}
