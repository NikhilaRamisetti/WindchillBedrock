package Windchill;


import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.ITestResult;
import org.testng.annotations.*;

import java.awt.*;
import java.awt.datatransfer.StringSelection;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.IOException;
import java.util.List;
import java.util.concurrent.TimeUnit;

public class ManageWindchillPart {


    private static File directory =  new File("");
    private static String Root;

    static {
        try {
            Root = directory.getCanonicalPath();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private ManageWindchillPart CommonAppsDriver;

    public ManageWindchillPart() throws IOException {
    }

    @BeforeSuite
    public synchronized void initiateReader() throws IOException {
        System.setProperty("webdriver.chrome.driver", Root + "\\chromedriver_win32\\chromedriver.exe");//setting driver property

    }

    ExcelReader credentialsReader= ExcelReader.getInstance(System.getProperty("user.dir") + "/GlobalSettings", Root + "\\src\\main\\java\\Windchill","TestDataInput.xlsx","Credentials");
    ExcelReader excelReader= ExcelReader.getInstance(System.getProperty("user.dir") + "/GlobalSettings", Root + "\\src\\main\\java\\Windchill","TestDataInput.xlsx","Part_Creation");
    WebDriver driver;
    ExtentReports extent;
    ExtentTest logger;
    java.util.List<String> DemoPartDetails= excelReader.getRowData(1,0);


    @BeforeClass
    public void Prerequisite() throws IOException {
        extent = ReportFactory.getInstance();
    }
    @BeforeTest
    public void initializeBrowser() {
        driver = new ChromeDriver();//initialize browser
        driver.manage().timeouts().pageLoadTimeout(60, TimeUnit.SECONDS);
        driver.manage().timeouts().implicitlyWait(60, TimeUnit.SECONDS);
        driver.manage().deleteAllCookies();
        driver.manage().window().maximize();//maximize browser
    }
    @Test(priority = 1)
    public void LaunchWindchillBrowser() throws InterruptedException, AWTException {
        logger = extent.startTest("Launching Windchill Browser");
        driver.get("http://windchilltest.accenture.com:81/Windchill/app");//GettingURL
        driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
        Thread.sleep(2000);
        Robot rb = new Robot();
        List<String> cred1= credentialsReader.getRowData(1,0);
        StringSelection userName = new StringSelection(cred1.get(0));//reading username from excel
        Toolkit.getDefaultToolkit().getSystemClipboard().setContents(userName, null);
        rb.keyPress(KeyEvent.VK_CONTROL);rb.keyPress(KeyEvent.VK_V);
        rb.keyRelease(KeyEvent.VK_CONTROL);rb.keyRelease(KeyEvent.VK_V);
        Thread.sleep(2000);
        rb.keyPress(KeyEvent.VK_TAB);
        StringSelection password = new StringSelection(cred1.get(1));//reading password from excel
        Toolkit.getDefaultToolkit().getSystemClipboard().setContents(password, null);
        rb.keyPress(KeyEvent.VK_CONTROL);rb.keyPress(KeyEvent.VK_V);
        rb.keyRelease(KeyEvent.VK_CONTROL);rb.keyRelease(KeyEvent.VK_V);
        Thread.sleep(2000);
        rb.keyPress(KeyEvent.VK_ENTER);
        System.out.println("Success: Launched Windchill Browser");
        logger.log(LogStatus.PASS, "Test Case is Passed");
    }

    /*
    @Test(priority=5)
    public void createPartwithAttachment() throws InterruptedException {
        implicitWait(10);
        driver.findElement(By.xpath("//button[@style='background-image: url(\"netmarkets/images/newpart.gif\");']")).click();
        String parent=driver.getWindowHandle();
        Set<String> windows = driver.getWindowHandles();
        System.out.println(windows);
        for (String window : windows)
        {
            driver.switchTo().window(window);
            System.out.println(driver.switchTo().window(window).getTitle());
            if (driver.getTitle().contains("New Part"))
            {
                implicitWait(10);
                driver.findElement(By.xpath("//select[@id='!~objectHandle~partHandle~!createType']")).click();
                implicitWait(2);
                driver.findElement(By.xpath("//option[text()=' Part ']")).click();
                implicitWait(10);
                Thread.sleep(5000);
                driver.findElement(By.xpath("//td[@attrid='name']/input[1]")).sendKeys("DemoPartWithCADDocument");
                implicitWait(10);
                driver.findElement(By.xpath("//input[@name='!~objectHandle~partHandle~!null___createCADDocForPart___createCADDocForPart!~objectHandle~partHandle~!']")).click();
                Thread.sleep(5000);
                driver.findElement(By.xpath("//button[@accesskey='n']")).click();
                Thread.sleep(5000);
                driver.findElement(By.xpath("//select[@id='DocTemplate']")).click();
                driver.findElement(By.xpath("//option[text()=' PartTemplate.CATPart ']")).click();
                Thread.sleep(5000);
                driver.findElement(By.xpath("//button[@accesskey='f']")).click();
                Thread.sleep(5000);
                System.out.println("Part with CAD Document created successfully");
            }
        }
        driver.switchTo().window(parent);
    }
    */
    @Test(priority=2)
    public void createPart() throws InterruptedException {
        logger = extent.startTest("Creating a Part");//log start message
        commonFunctions.viewProductMenuSection("Folders", driver, DemoPartDetails.get(3));//opening folders from product Menu Section
        implicitWait(10);
        try {
            for (int i = 1; i < excelReader.getRowCount(); i++) {
                List<String> PartDetails = excelReader.getRowData(i, 0);//reading partDetails from excel
                String PartName = new String(PartDetails.get(0));
                Thread.sleep(2000);
                driver.findElement(By.xpath("//div[@id='folderbrowser_PDM.toolBar']//button[text()='Actions']")).click();
                driver.findElement(By.xpath("//span[text()='New']")).click();
                driver.findElement(By.xpath("//span[text()='New Part']")).click();
                Thread.sleep(2000);
                String parent= driver.getWindowHandle();
                //Switch to new part window from parent window
                    if (commonFunctions.openNewWindowHandles("New Part",driver)) {
                        implicitWait(10);
                        driver.findElement(By.xpath("//select[@id='!~objectHandle~partHandle~!createType']")).click();
                        implicitWait(2);
                        driver.findElement(By.xpath("//option[text()=' Part ']")).click();
                        implicitWait(10);
                        Thread.sleep(5000);
                        driver.findElement(By.xpath("//td[@attrid='name']/input[1]")).sendKeys(PartName);
                        Thread.sleep(2000);
                        driver.findElement(By.xpath("//button[@accesskey='f']")).click();
                        Thread.sleep(5000);
                        System.out.println("Part created successfully");
                    }
                driver.close();
                commonFunctions.closeWindowHandle(driver, parent);
                Thread.sleep(2000);
            }
            logger.log(LogStatus.PASS, "Test Case is Passed");//log report result message
        }
        catch(Exception e){
            System.out.println(e.getLocalizedMessage());
            logger.log(LogStatus.ERROR, e.getLocalizedMessage());//log report error message
        }
    }

    @Test(priority=3)
    public void GetPartInfo() throws InterruptedException {
        logger = extent.startTest("Get Part Information");
        commonFunctions.viewProductMenuSection("Folders", driver, DemoPartDetails.get(3));
        implicitWait(10);
        Thread.sleep(2000);
        ViewPartInfoPage(DemoPartDetails.get(0));//passing the name of the part and open its info page
        try {
            String GeneralDetails = driver.findElement(By.xpath("//div[@id='dataStoreGeneral']")).getText();//general part info copied
            implicitWait(2);
            logger.log(LogStatus.INFO, GeneralDetails);
            System.out.println("Part Information retrieved successfully");
            logger.log(LogStatus.PASS, "Test Case is Passed");
        }
        catch(Exception e){
            System.out.println(e.getLocalizedMessage());
            logger.log(LogStatus.ERROR, e.getLocalizedMessage());
        }
    }

    @Test(priority = 4)
    public void deletePart() throws InterruptedException {
        logger = extent.startTest("Delete a Part");
        Thread.sleep(2000);
        commonFunctions.viewProductMenuSection("Folders", driver, DemoPartDetails.get(3));
        implicitWait(10);
        Thread.sleep(2000);
        ViewPartInfoPage(DemoPartDetails.get(0));
        try {
            driver.findElement(By.xpath("//button[text()='Actions']")).click();
            Thread.sleep(2000);
            driver.findElement(By.xpath("//span[text()='Delete']")).click();
            Thread.sleep(2000);
            String parent = driver.getWindowHandle();
            //switch to delete window from parent window
            if (commonFunctions.openNewWindowHandles("Delete",driver)) {
                    implicitWait(10);
                    driver.findElement(By.xpath("//div[@class='x-grid3-row-checker']")).click();
                    driver.findElement(By.xpath("//button[@accesskey='o']")).click();
                }
            commonFunctions.closeWindowHandle(driver, parent);
            Thread.sleep(2000);
            System.out.println("Part deleted successfully");
            logger.log(LogStatus.PASS, "Test Case is Passed");
        }
        catch(Exception e){
            System.out.println(e.getLocalizedMessage());
            logger.log(LogStatus.ERROR, e.getLocalizedMessage());
        }
    }


    public void ViewPartInfoPage(String Part){
        String partName = Part;
        try {
            List<WebElement> Parts= driver.findElements(By.xpath("//td//a[contains(text(),'" + partName + "')]"));
            for(WebElement part: Parts) {
                if(driver.findElements(By.xpath("//td//a[contains(text(),'" + partName + "')]/../../following-sibling::td[@class='x-grid3-col x-grid3-cell x-grid3-td-version ']//div[contains(text(),'" + DemoPartDetails.get(1) + "')]")).size() !=0){
                    part.click();
                }
                else{
                    logger.log(LogStatus.ERROR, "Re-Check the version of the part, It doesnt exist");
                }
            }
        }
        catch(Exception e){
            System.out.println(e.getLocalizedMessage());
            logger.log(LogStatus.ERROR, e.getLocalizedMessage());
        }
        implicitWait(5);
    }
    public void implicitWait(int seconds){
        driver.manage().timeouts().implicitlyWait(seconds, TimeUnit.SECONDS);
    }
    @AfterTest
    public void closeBrowser() {
        if (driver != null) {
            driver.close();
            driver.quit();
        }
    }

    @AfterMethod
    public void getResult(ITestResult result) throws Exception {
        if (result.getStatus() == ITestResult.FAILURE) {
            logger.log(LogStatus.FAIL, "Test Case Failed is " + result.getName());
            logger.log(LogStatus.FAIL, "Reason behind the failure " + result.getThrowable());
            String screenshotPath = commonFunctions.getScreenshot(driver, result.getName(),"/FailedTestsScreenshots/",Root);
            logger.log(LogStatus.FAIL, logger.addScreenCapture(screenshotPath));
        } else if (result.getStatus() == ITestResult.SKIP) {
            logger.log(LogStatus.SKIP, "Test Case Skipped is " + result.getName());
        }
        extent.endTest(logger);
    }

    @AfterClass
    public void endReport() {
        extent.flush();
    }

}

