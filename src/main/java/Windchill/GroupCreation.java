package Windchill;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
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

public class GroupCreation {
    private static File directory =  new File("");
    private static String Root;
    private static String parent;
    static {
        try {
            Root = directory.getCanonicalPath();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private ManageWindchillReport CommonAppsDriver;

    public GroupCreation() throws IOException {
    }

    @BeforeSuite
    public synchronized void initiateReader() throws IOException {
        System.setProperty("webdriver.chrome.driver", Root + "\\chromedriver_win32\\chromedriver.exe");

    }
    ExcelReader credentialsReader= ExcelReader.getInstance(Root + "\\src\\main\\java\\Windchill","TestDataInput.xlsx","Credentials");
    ExcelReader excelReader= ExcelReader.getInstance(Root + "\\src\\main\\java\\Windchill","TestDataInput.xlsx","GroupCreation");
    WebDriver driver;
    ExtentReports extent;
    ExtentTest logger;
    List<String> DemoGroupDetails = excelReader.getRowData(0, 0);
    List<String> DemoGroupDetails1 = excelReader.getRowData(1, 0);


    @BeforeClass
    public void Prerequisite() throws IOException {
        extent = ReportFactory.getInstance();
    }
    @BeforeTest
    public void initializeBrowser() {
        driver = new ChromeDriver();
        driver.manage().timeouts().pageLoadTimeout(60, TimeUnit.SECONDS);
        driver.manage().timeouts().implicitlyWait(60, TimeUnit.SECONDS);
        driver.manage().deleteAllCookies();
        driver.manage().window().maximize();


    }
    @Test(priority = 1)
    public void LaunchWindchillBrowser() throws InterruptedException, AWTException {
        logger = extent.startTest("Launching Windchill Browser");
        driver.get("http://windchilltest.accenture.com:81/Windchill/app");//GettingURL
        driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
        Thread.sleep(2000);
        Robot rb = new Robot();
        List<String> cred1= credentialsReader.getRowData(2,0);
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

    @Test(priority=2)
    public void createGroup() throws InterruptedException {
        System.out.println(DemoGroupDetails.get(0));
        System.out.println(DemoGroupDetails1.get(0));
        logger = extent.startTest("Creating a User Defined Group");
        driver.findElement(By.xpath("//a[@accesskey='W']")).click();
        Thread.sleep(2000);
        driver.findElement(By.xpath("//span[@class='x-tab-strip-text siteNavigation-icon']")).click();
        driver.findElement(By.xpath("//span[text()='Utilities']")).click();
        Thread.sleep(2000);
        try {
            driver.findElement(By.xpath("//a[contains(text(),'Participant Administration')]")).click();
            Thread.sleep(2000);
            driver.findElement(By.xpath("//button[@style='background-image: url(\"netmarkets/images/group_new.gif\");']")).click();
            Thread.sleep(2000);
            parent = driver.getWindowHandle();
            //switch to query builder window
            if (commonFunctions.openNewWindowHandles("New Group", driver)) {
                implicitWait(10);//Add report constraints
                driver.findElement(By.xpath("//td[@attrid='name']//input[@type='text']")).click();
                driver.findElement(By.xpath("//td[@attrid='name']//input[@type='text']")).sendKeys(DemoGroupDetails1.get(0));
                driver.findElement(By.xpath("//button[@accesskey='n']")).click();
                implicitWait(10);
                driver.findElement(By.xpath("//button[@style='background-image: url(\"netmarkets/images/add16x16.gif\");']")).click();
                String newGroupWindow= driver.getWindowHandle();
                if (commonFunctions.openNewWindowHandles("Find Participant", driver)) {
                    driver.findElement(By.xpath("//input[@id='fullName0_SearchTextBox']")).click();
                    driver.findElement(By.xpath("//input[@id='fullName0_SearchTextBox']")).sendKeys(DemoGroupDetails1.get(1));
                    driver.findElement(By.xpath("//button[text()='Search']")).click();
                    driver.findElement(By.xpath("//button[@accesskey='o']")).click();
                }
                else {
                    logger.log(LogStatus.ERROR, " Add Participant Window Not Found");
                }
                Thread.sleep(3000);
                commonFunctions.closeWindowHandle(driver, newGroupWindow);
                driver.findElement(By.xpath("//button[@accesskey='f']")).click();
                commonFunctions.closeWindowHandle(driver, parent);
                logger.log(LogStatus.PASS, "Test Case is Passed");
            } else {
                logger.log(LogStatus.ERROR, "New group template Window Not Found");
            }
        }
        catch(Exception e){
            System.out.println(e.getLocalizedMessage());
            logger.log(LogStatus.ERROR, e.getLocalizedMessage());
        }
    }

    public void implicitWait(int seconds){
        driver.manage().timeouts().implicitlyWait(seconds, TimeUnit.SECONDS);
    }
/*
    @AfterTest
    public void closeBrowser() {
        if (driver != null) {
            driver.close();
            driver.quit();
        }
    }
*/
    @AfterMethod
    public void getResult(ITestResult result) throws Exception {
        if (result.getStatus() == ITestResult.FAILURE) {
            logger.log(LogStatus.FAIL, "Test Case Failed is " + result.getName());
            logger.log(LogStatus.FAIL, "Reason behind the failure " + result.getThrowable());
            String screenshotPath = commonFunctions.getScreenshot(driver, result.getName(), "/FailedTestsScreenshots/", Root);
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

