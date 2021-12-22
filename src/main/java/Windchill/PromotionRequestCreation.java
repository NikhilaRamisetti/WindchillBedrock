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

public class PromotionRequestCreation {
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

    public PromotionRequestCreation() throws IOException {
    }

    @BeforeSuite
    public synchronized void initiateReader() throws IOException {
        System.setProperty("webdriver.chrome.driver", Root + "\\chromedriver_win32\\chromedriver.exe");

    }
    ExcelReader credentialsReader= ExcelReader.getInstance(Root + "\\src\\main\\java\\Windchill","TestDataInput.xlsx","Credentials");
    ExcelReader excelReader= ExcelReader.getInstance(Root + "\\src\\main\\java\\Windchill","TestDataInput.xlsx","PRCreation");
    WebDriver driver;
    ExtentReports extent;
    ExtentTest logger;
    List<String> DemoPRDetails = excelReader.getRowData(1, 0);


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
    public void createPromotionRequest() throws InterruptedException {
        logger = extent.startTest("Creating a Promotion Request");
        Thread.sleep(2000);
        commonFunctions.viewProductMenuSection("Folders", driver, DemoPRDetails.get(0));//opening folders from product Menu Section
        implicitWait(10);
        System.out.println(excelReader.getRowCount());
        System.out.println(DemoPRDetails.get(1));
        try {
            for (int i = 1; i <= excelReader.getRowCount(); i++) {
                List<String> ChangeRequestDetails = excelReader.getRowData(i, 0);
                Thread.sleep(2000);
                driver.findElement(By.xpath("//a[text()='"+ChangeRequestDetails.get(1)+"']")).click();
                Thread.sleep(1000);
                driver.findElement(By.xpath("//button[text()='Actions']")).click();
                driver.findElement(By.xpath("//span[text()='New']")).click();
                driver.findElement(By.xpath("//span[text()='New Promotion Request']")).click();
                Thread.sleep(2000);
                String parent = driver.getWindowHandle();
                //Switch to new part window from parent window
                if (commonFunctions.openNewWindowHandles("New Promotion Request", driver)) {
                    Thread.sleep(2000);
                    implicitWait(10);
                    //driver.findElement(By.xpath("//div[@class='cke_contents cke_reset']")).sendKeys(ChangeRequestDetails.get(2));
                    driver.findElement(By.xpath("//button[@accesskey='n']")).click();
                    Thread.sleep(2000);
                    driver.findElement(By.xpath("//select[@id='maturityState']")).click();
                    driver.findElement(By.xpath("//option[contains(text(),'"+ChangeRequestDetails.get(3)+"')]")).click();
                    Thread.sleep(2000);
                    driver.findElement(By.xpath("//div[@class='x-grid3-row-checker']")).click();
                    driver.findElement(By.xpath("//button[@style='background-image: url(\"netmarkets/images/defer_add.gif\");']")).click();
                    Thread.sleep(2000);
                    driver.findElement(By.xpath("//button[@accesskey='f']")).click();
                    Thread.sleep(5000);
                    implicitWait(10);
                    commonFunctions.closeWindowHandle(driver, parent);
                    driver.findElement(By.xpath("//a[text()='"+DemoPRDetails.get(0)+"']")).click();
                    Thread.sleep(1000);
                }
                logger.log(LogStatus.PASS, "Test Case is Passed");
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

