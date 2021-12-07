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
import java.util.Set;
import java.util.concurrent.TimeUnit;

public class customAction {
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
    private customAction CommonAppsDriver;

    public customAction() throws IOException {
    }

    @BeforeSuite
    public synchronized void initiateReader() throws IOException {
        System.setProperty("webdriver.chrome.driver", Root + "\\chromedriver_win32\\chromedriver.exe");

    }

    ExcelReader credentialsReader= ExcelReader.getInstance(System.getProperty("user.dir") + "/GlobalSettings", Root + "\\src\\main\\java\\Windchill","TestDataInput.xlsx","Credentials");
    ExcelReader excelReader= ExcelReader.getInstance(System.getProperty("user.dir") + "/GlobalSettings", Root + "\\src\\main\\java\\Windchill","TestDataInput.xlsx","CustomActions");
    java.util.List<String> actionDetails= excelReader.getRowData(1,0);
    WebDriver driver;
    ExtentReports extent;
    ExtentTest logger;
    //String DefaultPart = "Test Part1";


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
    @Test(priority=2)
    public void createCustomAction() throws InterruptedException {
        logger = extent.startTest("Creating a custom Action");
        viewCustomizationMenuSection();
        Thread.sleep(2000);
        try {
            driver.findElement(By.xpath("//button[text()='Search']")).click();
            //passing the action model name parameter
            driver.findElement(By.xpath("//a[@href='ptc1/carambola/tools/actionReport/actionModelDetails?actionModelName="+actionDetails.get(0)+"']//img")).click();
            driver.findElement(By.xpath("//button[text()='Actions']")).click();
            driver.findElement(By.xpath("//span[text()='Create Action']")).click();
            Thread.sleep(2000);
            openNewWindowHandles();
            implicitWait(10);
            driver.findElement(By.xpath("//input[@id='description']")).click();
            driver.findElement(By.xpath("//input[@id='description']")).sendKeys(actionDetails.get(1));//Passing the Action label parameter
            implicitWait(10);
            driver.findElement(By.xpath("//input[@id='name']")).sendKeys(actionDetails.get(2));//Passing the Action name parameter
            implicitWait(10);
            driver.findElement(By.xpath("//input[@id='objectType']")).sendKeys(actionDetails.get(3));//Passing the object type parameter
            Thread.sleep(2000);
            driver.findElement(By.xpath("//button[@accesskey='f']")).click();
            commonFunctions.closeWindowHandle(driver, parent);
            logger.log(LogStatus.PASS, "Test Case is Passed");
        }
        catch(Exception e){
            System.out.println(e.getLocalizedMessage());
            logger.log(LogStatus.ERROR, e.getLocalizedMessage());
        }

    }
    @Test(priority=3)
    public void DeleteCustomAction() throws InterruptedException {
        logger = extent.startTest("Delete Custom Action");
        viewCustomizationMenuSection();
        Thread.sleep(2000);
        try {
            driver.findElement(By.xpath("//button[text()='Search']")).click();
            Thread.sleep(2000);
            driver.findElement(By.xpath("//a[@href='ptc1/carambola/tools/actionReport/actionModelDetails?actionModelName="+actionDetails.get(0)+"']//img")).click();
            Thread.sleep(2000);
            driver.findElement(By.xpath("//div[contains(text(),'" + actionDetails.get(1) + "')]/../preceding-sibling::td[@class='x-grid3-col x-grid3-cell x-grid3-td-checker x-grid3-cell-first ']//div[@class='x-grid3-row-checker']")).click();
            driver.findElement(By.xpath("//button[text()='Actions']")).click();
            driver.findElement(By.xpath("//span[text()='Remove Action From Model']")).click();
            Thread.sleep(2000);
            logger.log(LogStatus.PASS, "Test Case is Passed");
        }
        catch(Exception e){
            System.out.println(e.getLocalizedMessage());
            logger.log(LogStatus.ERROR, e.getLocalizedMessage());
        }
    }
    public void viewCustomizationMenuSection() throws InterruptedException {
        implicitWait(15);
        try {
            driver.findElement(By.xpath("//a[@id='object_main_navigation_nav']")).click();
            implicitWait(5);
            driver.findElement(By.xpath("//span[@class='x-tab-strip-text customizationNavigation-icon']")).click();
            implicitWait(5);
            driver.findElement(By.xpath("//span[text()='Tools']")).click();
            implicitWait(5);
            implicitWait(20);
            Thread.sleep(2000);
            driver.findElement(By.xpath("//a[contains(text(),'Action Model')]")).click();
            Thread.sleep(2000);
            driver.findElement(By.xpath("//input[@name='modelName']")).sendKeys(actionDetails.get(0));
        }
        catch(Exception e){
            System.out.println(e.getLocalizedMessage());
            logger.log(LogStatus.ERROR, e.getLocalizedMessage());
        }
    }
    public void openNewWindowHandles() {
        parent = driver.getWindowHandle();
        Set<String> windows = driver.getWindowHandles();
        System.out.println(windows.size());
        for (String window : windows) {
            driver.switchTo().window(window);
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

