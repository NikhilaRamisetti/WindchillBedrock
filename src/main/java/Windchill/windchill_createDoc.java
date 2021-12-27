package Windchill;

//import WebAppExample.ExcelReader;
import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.ITestResult;
import org.testng.annotations.*;

import java.awt.*;
import java.awt.datatransfer.StringSelection;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Set;
import java.util.concurrent.TimeUnit;

public class windchill_createDoc {
    private static File directory =  new File("");
    private static String Root;

    static {
        try {
            Root = directory.getCanonicalPath();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    @BeforeSuite
    public synchronized void initiateReader() throws IOException {
        System.setProperty("webdriver.chrome.driver", Root + "\\chromedriver_win32\\chromedriver.exe");//setting driver property

    }

    static ExcelReader globalSettings;
    private static ExcelReader excelReader;
    WebDriver driver;
    ExtentTest logger;
    ExtentReports extent;
    String patternPath = System.getProperty("user.dir") + "\\PixelGraphics\\";


    @BeforeClass
    public void Prerequisite() throws IOException {
        extent = ReportFactory.getInstance();
    }

    @BeforeTest
    public void initializebrowser() {

        driver = new ChromeDriver();
        driver.manage().timeouts().pageLoadTimeout(60, TimeUnit.SECONDS);
        driver.manage().timeouts().implicitlyWait(60, TimeUnit.SECONDS);
        driver.manage().deleteAllCookies();
        driver.manage().window().maximize();
    }

    @Test(priority = 0)
    public void LaunchWebClient() throws InterruptedException, AWTException {
        logger = extent.startTest("LaunchWebClient");
        try {
            driver.get("http://windchilltest.accenture.com:81/Windchill/app");
            Thread.sleep(2000);
            Robot rb = new Robot();
            ExcelReader credentialsReader = ExcelReader.getInstance( Root + "\\src\\main\\java\\Windchill", "TestDataInput.xlsx", "Credentials");
            List<String> excelData = credentialsReader.getRowData(2, 0);
            String USERNAME = excelData.get(0);
            String PASSWORD = excelData.get(1);
            StringSelection str = new StringSelection(USERNAME);
            Toolkit.getDefaultToolkit().getSystemClipboard().setContents(str, null);
            rb.keyPress(KeyEvent.VK_CONTROL);
            rb.keyPress(KeyEvent.VK_V); // press Contol+V for pasting
            rb.keyRelease(KeyEvent.VK_CONTROL);
            rb.keyRelease(KeyEvent.VK_V);// release Contol+V for pasting
            Thread.sleep(2000);
            rb.keyPress(KeyEvent.VK_TAB);
            StringSelection str1 = new StringSelection(PASSWORD);
            Toolkit.getDefaultToolkit().getSystemClipboard().setContents(str1, null);
            rb.keyPress(KeyEvent.VK_CONTROL);
            rb.keyPress(KeyEvent.VK_V); // press Contol+V for pasting
            rb.keyRelease(KeyEvent.VK_CONTROL);
            rb.keyRelease(KeyEvent.VK_V);// release Contol+V for pasting
            Thread.sleep(2000);
            rb.keyPress(KeyEvent.VK_ENTER);
            logger.log(LogStatus.PASS, "Test Case is Passed");
        }
        catch(Exception e){
                System.out.println(e.getLocalizedMessage());
                logger.log(LogStatus.ERROR, e.getLocalizedMessage());
            }
    }

    @Test(priority = 2)
    public void createDoc() throws InterruptedException {
        logger = extent.startTest("createDoc");
        try {
            WebDriverWait wait = new WebDriverWait(driver, 30);
            Thread.sleep(3000);
            driver.findElement(By.xpath("//div[@class='x-tool x-tool-expand-west']")).click();
            Thread.sleep(2000);
            driver.findElement(By.id("navigatorTabPanel__object_main_navigation")).click();
            driver.findElement(By.xpath("//span[@class='x-tab-strip-text productNavigation-icon']")).click();
            driver.findElement(By.xpath("//a[text()='View All']")).click();
            Thread.sleep(3000);
            driver.findElement(By.xpath("//a[text()='product2']")).click();//select view all>product 2
            Thread.sleep(4000);
            wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[@id='folderbrowser_PDM.toolBar']//button[text()='Actions']"))).click();
            driver.findElement(By.xpath("//span[text()='New']")).click();
            driver.findElement(By.xpath("//span[text()='New Document']")).click();//select new>new documnet
            Set<String> winHandles = driver.getWindowHandles();
            ArrayList<String> list = new ArrayList<>(winHandles);
            driver.switchTo().window(list.get(1));
            Thread.sleep(4000);
            Select type = new Select(driver.findElement(By.id("createType")));//select the dropdown createtype
            type.selectByVisibleText("Document");
            Thread.sleep(2000);
            Select temp = new Select(driver.findElement(By.id("templatesCombo")));//select the template
            temp.selectByVisibleText("Requirements Template");
            Thread.sleep(3000);
            driver.findElement(By.xpath("//td[@attrid='name']/input[1]")).sendKeys("newDocument2");
            WebElement auto = driver.findElement(By.name("tcomp$attributesTable$OR:wt.pdmlink.PDMLinkProduct:98715$___Location_col_overrideFolder___radio"));
            auto.click();
            Thread.sleep(2000);
            driver.findElement(By.id("ext-gen37")).click();
            Thread.sleep(2000);
            driver.findElement(By.id("ext-gen39")).click();
            logger.log(LogStatus.PASS, "Test Case is Passed");
        } catch (Exception e) {
            System.out.println(e.getLocalizedMessage());
            logger.log(LogStatus.ERROR, e.getLocalizedMessage());
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
