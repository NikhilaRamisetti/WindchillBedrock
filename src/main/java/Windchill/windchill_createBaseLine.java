package Windchill;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
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

public class windchill_createBaseLine {
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
    public synchronized void initiatereader() {
        //ExcelReader excelReader = ExcelReader.getInstance("C:\\Program Files (x86)\\Accenture\\IXO\\Bedrock\\WebApplicationTestData", "TestData.xlsx", "Sheet2",
        // "Test");
        System.setProperty("webdriver.chrome.driver", "C:\\Program Files\\Accenture\\IX0\\Bedrock\\SetupFiles\\chromedriver\\chromedriver_win32\\chromedriver.exe");

    }


    WebDriver driver;
    ExtentTest logger;
    ExtentReports extent;
    String patternPath = System.getProperty("user.dir") + "\\PixelGraphics\\";


//    @BeforeClass
//    public void Prerequisite() {
//        extent = ReportFactory.getInstance();
//    }

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
       // logger = extent.startTest("LaunchWebClient");
        driver.get("http://windchilltest.accenture.com:82/Windchill/app");
        Thread.sleep(2000);
        ExcelReader credentialsReader= ExcelReader.getInstance(Root + "\\src\\main\\java\\Windchill","TestDataInput.xlsx","Credentials");
        List<String> excelData=credentialsReader.getRowData(3,0);
        String USERNAME=excelData.get(0);
        String PASSWORD=excelData.get(1);
        Robot rb = new Robot();
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
       // logger.log(LogStatus.PASS, "Test Case is Passed");
    }
    @Test(priority=1)
    public void CreateBL() throws InterruptedException{
        WebDriverWait wait = new WebDriverWait(driver, 30);
        logger = extent.startTest("CreateBaseLine");

        try {
            driver.findElement(By.xpath("//div[@class='x-tool x-tool-expand-west']")).click();
            Thread.sleep(2000);
            driver.findElement(By.id("navigatorTabPanel__object_main_navigation")).click();
            Thread.sleep(3000);
            driver.findElement(By.xpath("//img[@class='x-tree-ec-icon x-tree-elbow-plus']")).click();
            Thread.sleep(2000);
            driver.findElement(By.xpath("//span[text()='Details']")).click();
            Thread.sleep(4000);
            driver.findElement(By.xpath("//button[text()='Actions']")).click();
            Thread.sleep(2000);
            driver.findElement(By.xpath("//span[text()='New']")).click();
            Thread.sleep(2000);
            driver.findElement(By.xpath("//span[text()='New Baseline']")).click();//select actions>new >new baseline
            Set<String> winHandles = driver.getWindowHandles();
            ArrayList<String> list = new ArrayList<>(winHandles);
            driver.switchTo().window(list.get(1));
            ExcelReader credentialsReader = ExcelReader.getInstance(Root + "\\src\\main\\java\\Windchill", "TestDataInput.xlsx", "BaseLine_Creation");
            List<String> excelData = credentialsReader.getRowData(1, 0);
            String BaseLineName = excelData.get(0);
            String BaseLineDesrptn = excelData.get(1);
            driver.findElement(By.name("tcomp$attributesTable$OR:wt.pdmlink.PDMLinkProduct:101406$___name_col_name___textbox")).sendKeys(BaseLineName);//send name from the excel
            Thread.sleep(2000);
            driver.findElement(By.id("description")).sendKeys(BaseLineDesrptn);//send the baseline description from excel
            Thread.sleep(2000);
            WebElement lockNo = driver.findElement(By.name("tcomp$attributesTable$OR:wt.pdmlink.PDMLinkProduct:101406$___lock___radio"));//select lockno
            lockNo.click();
            WebElement auto = driver.findElement(By.name("tcomp$attributesTable$OR:wt.pdmlink.PDMLinkProduct:101406$___Location_col_overrideFolder___radio"));//select the folder path
            auto.click();
            Thread.sleep(2000);
            driver.findElement(By.xpath("//button[@accesskey='o']")).click();
            logger.log(LogStatus.PASS, "Test Case is Passed");

        }
        catch(Exception e){
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
