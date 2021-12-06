package Windchill;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.testng.ITestResult;
import org.testng.annotations.*;

import java.awt.*;
import java.awt.datatransfer.StringSelection;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.IOException;
import java.util.List;
import java.util.concurrent.TimeUnit;

public class ManageWindchillReport {
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

    public ManageWindchillReport() throws IOException {
    }

    @BeforeSuite
    public synchronized void initiateReader() throws IOException {
        System.setProperty("webdriver.chrome.driver", Root + "\\chromedriver_win32\\chromedriver.exe");

    }
    ExcelReader credentialsReader= ExcelReader.getInstance(Root + "\\src\\main\\java\\Windchill","TestDataInput.xlsx","Credentials");
    ExcelReader excelReader= ExcelReader.getInstance(Root + "\\src\\main\\java\\Windchill","TestDataInput.xlsx","ReportsManagement");
    WebDriver driver;
    ExtentReports extent;
    ExtentTest logger;
    List<String> DemoReportDetails = excelReader.getRowData(1, 0);



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
    public void createReport() throws InterruptedException {
        logger = extent.startTest("Creating a Report");
        commonFunctions.viewProductMenuSection("Utilities", driver, DemoReportDetails.get(0));//open products menu section
        Thread.sleep(2000);
        try {
            driver.findElement(By.xpath("//a[contains(text(),'Report Management')]")).click();
            Thread.sleep(2000);
            driver.findElement(By.xpath("//button[@style='background-image: url(\"netmarkets/images/report_template_new.png\");']")).click();
            Thread.sleep(2000);
            //switch to query builder window
            if (commonFunctions.openNewWindowHandles("Query Builder", driver)) {
                implicitWait(10);//Add report constraints
                driver.findElement(By.xpath("//input[@id='reportTemplateName']")).click();
                driver.findElement(By.xpath("//input[@id='reportTemplateName']")).sendKeys(DemoReportDetails.get(1));
                implicitWait(10);
                driver.findElement(By.xpath("//textarea[@id='description']")).click();
                driver.findElement(By.xpath("//textarea[@id='description']")).sendKeys("Demo Report for Bedrock Automation");
                Thread.sleep(1000);
                driver.findElement(By.xpath("//div[@id='tablesAndJoins']")).click();
                Thread.sleep(1000);
                driver.findElement(By.xpath("//button[@id='addTable']")).click();
                Thread.sleep(1000);
                driver.findElement(By.xpath("//a[text()='Abs Collection Criteria']")).click();
                implicitWait(10);
                driver.findElement(By.xpath("//button[text()='OK']")).click();
                Thread.sleep(2000);
                driver.findElement(By.xpath("//div[@id='selectOrConstrain']")).click();
                Thread.sleep(1000);
                driver.findElement(By.xpath("//div[@id='toolbarAddButton']")).click();
                implicitWait(10);
                driver.findElement(By.xpath("//a[text()='Reportable Item']")).click();
                Thread.sleep(1000);
                driver.findElement(By.xpath("//span[@id='TREETOPwt.dataops.objectcol.AbsCollectionCriteria']")).click();
                implicitWait(10);
                driver.findElement(By.xpath("//button[text()='OK']")).click();
                Thread.sleep(1000);
                driver.findElement(By.xpath("//button[text()='Apply']")).click();
                Thread.sleep(1000);
                driver.findElement(By.xpath("//button[text()='Close']")).click();
                Thread.sleep(1000);
                commonFunctions.closeWindowHandle(driver, parent);
                logger.log(LogStatus.PASS, "Test Case is Passed");
            } else {
                logger.log(LogStatus.ERROR, "Window Not Found");
            }
        }
        catch(Exception e){
            System.out.println(e.getLocalizedMessage());
            logger.log(LogStatus.ERROR, e.getLocalizedMessage());
        }
        }

    @Test(priority=3)
    public void ViewReport() throws Exception {
        logger = extent.startTest("View Report");
        try {
            driver.get("http://windchilltest.accenture.com:81/Windchill/app/#ptc1/comp/reporting.product.listReports");
            implicitWait(10);
            Thread.sleep(2000);
            driver.findElement(By.xpath("//div[@class='ux-maximgb-tg-elbow-active ux-maximgb-tg-elbow-plus']")).click();
            implicitWait(10);
            driver.findElement(By.xpath("//a[text()='Change Notice Log']")).click();
            Thread.sleep(2000);
            if (commonFunctions.openNewWindowHandles("View Change Notice Log",driver)) {
                Thread.sleep(2000);
                driver.findElement(By.xpath("//button[text()='Generate']")).click();
                implicitWait(2);
                Thread.sleep(2000);
                Thread.sleep(5000);
                //Click Manually for the next driver click
                //driver.findElement(By.xpath("//*[@value='Generate']")).click();
                JavascriptExecutor jse = (JavascriptExecutor) driver;
                jse.executeScript("window.confirm('Please click on generate');");
                Thread.sleep(2000);
                commonFunctions.getScreenshot(driver, "Change_Notice_Log", "/Screenshots/", Root);//save the report as screenshot
                Thread.sleep(2000);
                driver.close();
                commonFunctions.closeWindowHandle(driver, parent);
                logger.log(LogStatus.PASS, "Test Case is Passed");
            } else {
                logger.log(LogStatus.ERROR, "Window Not Found");
            }
        }
        catch(Exception e){
            System.out.println(e.getLocalizedMessage());
            logger.log(LogStatus.ERROR, e.getLocalizedMessage());
        }

    }
    //Edit Report is only possible if logged In as Admin (wcadmin)
    /*
    @Test(priority = 4)
    public void EditReport() throws InterruptedException {
        logger = extent.startTest("Edit a Report");
        Thread.sleep(2000);
        viewProductMenuSection("Reports");
        Thread.sleep(2000);
        driver.findElement(By.xpath("//div[@class='ux-maximgb-tg-elbow-active ux-maximgb-tg-elbow-plus']")).click();
        Thread.sleep(2000);
        driver.findElement(By.xpath("//a[contains(text(),'Change Notice Log')]/../..//following-sibling::td[@class='x-grid3-col x-grid3-cell x-grid3-td-updateReportIconAction ']//img")).click();
        Thread.sleep(2000);
        if(openNewWindowHandles("Edit Report")) {
            Thread.sleep(2000);
            driver.findElement(By.xpath("//button[@accesskey='o']")).click();
            closeWindowHandle();
            logger.log(LogStatus.PASS, "Test Case is Passed");
        }
        else{
            logger.log(LogStatus.ERROR, "Window Not Found");
        }
    }
    */
    @Test(priority = 5)
    public void ExportReport() throws InterruptedException {
        logger = extent.startTest("Export a Report");
        commonFunctions.viewProductMenuSection("Reports",driver, DemoReportDetails.get(0));
        Thread.sleep(2000);
        try {
            driver.findElement(By.xpath("//div[@class='ux-maximgb-tg-elbow-active ux-maximgb-tg-elbow-plus']")).click();
            Thread.sleep(2000);
            Actions actions = new Actions(driver);
            WebElement elementLocator = driver.findElement(By.xpath("//a[contains(text(),'Change Notice Log')]"));
            actions.contextClick(elementLocator).perform();
            driver.findElement(By.xpath("//span[text()='Export Report']")).click();//export report as zip file to local
            Thread.sleep(2000);
            logger.log(LogStatus.PASS, "Test Case is Passed");
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

