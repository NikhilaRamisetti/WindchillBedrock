package Windchill;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;
import org.apache.commons.io.FileUtils;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.testng.ITestResult;
import org.testng.annotations.*;
import org.openqa.selenium.JavascriptExecutor;
import java.awt.*;
import java.awt.datatransfer.StringSelection;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Set;
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
    public synchronized void initiatereader() throws IOException {
        System.setProperty("webdriver.chrome.driver", Root + "\\chromedriver_win32\\chromedriver.exe");

    }
    ExcelReader credentialsReader= ExcelReader.getInstance(Root + "\\src\\main\\java\\Windchill","TestDataInput.xlsx","Credentials");
    ExcelReader excelReader= ExcelReader.getInstance(Root + "\\src\\main\\java\\Windchill","TestDataInput.xlsx","ReportsManagement");
    WebDriver driver;
    ExtentReports extent;
    ExtentTest logger;
    List<String> DemoReportDetails = excelReader.getRowData(1, 0);

    public static String getScreenshot(WebDriver driver, String screenshotName, String FolderName) throws Exception {
        String dateName = new SimpleDateFormat("yyyyMMddhhmmss").format(new Date());
        TakesScreenshot ts = (TakesScreenshot) driver;
        File source = ts.getScreenshotAs(OutputType.FILE);
        String destination = Root + FolderName + screenshotName + dateName + ".png";
        File finalDestination = new File(destination);
        FileUtils.copyFile(source, finalDestination);
        return destination;
    }

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
    @Test(priority = 1)
    public void LaunchWindchillBrowser() throws InterruptedException, AWTException {
        logger = extent.startTest("Launching Windchill Browser");
        driver.get("http://windchilltest.accenture.com:81/Windchill/app");
        driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
        Thread.sleep(2000);
        Robot rb = new Robot();
        java.util.List<String> cred1= credentialsReader.getRowData(1,0);
        StringSelection str = new StringSelection(cred1.get(0));
        Toolkit.getDefaultToolkit().getSystemClipboard().setContents(str, null);
        rb.keyPress(KeyEvent.VK_CONTROL);rb.keyPress(KeyEvent.VK_V);
        rb.keyRelease(KeyEvent.VK_CONTROL);rb.keyRelease(KeyEvent.VK_V);
        Thread.sleep(2000);
        rb.keyPress(KeyEvent.VK_TAB);
        StringSelection str1 = new StringSelection(cred1.get(1));
        Toolkit.getDefaultToolkit().getSystemClipboard().setContents(str1, null);
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
        viewProductMenuSection("Utilities");
        Thread.sleep(2000);
        try {
            driver.findElement(By.xpath("//a[contains(text(),'Report Management')]")).click();
            Thread.sleep(2000);
            driver.findElement(By.xpath("//button[@style='background-image: url(\"netmarkets/images/report_template_new.png\");']")).click();
            Thread.sleep(2000);
            if (openNewWindowHandles("Query Builder")) {
                implicitWait(10);
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
                closeWindowHandle();
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
            if (openNewWindowHandles("View Change Notice Log")) {
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
                getScreenshot(driver, "Change_Notice_Log", "/Screenshots/");
                Thread.sleep(2000);
                driver.close();
                closeWindowHandle();
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
        viewProductMenuSection("Reports");
        Thread.sleep(2000);
        try {
            driver.findElement(By.xpath("//div[@class='ux-maximgb-tg-elbow-active ux-maximgb-tg-elbow-plus']")).click();
            Thread.sleep(2000);
            Actions actions = new Actions(driver);
            WebElement elementLocator = driver.findElement(By.xpath("//a[contains(text(),'Change Notice Log')]"));
            actions.contextClick(elementLocator).perform();
            driver.findElement(By.xpath("//span[text()='Export Report']")).click();
            Thread.sleep(2000);
            logger.log(LogStatus.PASS, "Test Case is Passed");
        }
        catch(Exception e){
            System.out.println(e.getLocalizedMessage());
            logger.log(LogStatus.ERROR, e.getLocalizedMessage());
        }
    }
    public void viewProductMenuSection(String section){
        implicitWait(15);
        try{
        driver.findElement(By.xpath("//a[@id='object_main_navigation_nav']")).click();
//        implicitWait(5);
//        driver.findElement(By.xpath("//li[@id='navigatorTabPanel__object_main_navigation']")).click();
        implicitWait(5);
        //driver.findElement(By.id("ext-gen169")).click();
        implicitWait(5);
        if(driver.findElements(By.xpath("//img[@class='x-tree-ec-icon x-tree-elbow-plus']")).size() !=0) {
            driver.findElement(By.xpath("//span[text()='"+DemoReportDetails.get(0)+"']")).click();
            implicitWait(5);
        }
        driver.findElement(By.xpath("//span[text()='"+section+"']")).click();

        }
        catch(Exception e){
            System.out.println(e.getLocalizedMessage());
            logger.log(LogStatus.ERROR, e.getLocalizedMessage());
        }
    }
    public boolean openNewWindowHandles(String WindowName) {
        parent = driver.getWindowHandle();
        Set<String> windows = driver.getWindowHandles();
        for (String window : windows) {
            driver.switchTo().window(window);
            System.out.println(WindowName);
            System.out.println(driver.getTitle());
            System.out.println(driver.getTitle().contains(WindowName));
            if (driver.getTitle().contains(WindowName)) {
                return true;
            }
        }
        return false;
    }
    public void closeWindowHandle() throws InterruptedException {
        driver.switchTo().window(parent);
        Thread.sleep(2000);
        }
    public void implicitWait(int seconds){
        driver.manage().timeouts().implicitlyWait(seconds, TimeUnit.SECONDS);
    }

    @AfterTest
    public void closebrowser() {
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
            String screenshotPath = CommonAppsDriver.getScreenshot(driver, result.getName(), "/FailedTestsScreenshots/");
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

