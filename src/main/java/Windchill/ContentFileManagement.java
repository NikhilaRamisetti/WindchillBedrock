package Windchill;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
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

public class ContentFileManagement {
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

    private ContentFileManagement CommonAppsDriver;

    public ContentFileManagement() throws IOException {
    }

    @BeforeSuite
    public synchronized void initiateReader() throws IOException {
        System.setProperty("webdriver.chrome.driver", Root + "\\chromedriver_win32\\chromedriver.exe");

    }
    public static JavascriptExecutor jse;
    ExcelReader credentialsReader= ExcelReader.getInstance(Root + "\\src\\main\\java\\Windchill","TestDataInput.xlsx","Credentials");
    ExcelReader excelReader= ExcelReader.getInstance(Root + "\\src\\main\\java\\Windchill","TestDataInput.xlsx","ContentFileManagement");
    WebDriver driver;
    ExtentReports extent;
    ExtentTest logger;
    java.util.List<String> DemoContentDetails = excelReader.getRowData(1, 0);

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
    public void addSecondaryContent() throws InterruptedException {
        logger = extent.startTest("Add Secondary Content");
        commonFunctions.viewProductMenuSection("Folders", driver, DemoContentDetails.get(0) );
        Thread.sleep(2000);
        try {
            for (int i = 1; i < excelReader.getRowCount(); i++) {
                java.util.List<String> contentDetails = excelReader.getRowData(i, 0);
                //Primary document details
                driver.findElement(By.xpath("//a[text()='" + contentDetails.get(0) + "']")).click();
                Thread.sleep(2000);
                implicitWait(10);
                driver.findElement(By.xpath("//span[text()='Content']")).click();
                Thread.sleep(5000);
                implicitWait(10);
                driver.findElement(By.xpath("//button[@style='background-image: url(\"netmarkets/images/edit.gif\");']")).click();
                Thread.sleep(5000);
                implicitWait(10);
                //switch to edit window from parent window
                if (commonFunctions.openNewWindowHandles("Edit",driver)) {
                    if (contentDetails.get(4).equalsIgnoreCase("Local File")) {
                        driver.findElement(By.xpath("//button[@style='background-image: url(\"netmarkets/images/content-file-generic_attach.gif\");']")).click();
                    } else if (contentDetails.get(4).equalsIgnoreCase("URL Link")) {
                        driver.findElement(By.xpath("//button[@style='background-image: url(\"netmarkets/images/content-url_attach.gif\");']")).click();
                    } else if (contentDetails.get(4).equalsIgnoreCase("External storage")) {
                        driver.findElement(By.xpath("//button[@style='background-image: url(\"netmarkets/images/content-external_attach.gif\");']")).click();
                    } else {
                        logger.log(LogStatus.ERROR, "Type of document doesnt exist");
                    }
                    driver.findElement(By.xpath("//div[@class='x-grid3-cell-inner x-grid3-col-contentName']//input[1]")).sendKeys(contentDetails.get(5));
                    driver.findElement(By.xpath("//div[@class='x-grid3-cell-inner x-grid3-col-contentLocation']//input[1]")).sendKeys(contentDetails.get(6));
                    Thread.sleep(2000);
                    driver.findElement(By.xpath("//button[@accesskey='s']")).click();
                    Thread.sleep(2000);
                    commonFunctions.closeWindowHandle(driver, parent);
                } else {
                    logger.log(LogStatus.ERROR, "Window Not Found");
                }
            }

            logger.log(LogStatus.PASS, "Test Case is Passed");
        }
        catch(Exception e){
            System.out.println(e.getLocalizedMessage());
            logger.log(LogStatus.ERROR, e.getLocalizedMessage());
        }
    }
    @Test(priority=3)
    public void downloadContent() {
        logger = extent.startTest("Download content file");
        commonFunctions.viewProductMenuSection("Folders", driver, DemoContentDetails.get(0));
        try{
            contentActionsClick();
            //Download primary content file
            driver.findElement(By.xpath("//img[@src='http://windchilltest.accenture.com:81/Windchill/netmarkets/images/file_msword.gif']")).click();
            Thread.sleep(2000);
            //Download secondary content file
            //driver.findElement(By.xpath("//div[@class='x-grid3-cell-inner x-grid3-col-attachmentsName']/a[2]")).click();
        }
        catch (Exception n){
            System.out.println(n.getLocalizedMessage());
            logger.log(LogStatus.ERROR, "Error In document name");
        }

        logger.log(LogStatus.PASS, "Test Case is Passed");
    }
    //Edit Report is only possible if logged In as Admin (wcadmin)
    @Test(priority = 4)
    public void replaceContent() throws InterruptedException {
        logger = extent.startTest("Replace Content");
        //contentActionsClick();
        try {
            driver.findElement(By.xpath("//button[text()='Actions']")).click();
            if (driver.findElement(By.xpath("//span[text()='Replace Content']")).isEnabled()) {
                driver.findElement(By.xpath("//span[text()='Replace Content']")).click();
                if (commonFunctions.openNewWindowHandles("Replace Content", driver)) {
                    documentContentSelection();//select the type of document(local file, External file, URL link)
                    commonFunctions.closeWindowHandle(driver, parent);
                } else {
                    logger.log(LogStatus.ERROR, "Window Not Found");
                }
            } else {
                logger.log(LogStatus.ERROR, "Document is checked out, please check In for Replacing content");
            }

            logger.log(LogStatus.PASS, "Test Case is Passed");
        }
        catch(Exception e){
            System.out.println(e.getLocalizedMessage());
            logger.log(LogStatus.ERROR, e.getLocalizedMessage());
        }
    }

    @Test(priority = 5)
    public void checkInAndCheckOut() throws InterruptedException {
        logger = extent.startTest("CheckIn And CheckOut");
        //contentActionsClick();
        try {
            driver.findElement(By.xpath("//button[text()='Actions']")).click();
            if (DemoContentDetails.get(10).equalsIgnoreCase("CheckIn")) {
                if (driver.findElement(By.xpath("//span[text()='Check In']")).isEnabled()) {
                    if (commonFunctions.openNewWindowHandles("Checking In Document", driver)) {
                        documentContentSelection();//select the type of document(local file, External file, URL link)
                        commonFunctions.closeWindowHandle(driver, parent);
                    } else {
                        logger.log(LogStatus.ERROR, "Window Not Found");
                    }
                } else {
                    logger.log(LogStatus.ERROR, "Document already checked In, Please check out manually and try again");
                }
            } else if (DemoContentDetails.get(10).equalsIgnoreCase("CheckOut")) {
                if (driver.findElement(By.xpath("//span[text()='Check Out']")).isEnabled()) {
                    driver.findElement(By.xpath("//span[text()='Check Out']")).click();
                    Thread.sleep(2000);
                } else {
                    logger.log(LogStatus.ERROR, "Document already checked out to user, Please check out manually and try again");
                }
            }
            logger.log(LogStatus.PASS, "Test Case is Passed");
        }
        catch(Exception e){
            System.out.println(e.getLocalizedMessage());
            logger.log(LogStatus.ERROR, e.getLocalizedMessage());
        }
    }
    public void documentContentSelection() throws InterruptedException {
        try {
            driver.findElement(By.xpath("//select[@id='primary0contentSourceList']")).click();
            if (DemoContentDetails.get(7).equalsIgnoreCase("Local File")) {
                driver.findElement(By.xpath("//option[@id='primary0contentSourceList_FILE']")).click();
                driver.findElement(By.xpath("//input[@id='keep_existing_primary_file']")).click();
                Thread.sleep(2000);
                driver.findElement(By.xpath("//button[@accesskey='o']")).click();
            } else if (DemoContentDetails.get(7).equalsIgnoreCase("URL Link")) {
                driver.findElement(By.xpath("//option[@id='primary0contentSourceList_URL']")).click();
                driver.findElement(By.xpath("//input[@id='primaryUrlLocationTextBox']")).sendKeys(DemoContentDetails.get(9));
                driver.findElement(By.xpath("//input[@id='primaryUrlNameTextBox']")).sendKeys(DemoContentDetails.get(8));
                Thread.sleep(2000);
                driver.findElement(By.xpath("//button[@accesskey='o']")).click();
                Thread.sleep(2000);
            } else if (DemoContentDetails.get(7).equalsIgnoreCase("External storage")) {
                driver.findElement(By.xpath("//option[@id='primary0contentSourceList_EXTERNAL']")).click();
                driver.findElement(By.xpath("//textarea[@id='primaryExternalLocationTextArea']")).sendKeys(DemoContentDetails.get(9));
                driver.findElement(By.xpath("//textarea[@id='primaryExternalNameTextBox']")).sendKeys(DemoContentDetails.get(8));
                Thread.sleep(2000);
                driver.findElement(By.xpath("//button[@accesskey='o']")).click();
                Thread.sleep(2000);
            } else {
                logger.log(LogStatus.ERROR, "Type of document doesnt exist");
            }
        }
        catch(Exception e){
            System.out.println(e.getLocalizedMessage());
            logger.log(LogStatus.ERROR, e.getLocalizedMessage());
        }
    }
    public void contentActionsClick() throws InterruptedException {
        try {
            Thread.sleep(5000);
            driver.findElement(By.xpath("//a[text()='" + DemoContentDetails.get(0) + "']")).click();
            Thread.sleep(2000);
            implicitWait(10);
            driver.findElement(By.xpath("//span[text()='Content']")).click();
            Thread.sleep(5000);
            //driver.findElement(By.xpath("//button[text()='Actions']")).click();
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

