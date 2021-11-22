package Windchill;


import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;
import org.apache.commons.io.FileUtils;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.ITestResult;
import org.testng.annotations.*;

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
    public synchronized void initiatereader() throws IOException {
        System.setProperty("webdriver.chrome.driver", Root + "\\chromedriver_win32\\chromedriver.exe");

    }

    ExcelReader credentialsReader= ExcelReader.getInstance(Root + "\\src\\main\\java\\Windchill","Credentials.xlsx","Sheet1");
    ExcelReader excelReader= ExcelReader.getInstance(Root + "\\src\\main\\java\\Windchill","TestData.xlsx","Sheet1");
    WebDriver driver;
    ExtentReports extent;
    ExtentTest logger;
    String DefaultPart = "Test Part1";

    
    public static String getScreenshot(WebDriver driver, String screenshotName) throws Exception {
        String dateName = new SimpleDateFormat("yyyyMMddhhmmss").format(new Date());
        TakesScreenshot ts = (TakesScreenshot) driver;
        File source = ts.getScreenshotAs(OutputType.FILE);
        String destination = Root + "/FailedTestsScreenshots/" + screenshotName + dateName + ".png";
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
        List<String> cred1= credentialsReader.getRowData(1,0);
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
        logger = extent.startTest("Creating a Part");
        viewProductFoldersPage();
        implicitWait(10);
        for(int i=1;i<excelReader.getRowCount(); i++) {
            List<String> PartDetails= excelReader.getRowData(i,0);
            String PartName = new String(PartDetails.get(0));
            System.out.println(PartName);
            driver.findElement(By.xpath("//button[@style='background-image: url(\"netmarkets/images/newpart.gif\");']")).click();
            String parent = driver.getWindowHandle();
            Set<String> windows = driver.getWindowHandles();
            System.out.println(windows);
            for (String window : windows) {
                driver.switchTo().window(window);
                System.out.println(driver.switchTo().window(window).getTitle());
                if (driver.getTitle().contains("New Part")) {
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
            }
            driver.switchTo().window(parent);
            Thread.sleep(2000);
        }
        logger.log(LogStatus.PASS, "Test Case is Passed");
    }

    @Test(priority=3)
    public void GetPartInfo() throws InterruptedException {
        logger = extent.startTest("Get Part Information");
        viewProductFoldersPage();
        implicitWait(10);
        Thread.sleep(2000);
        ViewPartInfoPage(DefaultPart);
        System.out.println(driver.findElement(By.xpath("//div[@id='dataStoreGeneral']")).getText());
        implicitWait(2);
        System.out.println("Part Information retrieved successfully");
        logger.log(LogStatus.PASS, "Test Case is Passed");
    }

    @Test(priority = 4)
    public void deletePart() throws InterruptedException {
        logger = extent.startTest("Delete a Part");
        Thread.sleep(2000);
        viewProductFoldersPage();
        implicitWait(10);
        Thread.sleep(2000);
        ViewPartInfoPage(DefaultPart);
        driver.findElement(By.xpath("//button[text()='Actions']")).click();
        Thread.sleep(2000);
        driver.findElement(By.xpath("//span[text()='Delete']")).click();
        Thread.sleep(2000);
        String parent=driver.getWindowHandle();
        Set<String> windows = driver.getWindowHandles();
        System.out.println(windows);
        for (String window : windows) {
            driver.switchTo().window(window);
            System.out.println(driver.switchTo().window(window).getTitle());
            if (driver.getTitle().contains("Delete")) {
                implicitWait(10);
                driver.findElement(By.xpath("//button[@accesskey='o']")).click();
            }
        }
        driver.switchTo().window(parent);
        Thread.sleep(2000);
        System.out.println("Part deleted successfully");
        logger.log(LogStatus.PASS, "Test Case is Passed");
    }

    public void viewProductFoldersPage(){
        implicitWait(15);
        driver.findElement(By.xpath("//a[@id='object_main_navigation_nav']")).click();
//        implicitWait(5);
//        driver.findElement(By.xpath("//li[@id='navigatorTabPanel__object_main_navigation']")).click();
        implicitWait(5);
        //driver.findElement(By.id("ext-gen169")).click();
        implicitWait(5);
        if(driver.findElements(By.xpath("//img[@class='x-tree-ec-icon x-tree-elbow-plus']")).size() !=0) {
            driver.findElement(By.xpath("//span[text()='Sample Product1']")).click();
            implicitWait(5);
        }
        driver.findElement(By.xpath("//span[text()='Folders']")).click();
        implicitWait(20);
    }

    public void ViewPartInfoPage(String Part){
        String partName = Part;
        driver.findElement(By.xpath("//a[text()='" + partName + "']")).click();
        implicitWait(5);
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
            String screenshotPath = CommonAppsDriver.getScreenshot(driver, result.getName());
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

