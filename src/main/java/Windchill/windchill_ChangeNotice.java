package Windchill;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;
import drivers.CommonAppsDriver;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
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

public class windchill_ChangeNotice {
    private static File directory = new File("");
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

        System.setProperty("webdriver.chrome.driver", "C:\\Program Files\\Accenture\\IX0\\Bedrock\\SetupFiles\\chromedriver\\chromedriver_win32\\chromedriver.exe");

    }


    WebDriver driver;
    ExtentTest logger;
    ExtentReports extent;
    String patternPath = System.getProperty("user.dir") + "\\PixelGraphics\\";
    ArrayList<String> list1;


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

        driver.get("http://windchilltest.accenture.com:82/Windchill/app");
        Thread.sleep(2000);
        Robot rb = new Robot();
        ExcelReader credentialsReader = ExcelReader.getInstance(System.getProperty("user.dir") + "/GlobalSettings", Root + "\\src\\main\\java\\Windchill", "TestDataInput.xlsx", "Credentials");
        List<String> excelData = credentialsReader.getRowData(3, 0);
        String USERNAME = excelData.get(0);
        String PASSWORD = excelData.get(1);
        StringSelection str = new StringSelection(USERNAME);
        Toolkit.getDefaultToolkit().getSystemClipboard().setContents(str, null);
        rb.keyPress(KeyEvent.VK_CONTROL);
        rb.keyPress(KeyEvent.VK_V); // press Contol+V to paste
        rb.keyRelease(KeyEvent.VK_CONTROL);
        rb.keyRelease(KeyEvent.VK_V);// release Contol+V to paste
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

    @Test(priority = 1)
    public void CreateDoc() throws InterruptedException {
        logger = extent.startTest("CreateDoc");

        try {

            WebDriverWait wait = new WebDriverWait(driver, 30);
            driver.findElement(By.xpath("//div[@class='x-tool x-tool-expand-west']")).click();
            Thread.sleep(2000);
            driver.findElement(By.id("navigatorTabPanel__object_main_navigation")).click();
            Thread.sleep(3000);
            driver.findElement(By.xpath("//img[@class='x-tree-ec-icon x-tree-elbow-plus']")).click();//product selection
            Thread.sleep(2000);
            driver.findElement(By.xpath("//span[text()='Folders']")).click();//folder selection
            Thread.sleep(4000);
            Actions actions = new Actions(driver);
            WebElement ele = driver.findElement(By.xpath("//div[@class='x-grid3-cell-inner x-grid3-col-number']"));
            actions.contextClick(ele).perform();
            Thread.sleep(2000);
            driver.findElement(By.xpath("//span[text()='Copy']")).click();
            Thread.sleep(2000);
            WebElement element = driver.findElement(By.xpath("//div[@class='x-grid3-cell-inner x-grid3-col-name']"));
            actions.contextClick(element).perform();
            Thread.sleep(2000);
            driver.findElement(By.xpath("//span[text()='New']")).click();//windchill actions > new >new change notice.
            driver.findElement(By.xpath("//span[text()='New Change Notice']")).click();
            Set<String> winHandles1 = driver.getWindowHandles();
            list1 = new ArrayList<>(winHandles1);
            driver.switchTo().window(list1.get(1));
            Thread.sleep(4000);
            logger.log(LogStatus.PASS, "Test Case is Passed");

        } catch (Exception e) {
            System.out.println(e.getLocalizedMessage());
            logger.log(LogStatus.ERROR, e.getLocalizedMessage());
        }


    }


    @Test(priority = 2)
    public void newCN() throws InterruptedException {
        logger = extent.startTest("NewChnageNotice");

        try {
            Select template = new Select(driver.findElement(By.id("ChangeObjectTemplatePicker")));//change notice template
            template.selectByIndex(0);
            Thread.sleep(4000);
            ExcelReader credentialsReader = ExcelReader.getInstance(System.getProperty("user.dir") + "/GlobalSettings", Root + "\\src\\main\\java\\Windchill", "TestDataInput.xlsx", "ChangeNotice");
            List<String> excelData = credentialsReader.getRowData(1, 0);
            System.out.println(excelData);
            String CNname = excelData.get(0);
            driver.findElement(By.xpath("//td[@attrid='name']/input[1]")).sendKeys(CNname);//CN name
            Thread.sleep(3000);
            Select complexity = new Select(driver.findElement(By.xpath("//select[@name='tcomp$attributesTable$VR:wt.part.WTPart:106324$___theChangeNoticeComplexity___combobox']")));//select dropdown complexity
            complexity.selectByVisibleText("Basic");
            Thread.sleep(2000);
            driver.findElement(By.xpath("//button[@accesskey='n']")).click();//new CN window details.
            logger.log(LogStatus.PASS, "Test Case is Passed");

        } catch (Exception e) {
            System.out.println(e.getLocalizedMessage());
            logger.log(LogStatus.ERROR, e.getLocalizedMessage());
        }
    }

    @Test(priority = 3)
    public void ChangeTask() throws InterruptedException {
        logger = extent.startTest("ChangeTask");

        try {
            driver.findElement(By.cssSelector("button[style='background-image: url(\"netmarkets/images/task_create.gif\");']"));//change task creation
            ExcelReader credentialsReader = ExcelReader.getInstance(System.getProperty("user.dir") + "/GlobalSettings", Root + "\\src\\main\\java\\Windchill", "TestDataInput.xlsx", "ChangeNotice");
            List<String> excelData = credentialsReader.getRowData(1, 0);
            String CTname = excelData.get(1);
            String approver = excelData.get(2);
            String reviewer = excelData.get(3);
            driver.findElement(By.xpath("//button[@class=' x-btn-text blist']")).click();
            Thread.sleep(3000);
            WebElement frame = driver.findElement(By.xpath("//iframe[@src='templates/htmlcomp/Empty.html']"));//switch to frame
            driver.switchTo().frame(frame);
            driver.findElement(By.xpath("//td[@attrid='name']/input[1]")).sendKeys(CTname);//send CT name from the excel
            Thread.sleep(1000);
            driver.findElement(By.xpath("//td[@attrid='ROLE_ASSIGNEE']/span/div/input[@type='text']")).clear();//Erase already existing data
            Thread.sleep(2000);
            driver.findElement(By.xpath("//td[@attrid='ROLE_ASSIGNEE']/span/div/input[@type='text']")).sendKeys(approver);//send approver name from the excel
            Thread.sleep(2500);

            driver.findElement(By.xpath("//td[@attrid='ROLE_REVIEWER']/span/div/input[@type='text']")).clear();//Erase alraedy eisting data
            Thread.sleep(2000);
            driver.findElement(By.xpath("//td[@attrid='ROLE_REVIEWER']/span/div/input[@type='text']")).sendKeys(reviewer);//send reviwer name from excel
            Thread.sleep(2000);
            driver.findElement(By.xpath("//button[@accesskey='n']")).click();//click on the next button
            Thread.sleep(2000);
            logger.log(LogStatus.PASS, "Test Case is Passed");

        } catch (Exception e) {
            System.out.println(e.getLocalizedMessage());
            logger.log(LogStatus.ERROR, e.getLocalizedMessage());
        }
    }

    @Test(priority = 4)
    public void AffectedObjects() throws InterruptedException {
        logger = extent.startTest("AffectedObjects");

        try {
            driver.findElement(By.id("ext-gen88")).click();//paste the part number
            Thread.sleep(2000);
            driver.findElement(By.id("ext-gen179")).click();//paste the part number
            Thread.sleep(3000);
            driver.findElement(By.xpath("//button[@accesskey='f']")).click();//click on the finish button
            Thread.sleep(2000);
            driver.switchTo().defaultContent();
            logger.log(LogStatus.PASS, "Test Case is Passed");

        } catch (Exception e) {
            System.out.println(e.getLocalizedMessage());
            logger.log(LogStatus.ERROR, e.getLocalizedMessage());
        }
    }

    @Test(priority = 5)
    public void ImplementaionPlan() throws InterruptedException {
        logger = extent.startTest("ImplementationPlan");

        try {
            driver.findElement(By.xpath("//button[@accesskey='n']")).click();//click on the next button
            Thread.sleep(2000);
            driver.findElement(By.xpath("//button[@accesskey='f']")).click();//click on the finish button
            driver.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);
            driver.findElement(By.xpath("//b[text()='Submit Now']")).click();
            Thread.sleep(5000);
            driver.switchTo().window(list1.get(0));
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


