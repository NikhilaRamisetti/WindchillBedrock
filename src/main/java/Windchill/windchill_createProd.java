package Windchill;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;
import drivers.CommonAppsDriver;
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

public class windchill_createProd {

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
        System.setProperty("webdriver.chrome.driver", Root + "\\chromedriver_win32\\chromedriver.exe");

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
    public void LaunchWebClient() throws  Exception {
        logger = extent.startTest("LaunchWebClient");
        Thread.sleep(2000);
        driver.get("http://windchilltest.accenture.com:82/Windchill/app");
        ExcelReader credentialsReader= ExcelReader.getInstance(System.getProperty("user.dir") + "/GlobalSettings", Root + "\\src\\main\\java\\Windchill","TestDataInput.xlsx","Credentials");
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
        logger.log(LogStatus.PASS, "Test Case is Passed");
    }


    @Test(priority = 1)
    public void  createProduct() throws InterruptedException {
        WebDriverWait wait = new WebDriverWait(driver, 30);
        logger = extent.startTest("CreateProduct");

        try {


            //span[contains(class, 'class')]
            // wait.until(ExpectedConditions.visibilityOfElementLocated(By.className("x-tool x-tool-close"))).click();
            wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[@class='x-tool x-tool-expand-west']"))).click();
            //driver.findElement(By.xpath("//div[@class='x-tool x-tool-expand-west']")).click();
            Thread.sleep(2000);
            driver.findElement(By.id("navigatorTabPanel__object_main_navigation")).click();
            driver.findElement(By.xpath("//span[@class='x-tab-strip-text productNavigation-icon']")).click();
            driver.findElement(By.xpath("//a[text()='View All']")).click();
            // driver.findElement(By.className("//td[@class='x-btn-mc/button']")).click();
            wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//button[@class=' x-btn-text blist']"))).click();
            Set<String> winHandles = driver.getWindowHandles();
            ArrayList<String> list = new ArrayList<>(winHandles);
            driver.switchTo().window(list.get(1));//switch to different window
            ExcelReader credentialsReader = ExcelReader.getInstance(System.getProperty("user.dir") + "/GlobalSettings", Root + "\\src\\main\\java\\Windchill", "TestDataInput.xlsx", "Product_Creation");
            List<String> excelData = credentialsReader.getRowData(1, 0);
            String ProductName = excelData.get(0);//fetch the data from excel
            String ProductDesrptn = excelData.get(1);
            wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//td[@attrid='containerInfo.name']/input[1]"))).sendKeys(ProductName);//send the product name from excel
            Select temp = new Select(driver.findElement(By.name("tcomp$attributesTable$$___templateRef___combobox")));
            Thread.sleep(2000);
            temp.selectByVisibleText("General Product");//select the template dropdown
            driver.findElement(By.id("containerInfo.description")).sendKeys(ProductDesrptn);//enter the product description from excel

            WebElement yes = driver.findElement(By.name("tcomp$attributesTable$$___pdmlinkContainerPrivateAccess_col_pdmlinkContainerPrivateAccess___com.ptc.core.components.rendering.renderers.RadioButtonGroupRendererGROUP_ID-166259175___radio"));
            yes.click(); //select the dropdown
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






