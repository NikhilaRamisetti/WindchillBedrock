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
import java.util.ArrayList;
import java.util.Set;
import java.util.concurrent.TimeUnit;

public class windchill_createProd {
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


    @BeforeClass
    public void Prerequisite() {
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
        driver.get("http://windchilltest.accenture.com:81/Windchill/app");
        Thread.sleep(2000);
        Robot rb = new Robot();
        StringSelection str = new StringSelection("testuser1");
        Toolkit.getDefaultToolkit().getSystemClipboard().setContents(str, null);
        rb.keyPress(KeyEvent.VK_CONTROL);
        rb.keyPress(KeyEvent.VK_V); // press Contol+V for pasting
        rb.keyRelease(KeyEvent.VK_CONTROL);
        rb.keyRelease(KeyEvent.VK_V);// release Contol+V for pasting
        Thread.sleep(2000);
        rb.keyPress(KeyEvent.VK_TAB);
        StringSelection str1 = new StringSelection("123");
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

        //span[contains(class, 'class')]
        driver.findElement(By.xpath("//div[@class='x-tool x-tool-expand-west']")).click();
        Thread.sleep(2000);
        driver.findElement(By.id("navigatorTabPanel__object_main_navigation")).click();
        driver.findElement(By.xpath("//span[@class='x-tab-strip-text productNavigation-icon']")).click();
        driver.findElement(By.xpath("//a[text()='View All']")).click();
        // driver.findElement(By.className("//td[@class='x-btn-mc/button']")).click();
        wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ext-gen197"))).click();
        Set<String> winHandles = driver.getWindowHandles();
        ArrayList<String> list = new ArrayList<>(winHandles);
        driver.switchTo().window(list.get(1));
        wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//td[@attrid='containerInfo.name']/input[1]"))).sendKeys("product2");
        Select temp = new Select(driver.findElement(By.name("tcomp$attributesTable$$___templateRef___combobox")));
        Thread.sleep(2000);
        temp.selectByVisibleText("General Product");
        driver.findElement(By.id("containerInfo.description")).sendKeys("product2");

        WebElement yes = driver.findElement(By.name("tcomp$attributesTable$$___pdmlinkContainerPrivateAccess_col_pdmlinkContainerPrivateAccess___com.ptc.core.components.rendering.renderers.RadioButtonGroupRendererGROUP_ID-166259175___radio"));
        yes.click();
        driver.findElement(By.id("ext-gen39")).click();
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


    //driver.findElement(By.xpath("//button[@class='x-btn-text blist']")).click();



//        driver.findElement(By.id("ext-gen113")).click();
//        driver.findElement(By.id("ext-gen50")).click();
//        driver.findElement(By.id("ext-gen152")).click();
//        driver.findElement(By.xpath("//a[contains(text(),'View All')]")).click();
//        wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ext-gen197"))).click();
//        Set<String> winHandles = driver.getWindowHandles();
//        ArrayList<String> list = new ArrayList<>(winHandles);
//        driver.switchTo().window(list.get(1));
//        //driver.findElement(By.xpath("//td[@attrid='name']/input[1]")).sendKeys("DemoPart")
//        wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//td[@attrid='containerInfo.name']/input[1]"))).sendKeys("newproduct1");
//        Select temp=new Select(driver.findElement(By.xpath("//td[@attrid='containerTemplateReference'/select[1]]")));
//        temp.selectByVisibleText("General Product");

}






