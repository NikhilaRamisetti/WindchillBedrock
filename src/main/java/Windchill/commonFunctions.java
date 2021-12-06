package Windchill;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;

import java.io.File;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Set;
import java.util.concurrent.TimeUnit;

public class commonFunctions {


    public static void viewProductMenuSection(String section, WebDriver driver, String ProductName){
        driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);
        try{
            driver.findElement(By.xpath("//a[@id='object_main_navigation_nav']")).click();
//        implicitWait(5);
//        driver.findElement(By.xpath("//li[@id='navigatorTabPanel__object_main_navigation']")).click();
            driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
            if(driver.findElements(By.xpath("//img[@class='x-tree-ec-icon x-tree-elbow-plus']")).size() !=0) {
                driver.findElement(By.xpath("//span[text()='"+ProductName+"']")).click();
                driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);
            }
            driver.findElement(By.xpath("//span[text()='"+section+"']")).click();

        }
        catch(Exception e){
            System.out.println(e.getLocalizedMessage());
            //logger.log(LogStatus.ERROR, e.getLocalizedMessage());
        }
    }
    public static boolean openNewWindowHandles(String WindowName,WebDriver driver) {
        String parent = driver.getWindowHandle();
        Set<String> windows = driver.getWindowHandles();
        for (String window : windows) {
            driver.switchTo().window(window);
            if (driver.getTitle().contains(WindowName)) {
                return true;
            }
        }
        return false;
    }
    public static void closeWindowHandle(WebDriver driver, String parent) throws InterruptedException {
        driver.switchTo().window(parent);
        Thread.sleep(2000);
    }
    public static String getScreenshot(WebDriver driver, String screenshotName, String FolderName, String RootFolder) throws Exception {
        String dateName = new SimpleDateFormat("yyyyMMddhhmmss").format(new Date());
        TakesScreenshot ts = (TakesScreenshot) driver;
        File source = ts.getScreenshotAs(OutputType.FILE);
        String destination = RootFolder + FolderName + screenshotName + dateName + ".png";
        File finalDestination = new File(destination);
        FileUtils.copyFile(source, finalDestination);
        return destination;
    }
    /*
    public static void runTest(Runnable code, ExtentTest logger) {
        try {
            code.run();
        } catch(Exception e) {
                System.out.println(e.getLocalizedMessage());
                logger.log(LogStatus.ERROR, e.getLocalizedMessage());
        }
    }
    */
}
