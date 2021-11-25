package Windchill;

import io.appium.java_client.windows.WindowsDriver;
import org.openqa.selenium.remote.DesiredCapabilities;

import java.net.MalformedURLException;
import java.net.URL;

public class Testing {
    public static void main(String args[]) throws MalformedURLException, InterruptedException {
            DesiredCapabilities sessionForWindchill = new DesiredCapabilities();
            sessionForWindchill.setCapability("deviceName", "BDC7-L-DKCX693");
            sessionForWindchill.setCapability("platformName", "Windows");
            sessionForWindchill.setCapability("app", "C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\Accessories\\notepad.exe");
            WindowsDriver customActionSession = new WindowsDriver(new URL("http://localhost:4723/wd/hub"), sessionForWindchill);
// Control the Notepad app
            customActionSession.findElementByClassName("Edit").sendKeys("This is Input");
            Thread.sleep(5000);
    }
}