package Automation;

import java.util.concurrent.TimeUnit;

//TODO Auto-generated method stub
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;



public class OutlookAutomation {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		// Set the path to the chromedriver executable
        //System.setProperty("webdriver.chrome.driver", "C:\\Users\\Vinayak\\Desktop\\chromedriver.exe");
		//System.setProperty("webdriver.gecko.driver","C:\\Users\\Vinayak\\Desktop\\geckodriver.exe");
		DesiredCapabilities capabilities = DesiredCapabilities.firefox();  
        capabilities.setCapability("marionette",true);  // Setting system properties of FirefoxDriver
				WebDriver driver = new FirefoxDriver(capabilities);
		//System.setProperty("webdriver.edge.driver", "C:\\Users\\Vinayak\\Desktop\\msedgedriver.exe");

        // Initialize ChromeDriver
        //WebDriver driver = new EdgeDriver();
        
        driver.manage().window().maximize();
        
        // Open Outlook web version
        //driver.get("https://outlook.live.com/mail/0/");
        driver.get("https://login.live.com/ppsecure/post.srf?cobrandid=ab0455a0-8d03-46b9-b18b-df2f57b9e44c&id=292841&contextid=7B472805B4902BCC&opid=82341CEE11DF31E7&bk=1709457876&uaid=4b8866ed61194df688d5837c48fa0e0f&pid=0"); 
        
        try {
			Thread.sleep(10000);
		} catch (InterruptedException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
       
        
      WebElement buttonN= driver.findElement(By.xpath("//a[contains(text(),'Sign in')]"));
     
        buttonN.click();
        WebDriverWait wait = new WebDriverWait(driver, 1);
        driver.manage().timeouts().implicitlyWait(120,TimeUnit.SECONDS);
        wait.until(ExpectedConditions.titleContains("Microsoft Outlook Personal Email and Calendar | Microsoft 365"));
        try {
			Thread.sleep(30000);
		} catch (InterruptedException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
       
      /*  WebElement Username= driver.findElement(By.id("i0116"));
        Username.sendKeys("vjadhav48@outlook.com");
        try {
			Thread.sleep(10000);
		} catch (InterruptedException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
        WebElement Pass= driver.findElement(By.id("i0118"));
        Username.sendKeys("Selenium@2024");
        
        WebElement buttonN1= driver.findElement(By.id("acceptButton"));
        buttonN1.click();*/
      //*[@id="i0118"]
      //*[@id="i0116"]
        // Log in to Outlook (assuming manual login)
        // You may need to handle login programmatically if possible

        // Wait for the inbox to load
        WebDriverWait wait1 = new WebDriverWait(driver, 1);
        driver.manage().timeouts().implicitlyWait(120,TimeUnit.SECONDS);
        wait1.until(ExpectedConditions.titleContains("Mail - Vinayak Jadhav - Outlook"));

        // Locate and open the most recent email
        WebElement latestEmail = driver.findElement(By.xpath("/html/body/div[1]/div/div[2]/div/div[2]/div[2]/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div/div/div/div/div/div[2]/div/div/div/div/div[2]/div[2]/div[3]/div/div/span"));
        latestEmail.click();
        
        

        // Wait for the email to load
        wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("/html/body/div[1]/div/div[2]/div/div[2]/div[2]/div/div/div/div[3]/div/div/div[3]/div/div/div/div/div/div[2]/div/div/div[1]/div/div/div/div")));
      // JavascriptExecutor js = (JavascriptExecutor) driver;
      
       //WebElement content = driver.findElement(By.xpath("/html/body/div[1]/div/div[2]/div/div[2]/div[2]/div/div/div/div[2]/div/div/div[3]/div/div/div/div/div/div[2]/div/div/div[1]/div/div/div/div/div[1]/div[1]/div[1]/div[2]/span[2]/div"));
        //content.click();

        // Scroll down the page by a specified number of pixels
       //js.executeScript("window.scrollBy(0, 20000);");
        //Actions actions = new Actions(driver);

        // Scroll down the page
        //actions.moveToElement(driver.findElement(By.xpath("/html/body/div[1]/div/div[2]/div/div[2]/div[2]/div/div/div/div[3]/div/div/div[3]/div/div/div/div/div/div[2]/div"))).clickAndHold().moveByOffset(0, 10000).release().perform();

        // Find the button within the email and click it
       WebElement button= driver.findElement(By.xpath("//a[contains(text(),'Access Assigned Activities')]"));
       //WebElement button= driver.findElement(By.xpath("//a[@href='https://u7177183.ct.sendgrid.net/ls/click?upn=u001.pczFKVg0Srke7du83-2FSGPzLdFSeLJv01BXVWdZhD2VLhfM8w7YIYPdJVXrfp7izQSizpTcfYX074K8dQfpa3r10GHgriDQiqpopIpW2dOx6Q0hZC-2BPcU0DmLCFHThIbjIe21and4WN7NgQj9b4awMWmpxOhgcSAXvGNYOlK5w-2BqISEZwN2ql9r-2FKAM1WdC5-2Fruz7_r-2BkTEWmzr-2Bvp1DbC9VuTOm0vjoaBDb-2BuCpunSWhQj2u8PQr8y4-2FjK57v2-2B2qwOj7Rh95Sk5rt198muZT0iN03QsZEM4H091CV-2Fm7EMtF39-2B2W36UyCLnVevqVJAMhPCx-2BN2HgnzasY-2BNm2JW6MWwCpzscAf9l8ZbHW9gubwzEGBAD3AocGsCQlbEJ479y0KXKA8-2BagSwkCGHOcwkz9JCKf8TduRk6Tnu46V4kXSKBuA-3D']"));
      //a[@class='button-class'
        //color:#4a54a4
        //WebElement button= driver.findElement(By.xpath("//a[contains(target(),'_blank')]"));
        Actions actions = new Actions(driver);

        // Perform mouse click on the button
        actions.click(button).perform();
//      //<a href="https://u7177183.ct.sendgrid.net/ls/click?upn=u001.pczFKVg0Srke7du83-2FSGPzLdFSeLJv01BXVWdZhD2VLhfM8w7YIYPdJVXrfp7izQSizpTcfYX074K8dQfpa3r10GHgriDQiqpopIpW2dOx6Q0hZC-2BPcU0DmLCFHThIbjIe21and4WN7NgQj9b4awMUKAOZwVDqSCe15aKoZZptRNHhX8yL4euRcnXQkO15Q0QO5m_r-2BkTEWmzr-2Bvp1DbC9VuTOm0vjoaBDb-2BuCpunSWhQj2vvedqQTukoqkRMmmzAlHmpFWmdDbIeAtQEtvcb-2BpzhNuNMG-2BswtmBojCCsCUOwfW-2BxwHpTfP8JZkTpzS23t8cK2v-2BIwfKqd4c4yjqO5TDzt9ulxOinLwGMLkpt4CKyz-2BJpGVFR5eGdWBvPNjtWfDxPePZnOeew9KUk-2B3RCak7vhI1xtKWlOrPXTJXNANuHOJk-3D" target="_blank" rel="noopener noreferrer" data-auth="NotApplicable" style="text-decoration:none; line-height:42px; display:block; color:#4a54a4" data-linkindex="3">Access Assigned Activities </a>
        
        //js.executeScript("arguments[0].scrollIntoView();", button);
        //buttonLink.click();
//https://u7177183.ct.sendgrid.net/ls/click?upn=u001.pczFKVg0Srke7du83-2FSGPzLdFSeLJv01BXVWdZhD2VLhfM8w7YIYPdJVXrfp7izQSizpTcfYX074K8dQfpa3r10GHgriDQiqpopIpW2dOx6Q0hZC-2BPcU0DmLCFHThIbjIe21and4WN7NgQj9b4awMWmpxOhgcSAXvGNYOlK5w-2BqISEZwN2ql9r-2FKAM1WdC5-2Fruz7_r-2BkTEWmzr-2Bvp1DbC9VuTOm0vjoaBDb-2BuCpunSWhQj2u8PQr8y4-2FjK57v2-2B2qwOj7Rh95Sk5rt198muZT0iN03QsZEM4H091CV-2Fm7EMtF39-2B2W36UyCLnVevqVJAMhPCx-2BN2HgnzasY-2BNm2JW6MWwCpzscAf9l8ZbHW9gubwzEGBAD3AocGsCQlbEJ479y0KXKA8-2BagSwkCGHOcwkz9JCKf8TduRk6Tnu46V4kXSKBuA-3D
        // Close the browser
        //driver.quit();
    }
}


	


