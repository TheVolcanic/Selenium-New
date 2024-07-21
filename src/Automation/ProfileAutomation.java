package Automation;

import org.openqa.selenium.PageLoadStrategy;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;

public class ProfileAutomation {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		System.setProperty("webdriver.chrome.driver",
				"\\\\depsuse2wf02\\Profiles\\vjadhav\\Desktop\\chromedriver(1).exe");




				       ChromeOptions
				chromeOptions =
				new ChromeOptions();

				       chromeOptions.setPageLoadStrategy(PageLoadStrategy.NORMAL);




				       String
				chromeProfilePath =
				"C:\\Users\\vinayak\\AppData\\Local\\Google\\Chrome\\User Data";

				       chromeOptions.addArguments("user-data-dir="
				 + chromeProfilePath);

				       chromeOptions.addArguments("--profile-directory=Profile 8");
				       //chromeOptions.addArguments("--remote-allow-origins=*");




				       ChromeDriver driver =
				new ChromeDriver(chromeOptions);




				       // Navigate to URL

				       driver.get("https://outlook.office.com/mail/inbox/id/AAQkADQxMjlkNTA3LTJkNzMtNGI2NS05ZGI1LTY5YTRlNWY2MWZhYgAQALHELfhh71ZFjPRnOEU4yvw%3D");

	}

}
