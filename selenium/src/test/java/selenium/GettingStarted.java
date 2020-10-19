package selenium;

import org.openqa.selenium.*;

import org.openqa.selenium.chrome.*;

public class GettingStarted {
	public void testGoogleSearch() throws InterruptedException {
		// Optional. If not specified, WebDriver searches the PATH for chromedriver.
		System.setProperty("webdriver.chrome.driver", "D:\\Software\\chromedriver_85.exe");

		WebDriver driver = new ChromeDriver();
		driver.get("http://www.google.com/");
		Thread.sleep(5000); // Let the user actually see something!
		WebElement searchBox = driver.findElement(By.name("q"));
		searchBox.sendKeys("ChromeDriver");
		searchBox.submit();
		Thread.sleep(5000); // Let the user actually see something!
		driver.quit();
	}

	public static void main(String[] args) throws InterruptedException {
		new GettingStarted().testGoogleSearch();
	}
}