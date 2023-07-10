package tests;

import java.util.concurrent.TimeUnit;

import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.Assert;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import io.github.bonigarcia.wdm.WebDriverManager;

public class DataDrivenTest {

	public static String path = "E:\\New folder\\DataDriven.xlsx";
	WebDriver driver;

	@BeforeTest
	public void initialization() {

		// To set the path of the Chrome driver.
		WebDriverManager.chromedriver().setup();
		driver = new ChromeDriver();
		// To launch application
		driver.get("https://www.saucedemo.com/");
		// To maximize the browser
		driver.manage().window().maximize();
		// implicit wait
		driver.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);
	}

	@Test
	public void saucePageLogin() throws Exception {
	    XLUtils.setExcelFile(path, "test steps");
	    int rows = XLUtils.getRowCount(path, "test steps");
	    for (int i = 1; i <= rows; i++) {
	        String username = XLUtils.getCellData(path, "test steps", i, 0);
	        String pwd = XLUtils.getCellData(path, "test steps", i, 1);
	        Thread.sleep(5000);
	        
	        // Check if username and password are not empty or null
	        if (username != null && !username.isEmpty() && pwd != null && !pwd.isEmpty()) {
	            // To read page title
	            String pgTitle = driver.getTitle();
	            if (pgTitle.equals("Swag Labs")) {
	                driver.findElement(By.name("user-name")).sendKeys(username);
	                driver.findElement(By.name("password")).sendKeys(pwd);
	                driver.findElement(By.name("login-button")).click();
	                System.out.println("Test passed");
	                Thread.sleep(5000);
	                XLUtils.setCellData(path, "test steps", i, 2, "successful login operation");
	            } else {
	                System.out.println("Test Failed");
	                XLUtils.setCellData(path, "test steps", i, 2, "Unsuccessful login operation");
	            }
	            
	            driver.findElement(By.xpath("//*[@class ='bm-burger-button']")).click();
	            driver.findElement(By.linkText("Logout")).click();
	        } else {
	        	XLUtils.setCellData(path, "test steps", i, 2, "Invalid username or password");
	        }
	    }
	}
	
	@AfterTest
	public void tearDown() {
		driver.close();
	}
}
