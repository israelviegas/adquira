package adquira;

import java.util.List;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

import adquira.util.Util;

public class Navegacao2Aula {
	

	public static void main(String[] args) throws InterruptedException {
	
	System.setProperty("webdriver.chrome.driver", "C:\\Viegas\\desenvolvimento\\Selenium\\drivers\\chromedriver.exe");
		
		
		WebDriver driver = new ChromeDriver();
		WebDriverWait wait = new WebDriverWait(driver, 10);
		
		// Viegas
       // wait.until(ExpectedConditions.elementToBeClickable(By.id("idSIButton9"))).click();
		// Viegas
		
		
		
		//Abrir o site automation usando uma URL
		driver.get("http://automationpractice.com/index.php");
		driver.manage().timeouts().implicitlyWait(60, TimeUnit.SECONDS);
		driver.manage().window().maximize();
		System.out.println(driver.getTitle());
		
		//clicando no botão "Sing In"
		WebElement SignIn = driver.findElement(By.linkText("Sign in"));
		SignIn.click();
		//Thread.sleep(10000);
		
		//System.out.println(SignIn.getAttribute("title"));
		//System.out.println(SignIn.getText());
		
		//wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("email_create")));
		//wait.until(ExpectedConditions.elementToBeClickable(By.id("email_create"))).click();
		

		
		//preencher o campo de create com um emal e clicar
		driver.findElement(By.id("email_create")).click();
		//Thread.sleep(100);
		driver.findElement(By.id("email_create")).sendKeys("teste1millsz@gmail.com");
		//Thread.sleep(100);
		driver.findElement(By.id("SubmitCreate")).click();
		//Thread.sleep(5000);
		
		//wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("SubmitCreate")));
		//wait.until(ExpectedConditions.elementToBeClickable(By.id("SubmitCreate"))).click();
		//Thread.sleep(5000);
		
		
		//
		//wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//input[@id='id_gender1']")));
		WebElement RadioInline = driver.findElement(By.xpath("//input[@id='id_gender1']"));
		Thread.sleep(100);
		if (!(RadioInline.isSelected())) {	
			RadioInline.click();
		}

		
		Select comboboxDia = new Select(driver.findElement(By.id("days")));
		comboboxDia.selectByIndex(29);
		List<WebElement> dias = comboboxDia.getOptions();
		for(int i=0; i < dias.size(); i++) {
			System.out.println(dias.get(i).getText());
		}
		Select comboboxMes = new Select(driver.findElement(By.id("months")));
		comboboxMes.selectByValue("1");
		List<WebElement> meses = comboboxMes.getOptions();
		for(int i=0; i < meses.size(); i++) {
			System.out.println(meses.get(i).getText());
		}
		Select comboboxAno = new Select(driver.findElement(By.id("years")));
		comboboxAno.selectByVisibleText("2022  ");
		List<WebElement> anos = comboboxAno.getOptions();
		for(int i=0; i < anos.size(); i++) {
			System.out.println(anos.get(i).getText());
		}
		WebElement newsletter = driver.findElement(By.id("newsletter"));
		Thread.sleep(100);
		if (!(newsletter.isSelected())) {	
			newsletter.click();
		}
	
		WebElement uniformoptin = driver.findElement(By.id("uniform-optin"));
		Thread.sleep(100);
		if (!(uniformoptin.isSelected())) {	
			uniformoptin.click();
		}
		
		
	}	


}
