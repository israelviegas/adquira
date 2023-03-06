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
		
		
		 // Criar uma instância do navegador
	    //WebDriver driver = new ChromeDriver();

	    // Abrir o site do Google
	    driver.get("https://www.google.com");

	    // Localizar a caixa de pesquisa do Google
	    driver.findElement(By.name("q")).sendKeys("youtube");

	    // Enviar o formulário de pesquisa
	    driver.findElement(By.name("q")).submit();
	    Thread.sleep(5000);

	    // Localizar o primeiro link resultante da pesquisa
	    driver.findElement(By.cssSelector("h3 > a")).click();
		
		
		// Viegas
		
		
		
		
		
	}	


}
