package adquira;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.net.URL;
import java.util.Collections;
import java.util.HashMap;
import java.util.Map;
import java.util.Properties;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.firefox.FirefoxOptions;
import org.openqa.selenium.firefox.FirefoxProfile;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.ie.InternetExplorerOptions;
import org.openqa.selenium.remote.Command;
import org.openqa.selenium.remote.CommandExecutor;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.remote.HttpCommandExecutor;
import org.openqa.selenium.remote.RemoteWebDriver;
import org.openqa.selenium.remote.Response;
import org.openqa.selenium.remote.SessionId;
import org.openqa.selenium.remote.codec.w3c.W3CHttpResponseCodec;
import org.openqa.selenium.remote.http.W3CHttpCommandCodec;

public class Teste4 {
	
	public static void main(String[] args) throws InterruptedException, IOException {
		
		
		//WebDriver driver = null;
		//driver = getWebDriver();
		
		File file = new File(getValor("driver.Chrome.selenium"));
		System.setProperty(getValor("propriedade.sistema.para.driver.Chrome.selenium"), file.getAbsolutePath());
	    DesiredCapabilities caps = DesiredCapabilities.chrome();
	    caps.setJavascriptEnabled(true);
	    caps.setCapability("ignoreZoomSetting", true);
	    caps.setCapability("nativeEvents",false);
	    ChromeOptions chromeOptions = new ChromeOptions(); 
	    Map<String, Object> chromePreferences = new HashMap<String, Object>();
		chromePreferences.put("profile.default_content_settings.popups", 0);
	    chromePreferences.put("download.default_directory",getValor("caminho.download.relatorios"));
	    chromePreferences.put("browser.helperApps.neverAsk.saveToDisk", "text/plain, application/vnd.ms-excel, text/csv, text/comma-separated-values, application/octet-stream");
	    
	    
	    chromeOptions.setExperimentalOption("prefs", chromePreferences);
	    chromeOptions.addArguments("user-data-dir=" + getValor("caminho.dados.usuario.Chrome"));
	    // C:\Users\israel.viegas\AppData\Local\Google\Chrome\User Data
	    chromeOptions.addArguments("--lang=pt");
	    
	    // browserOpt.add_argument("user-data-dir=C:/Users/rumenik.andrade/AppData/Local/Google/Chrome/User Data")
		 
		
		ChromeDriver driver = new ChromeDriver(chromeOptions);
		HttpCommandExecutor executor = (HttpCommandExecutor) driver.getCommandExecutor();
		URL url = executor.getAddressOfRemoteServer();
		SessionId session_id = driver.getSessionId();
		 
		  
		RemoteWebDriver driver2 = createDriverFromSession(session_id, url);
		  
		//driver.get(getValor("url.sharepoint"));
		driver2.get(getValor("url.sharepoint"));
		
        //fazerLogout(wait);
		if (driver2 != null) {
			//driver2.quit();
		}
		
	}
	
	public class AttachedWebDriver extends RemoteWebDriver {

	    public AttachedWebDriver(URL url, String sessionId) {
	        super();
	        setSessionId(sessionId);
	        setCommandExecutor(new HttpCommandExecutor(url) {
	            @Override
	            public Response execute(Command command) throws IOException {
	                if (command.getName() != "newSession") {
	                    return super.execute(command);
	                }
	                return super.execute(new Command(getSessionId(), "getCapabilities"));
	            }
	        });
	        startSession(new DesiredCapabilities());
	    }
	}
	
	
	public static RemoteWebDriver createDriverFromSession(final SessionId sessionId, URL command_executor){
	    CommandExecutor executor = new HttpCommandExecutor(command_executor) {

	    @Override
	    public Response execute(Command command) throws IOException {
	        Response response = null;
	        if (command.getName() == "newSession") {
	            response = new Response();
	            response.setSessionId(sessionId.toString());
	            response.setStatus(0);
	            response.setValue(Collections.<String, String>emptyMap());

	            try {
	                Field commandCodec = null;
	                commandCodec = this.getClass().getSuperclass().getDeclaredField("commandCodec");
	                commandCodec.setAccessible(true);
	                commandCodec.set(this, new W3CHttpCommandCodec());

	                Field responseCodec = null;
	                responseCodec = this.getClass().getSuperclass().getDeclaredField("responseCodec");
	                responseCodec.setAccessible(true);
	                responseCodec.set(this, new W3CHttpResponseCodec());
	            } catch (NoSuchFieldException e) {
	                e.printStackTrace();
	            } catch (IllegalAccessException e) {
	                e.printStackTrace();
	            }

	        } else {
	            response = super.execute(command);
	        }
	        return response;
	    }
	    };

	    return new RemoteWebDriver(executor, new DesiredCapabilities());
	}
	
    // Arquivo de properties sendo usado fora do projeto
    public static Properties getProp() throws IOException {
        Properties props = new Properties();
       
        
        // String arquivoProperties = "D:/JOBS/Automacao Adquira/configuracoes/gerar relatorio pedidos.properties"; 
        // String arquivoProperties = "C:/Automacao Adquira/configuracoes/gerar relatorio pedidos.properties"; 
        // Quando a integração do Sharepoibt estiver pronta, usar o arquivo propriedadesAdquira com Sharepoint.properties
         String arquivoProperties = "C:/Viegas/desenvolvimento/Selenium/arquivos propriedades/propriedadesAdquira.properties";
         
        FileInputStream file = new FileInputStream(arquivoProperties);
        props.load(file);
        return props;
    }

	
    public static String getValor(String chave) throws IOException{
    	Properties props = getProp();
        return (String)props.getProperty(chave);
    }
	
	   // Propiedades do driver para abrir no IE, Chrome ou Firefox
    public static WebDriver getWebDriver() throws InterruptedException {
    	
    	WebDriver driver = null;
    		
            try {
				
            	if ("Chrome".equals(getValor("navegador"))) {
				    
					File file = new File(getValor("driver.Chrome.selenium"));
					System.setProperty(getValor("propriedade.sistema.para.driver.Chrome.selenium"), file.getAbsolutePath());
				    DesiredCapabilities caps = DesiredCapabilities.chrome();
				    caps.setJavascriptEnabled(true);
				    caps.setCapability("ignoreZoomSetting", true);
				    caps.setCapability("nativeEvents",false);
				    ChromeOptions chromeOptions = new ChromeOptions(); 
				    Map<String, Object> chromePreferences = new HashMap<String, Object>();
					chromePreferences.put("profile.default_content_settings.popups", 0);
				    chromePreferences.put("download.default_directory",getValor("caminho.download.relatorios"));
				    chromePreferences.put("browser.helperApps.neverAsk.saveToDisk", "text/plain, application/vnd.ms-excel, text/csv, text/comma-separated-values, application/octet-stream");
				    chromeOptions.setExperimentalOption("prefs", chromePreferences);
				    chromeOptions.addArguments("--lang=pt");
				    
				    driver = new ChromeDriver(chromeOptions);
				    // Limpa o cache usando método do driver
				    driver.manage().deleteAllCookies();
				
				} else if ("internetExplorer".equals(getValor("navegador"))) {
				
					File file = new File(getValor("driver.internetExplorer.selenium"));
					System.setProperty(getValor("propriedade.sistema.para.driver.internetExplorer.selenium"), file.getAbsolutePath());
				    DesiredCapabilities caps = DesiredCapabilities.internetExplorer();
				    caps.setJavascriptEnabled(true);
				    //caps.setPlatform(org.openqa.selenium.Platform.WINDOWS);
				    caps.setCapability("ignoreZoomSetting", true);
				    caps.setCapability("nativeEvents",false);
					InternetExplorerOptions ieOptions = new InternetExplorerOptions();
					ieOptions.setCapability("ignoreZoomSetting", true);
					ieOptions.setCapability("nativeEvents",false);
					ieOptions.setCapability("browser.download.folderList", 2);
					ieOptions.setCapability("browser.helperApps.alwaysAsk.force", false);
					ieOptions.setCapability("browser.download.manager.showWhenStarting",false);
					//ieOptions.setCapability("browser.download.dir",getValor("caminho.download.relatorios"));
					//ieOptions.setCapability("browser.helperApps.neverAsk.saveToDisk", "text/plain, application/vnd.ms-excel, text/csv, text/comma-separated-values, application/octet-stream");
					//ieOptions.setCapability("browser.helperApps.alwaysAsk.force", true);
					// Limpando o cache com propriedades do Internet Explorer
					//caps.setCapability(InternetExplorerDriver.IE_ENSURE_CLEAN_SESSION,true);
					//driver = new InternetExplorerDriver(caps);
					//driver.manage().timeouts().implicitlyWait(0, TimeUnit.SECONDS);
					
					driver = new InternetExplorerDriver(ieOptions);
				    // Limpa o cache usando método do driver
				    driver.manage().deleteAllCookies();
				
				} else if ("Firefox".equals(getValor("navegador"))) {
					
					File file = new File(getValor("driver.Firefox.selenium"));
					System.setProperty(getValor("propriedade.binario.Firefox.selenium"),getValor("binario.Firefox")); 
					System.setProperty(getValor("propriedade.sistema.para.driver.Firefox.selenium"),file.getAbsolutePath());
				    FirefoxProfile fxProfile = new FirefoxProfile();
				    fxProfile.setPreference("browser.download.folderList",2);
				    fxProfile.setPreference("browser.download.manager.showWhenStarting",false);
				    fxProfile.setPreference("browser.download.dir",getValor("caminho.download.relatorios"));
				    fxProfile.setPreference("browser.helperApps.neverAsk.saveToDisk", "text/plain, application/vnd.ms-excel, application/zip, text/csv, text/comma-separated-values, application/octet-stream");
				    // Limpando o cache com propriedades do Firefox 
				    /*
				    fxProfile.setPreference("browser.cache.disk.enable", false);
				    fxProfile.setPreference("browser.cache.memory.enable", false);
				    fxProfile.setPreference("browser.cache.offline.enable", false);
				    fxProfile.setPreference("network.http.use-cache", false);
				    fxProfile.setPreference("network.cookie.cookieBehavior", 2);
				    */
				    
				    FirefoxOptions fxOptions = new FirefoxOptions();
				    fxOptions.setProfile(fxProfile);
				    driver = new FirefoxDriver(fxOptions);
				    // Limpa o cache usando método do driver
				    driver.manage().deleteAllCookies();
				}
			
            } catch (IOException e) {
				System.out.println("Ocorreu um erro no metodo getWebDriver: " + e.getMessage());
			}
            
            return driver;
    } 

}

