package adquira;
import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Properties;

import org.openqa.selenium.WebDriverException;

public class Teste5 {
	
	public static void main(String[] args) throws InterruptedException, IOException {
		

		String caminho = getValor("caminho.executavel.planilha.contract.numbers");
		
		Runtime.getRuntime().exec(caminho);
		
		Thread.sleep(10000);
		
        try {
            
        	Robot robot = new Robot();
            // Comando que pressiona as teclas CTRL + ALT + F5 que
            // que é usado para atualizar a planilha
            robot.keyPress(KeyEvent.VK_CONTROL);
            robot.keyPress(KeyEvent.VK_ALT);
            robot.keyPress(KeyEvent.VK_F5);
            robot.keyRelease(KeyEvent.VK_CONTROL);
            robot.keyRelease(KeyEvent.VK_ALT);
            robot.keyRelease(KeyEvent.VK_F5);
            Thread.sleep(20000);
            
           // Comando que pressiona as teclas CTRL + S que
           // que é usado para salvar a planilha
            robot.keyPress(KeyEvent.VK_CONTROL);
            robot.keyPress(KeyEvent.VK_S);
            robot.keyRelease(KeyEvent.VK_CONTROL);
            robot.keyRelease(KeyEvent.VK_S);
            Thread.sleep(5000);
            
            // Fechar a planilha
            robot.keyPress(KeyEvent.VK_ALT);
            robot.keyPress(KeyEvent.VK_F4);
            robot.keyRelease(KeyEvent.VK_ALT);
            robot.keyRelease(KeyEvent.VK_F4);
            Thread.sleep(5000);
            
        } catch (AWTException ex) {
            throw new WebDriverException("Erro ao digitar comandos", ex);
        } 
		
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
	

}

