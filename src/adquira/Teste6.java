package adquira;
import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.FileInputStream;
import java.io.IOException;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Properties;

import org.openqa.selenium.WebDriverException;

public class Teste6 {
	
	public static void main(String[] args) throws InterruptedException, IOException, SQLException {
		
		System.out.println("Teste do Loop");
		
		String dataAtualPlanilhaFinal = new SimpleDateFormat("dd/MM/yyyy HH:mm:ss").format(new Date());
		
		inserirStatusExecucaoNoBanco("Adquira", dataAtualPlanilhaFinal, "Teste");
		
		Thread.sleep(1000);
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
    
	   public static void inserirStatusExecucaoNoBanco(String servico, String dataHora, String status) throws IOException, SQLException{
			  
		   HistoricoExecucaoDao historicoExecucaoDao = new HistoricoExecucaoDao();
		   historicoExecucaoDao.inserirStatusExecucao(servico, dataHora, status);
	   
	   }
	

}

