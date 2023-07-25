package adquira;
import java.awt.Color;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStreamReader;
import java.math.BigDecimal;
import java.nio.file.attribute.BasicFileAttributes;
import java.sql.SQLException;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Properties;
import java.util.Set;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

import javax.swing.JFrame;
import javax.swing.JOptionPane;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.commons.io.FileUtils;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.firefox.FirefoxOptions;
import org.openqa.selenium.firefox.FirefoxProfile;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.ie.InternetExplorerOptions;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

import adquira.util.Pdf;
import adquira.util.Util;

public class AutomacaoAdquiraExtracaoPedidos {
	
	private static String nomeRelatorioBaixado;
	private static String nomeZipBaixado;
	private static boolean extracaoPossuiPedidos;
	private static boolean existemPedidos = false;
	private static String dataAtual = null;
	private static String dataAtualPlanilhaFinal = null;
	private static String dataAtualSharepoint = null;
	private static String diretorioLogs = "";
	private static String subdiretorioPdfsBaixados = "";
	private static List<ContractNumber> listaContractNumbers = null;
	private static List<ContractNumber> listaContractNumbersTemporaria = null;
	private static Set<String> listaNumerosContractNumbersDistintos = null;
	private static List<Pedido> listaPedidos = null;
	private static List<Pedido> listaPedidosFaturados = null;
	private static List<Pedido> listaPedidosNaoFaturados = null;
	private static List<Pedido> listaPedidosNaoFaturadosAuxiliar = null;
	private static String listaPedidosComErrosNasRegraDePreenchimentoNoSharePoint = "";
	private static int contadorErros;
	private static int contadorErroslerRelatorioExcel = 0;
	private static int contadorfazerDownlodRelatorioPorPeriodo = 0;
	private static int contadorErrosMoverArquivos = 0;
	private static int contadorErrosLogin = 0;
	private static int contadorErrosLogout = 0;
	private static int contadorErrosRecuperaContractNumbersSharepoint = 0;
	private static int contadorErrosRecuperaPedidosSharepoint = 0;
	private static int contadorErrosPreencherCamposBiling = 0;
	private static int contadorExecutaAutomacaoAdquiraSharepoint = 0;
	private static int contadorLogin = 0;
	private static String diretorioRelatorio = null;
	private static String subdiretorioRelatoriosBaixados = null;
	private static String subdiretorioRelatoriosBaixados2 = null;
	private static String subdiretorioRelatorioFinal = null;
	private static String subdiretorioRelatorioFinal2  = null;
	private static String subdiretorioRelatorioIncremental = null;   
	private static String subdiretorioPdfsBaixados2  = null;
	private static String caminhoExecutavelPlanilhaContractNumbers = null;
	private static String caminhoExecutavelPlanilhaPedidosFaturados = null;
	
	@SuppressWarnings("unused")
	public static void main(String[] args) throws Exception {
		
		String usernameSP = Util.getValor("usernameSP");
		String senhaSP = Util.getValor("senhaSP");
		automacaoAdquiraSharepoint(usernameSP, senhaSP);
		
		String usernameSBC = Util.getValor("usernameSBC");
		String senhaSBC = Util.getValor("senhaSBC");
		automacaoAdquiraSharepoint(usernameSBC, senhaSBC);
		
    		
	}
	
	public static void automacaoAdquiraSharepoint(String usuario, String senha) throws Exception{
		
		WebDriver driver = null;
		
		try {
			
			dataAtual = new SimpleDateFormat("yyyy_MM_dd HH_mm_ss").format(new Date());
			dataAtualPlanilhaFinal = new SimpleDateFormat("dd/MM/yyyy HH:mm:ss").format(new Date());
			dataAtualSharepoint = new SimpleDateFormat("MM/dd/yyyy").format(new Date());
			diretorioLogs = Util.getValor("caminho.diretorio.relatorios") + "/" + dataAtual;
			diretorioRelatorio = Util.getValor("caminho.download.relatorios") + "\\" + dataAtual;
			subdiretorioRelatoriosBaixados = diretorioRelatorio + "\\" + "relatorios baixados " + dataAtual;
			subdiretorioRelatoriosBaixados2 = diretorioLogs + "/" + "relatorios baixados " + dataAtual;
			subdiretorioRelatorioFinal = diretorioRelatorio + "\\" + "relatorio final " + dataAtual;
			subdiretorioRelatorioFinal2  = diretorioRelatorio + "/" + "relatorio final " + dataAtual;
			subdiretorioRelatorioIncremental = Util.getValor("caminho.diretorio.relatorio.incremental");   
			subdiretorioPdfsBaixados  = diretorioRelatorio + "\\" + "pdfs baixados " + dataAtual;
			subdiretorioPdfsBaixados2  = diretorioRelatorio + "/" + "pdfs baixados " + dataAtual;
			caminhoExecutavelPlanilhaContractNumbers = Util.getValor("caminho.executavel.planilha.contract.numbers");
			caminhoExecutavelPlanilhaPedidosFaturados = Util.getValor("caminho.executavel.planilha.pedidos.faturados");
			criaDiretorio(subdiretorioRelatoriosBaixados);
			criaDiretorio(subdiretorioRelatorioFinal);
			criaDiretorio(subdiretorioRelatorioIncremental);
			criaDiretorio(subdiretorioPdfsBaixados);
			
			// Deleta os diret�rios que possu�rem data de cria��o anterior � data de 7 dias atr�s
			apagaDiretoriosDeRelatorios(Util.getValor("caminho.download.relatorios"));
			
			// As vezes o diret�rio que armazena dados tempor�rios do Chome simplesmente some, da� o Selenium d� pau na hora de chamar o browser
			// Com o m�todo abaixo, crio essa pasta se ela n�o existir
			criaDiretorioTemp();
			
			executaAutomacaoAdquiraSharepoint(driver, usuario, senha);
			
		} catch (Exception e) {
			gravarArquivo(diretorioLogs, "Erro Adquira" + " " + dataAtual, ".txt", e.getMessage(), "Ocorreu um erro na automacao de extracao de pedidos: ");
			inserirStatusExecucaoNoBanco("Adquira", dataAtualPlanilhaFinal, "Erro de execucao do robo");
		} finally {
			//mensagemErro("Houve um problema na extra��o dos pedidos no Adquira\n");
			//fazerLogout(wait);
			if (driver != null) {
				driver.quit();
			}
			
			mataProcessosGoogle();
			mataProcessosFirefox();
			
		}
	
	}
	
    public static void executaAutomacaoAdquiraSharepoint(WebDriver driver, String usuario, String senha) throws Exception{
    	
    	try {
    		
    		String mensagemResultadoAdquira = "0 Pedidos encontrados";
    		
    		if (driver != null) {
    			driver.quit();
    		}
    		
    		mataProcessosGoogle();
    		mataProcessosFirefox();
    		
    		listaContractNumbers = new ArrayList<ContractNumber>();
    		listaContractNumbersTemporaria = new ArrayList<ContractNumber>();
    		listaNumerosContractNumbersDistintos = new HashSet<String>();
    		listaPedidos = new ArrayList<Pedido>();
    		listaPedidosFaturados = new ArrayList<Pedido>();
    		listaPedidosNaoFaturados = new ArrayList<Pedido>();
    		listaPedidosNaoFaturadosAuxiliar = new ArrayList<Pedido>();
    		
    		// Abre e atualiza planilha de Contract Numbers
    		//abreAtualizaPlanilha(caminhoExecutavelPlanilhaContractNumbers);
    		
    		// Abre e atualiza planilha de Pedidos Faturados
    		//abreAtualizaPlanilha(caminhoExecutavelPlanilhaPedidosFaturados);
    		
    		System.out.println("Inicio: " + new SimpleDateFormat("dd_MM_yyyy HH_mm_ss").format(new Date()));
    		// Abrindo a URl do SharePoint somente para fazer o login
    		//fazerLoginSharepoint(driver, wait);
    		
    		driver = getWebDriver();
    		JavascriptExecutor js = (JavascriptExecutor) driver;
    		WebDriverWait wait = new WebDriverWait(driver, 60);
    		
    		// Obtendo os contract numbers que est�o na planilha
    		//lerPlanilhaContractNumbers(Util.getValor("caminho.contract.numbers"));
    		
    		// Obtendo os contract numbers ativos do site do sharepoint
    		recuperaContractNumbersSharepoint(driver, wait);
    		
    		// Obtendo os pedidos do site do sharepoint
    		recuperaPedidosSharepoint(driver, wait);
    		
    		// Criando a lista de contract numbers sem repeti��o de n�meros de contratos
    		criaListaContractNumbersDistintos();
    		
    		// Faz login no Adquira
    		fazerLoginAdquira(driver, wait, js, usuario, senha);
    		
    		//acessarPaginaInicial(driver, wait);
    		
    		// Viegas
    		// Trecho de c�digo para testar com algum contract number espec�fico
    		/*
    		listaContractNumbers = new ArrayList<ContractNumber>();
    		ContractNumber contractNumberTeste = new ContractNumber();
    		
    		contractNumberTeste.setContrato("Projeto Teste");
    		contractNumberTeste.setFrente("Frente Teste");
    		contractNumberTeste.setNumero("4100102109");
    		
    		listaContractNumbers.add(contractNumberTeste);
    		*/
    		
    		fazerDownlodRelatorioPorPeriodo(driver, wait, js, usuario, senha);
    		
    		if (extracaoPossuiPedidos) {
    			existemPedidos = true;
    			//Move o relat�rio baixado do diret�rio relatorios para o diret�rio correto
    			moverArquivosEntreDiretorios(driver, wait, js, Util.getValor("caminho.download.relatorios") + "\\" + nomeRelatorioBaixado, subdiretorioRelatoriosBaixados, usuario, senha);
    			Thread.sleep(5000);
    			
    			// L� o relat�rio baixado
    			lerRelatorioExcel(driver, wait, js, subdiretorioRelatoriosBaixados2 + "/" + nomeRelatorioBaixado, subdiretorioRelatoriosBaixados, usuario, senha);
    			
    		}
    		
    		// Se existem pedidos, fa�o a subtra��o da lista desses pedidos com a lista de pedidos faturados
    		// Teremos ent�o uma lista de pedidos n�o faturados, onde desses farei o download dos arquivos pdfs zipados e gero o relat�rio final
    		if (existemPedidos) {
    			
    			if (listaPedidos != null && listaPedidos.size() > 0) {
    				
    				// Crio a lista de pedidos n�o faturados inicialmente com a lista completa de pedidos
    				listaPedidosNaoFaturados.addAll(listaPedidos);
    				
    				// Obtenho a lista de pedidos faturados atrav�s da planilha
    				//lerPlanilhaPedidosFaturados(Util.getValor("caminho.pedidos.faturados")); 
    				
    				for (Pedido pedido : listaPedidos) {
    					
    					if (listaPedidosFaturados != null && listaPedidosFaturados.size() > 0) {
    						
    						for (Pedido pedidoFaturado : listaPedidosFaturados) {
    							
    							if (pedido.getNumero().trim().equals(pedidoFaturado.getNumero().trim())) {
    								
    								pedido.setFaturado(true);
    								
    								listaPedidosNaoFaturados.remove(pedido);
    								
    							}
    							
    						}
    						
    					}
    					
    				}
    				
    				if (listaPedidosNaoFaturados != null && listaPedidosNaoFaturados.size() > 0) {
    					
    					for (Pedido pedidoNaoFaturado: listaPedidosNaoFaturados) {
    						
    						fazerDownlodPdfPedidoMoveArquivosEDescompacta(driver, wait, js, pedidoNaoFaturado, subdiretorioPdfsBaixados2, usuario, senha);
    						
    						if (pedidoNaoFaturado.isEncontrouPdfAnexo()) {
    							encontrarContractNumberNoPdfDoPedido(driver, wait, js, pedidoNaoFaturado, subdiretorioPdfsBaixados2);
    						}
    						
    					}
    					
    					// Retirando os pedidos que contenham contract numbers inv�lidos
    					// Crio uma lista auxiliar com os pedidos n�o faturados somente para retirar os pedidos
    					// que contenham contract numbers inv�lidos da lista de pedidos n�o faturados.
    					listaPedidosNaoFaturadosAuxiliar.addAll(listaPedidosNaoFaturados);
    					
    					for (Pedido pedidoNaoFaturadoAuxiliar: listaPedidosNaoFaturadosAuxiliar) {
    						
    						if (pedidoNaoFaturadoAuxiliar.isEncontrouPdfAnexo()) {
    							
    							if (isPedidoComContractNumberInvalido(pedidoNaoFaturadoAuxiliar)) {
    								
    								listaPedidosNaoFaturados.remove(pedidoNaoFaturadoAuxiliar);
    								
    							}
    						
    						}
    						
    					}
    					
    					// Verificando novamente, pois pode ser que n�o tenha nada na lista depois da retirada dos pedidos 
    					// com contract numbers inv�lidos feito acima
    					if (listaPedidosNaoFaturados != null && listaPedidosNaoFaturados.size() > 0) {
    						
    						fazerLogoutAdquira(driver, wait);
    						
    						criarPedidosSharepoint(driver, wait, js, listaPedidosNaoFaturados);
    						
    						// Cria arquivo excel que conter� o relat�rio final contendo os pedidos n�o faturados
    						// Tamb�m armazenar� se o pedido foi salvo no sharepoint
    						//gravarArquivo(subdiretorioRelatorioFinal2, "relatorio final" + " " + dataAtual, ".xls", "", "");
    						
    						// Gera o relat�rio final de pedidos n�o faturados
    						//String relatorioFinal = subdiretorioRelatorioFinal2 + "/" + "relatorio final" + " " + dataAtual + ".xls";
    						//criarRelatorioFinal(relatorioFinal);
    						
    						// Gera o relat�rio incremental de pedidos faturados e n�o faturados
    						//String relatorioIncremental = subdiretorioRelatorioIncremental + "/" + "relatorio incremental" + ".xls";
    						// Se o relat�rio n�o existir, crio um novo
    						//if (!existeArquivo(relatorioIncremental)) {
    							// N�o vou criar o relat�rio em excel por enquanto, pois j� est� sendo gravado no banco
    							//criarRelatorioFinal(relatorioIncremental);
    							// Se o relat�rio existir, o incremento
    						//} else {
    							// N�o vou criar o relat�rio em excel por enquanto, pois j� est� sendo gravado no banco
    							//preenchePlanilhaRelatorioIncremental(relatorioIncremental);
    						//}
    						
    						// Insere no banco a lista contendo todos os pedidos n�o faturados
    						inserePedidosNoBanco();
    						
    						// Cria arquivo contendo a data e hora da extra��o dos pedidos no Adquira
    						// Ele ser� usado pela automa��o do sharepoint
    						//gravarArquivo(getValor("caminho.diretorio.relatorios"), "data e hora dos pedidos para o sharepoint", ".txt", dataAtual, "data.hora.pedidos.para.sharepoint=");
    						
    						// Encontrando a quantidade de pedidos salvos no sharepoint
    						int quantidadePedidosSalvosNoSharepoint = 0;
    						for (Pedido pedidoNaoFaturado: listaPedidosNaoFaturados) {
    							
    							if (pedidoNaoFaturado.isSalvoNoSharepoint()) {
    								
    								quantidadePedidosSalvosNoSharepoint = quantidadePedidosSalvosNoSharepoint + 1;
    							}
    						}
    						
    						String mensagemPedidos = "";
    						
    						if (listaPedidosNaoFaturados.size() > 1) {
    							
    							mensagemPedidos = " Pedidos encontrados e ";
    							
    						} else {
    							
    							mensagemPedidos = " Pedido encontrado e ";
    							
    						}
    						
    						String mensagemPedidosSalvosNoSharepoint = "";
    						
    						if (quantidadePedidosSalvosNoSharepoint > 1) {
    							
    							mensagemPedidosSalvosNoSharepoint = " cadastrados";
    							
    						} else {
    							
    							mensagemPedidosSalvosNoSharepoint = " cadastrado";
    							
    						}
    						
    						mensagemResultadoAdquira = listaPedidosNaoFaturados.size() + mensagemPedidos + quantidadePedidosSalvosNoSharepoint + mensagemPedidosSalvosNoSharepoint;
    	    			
    					} else {
    	    				fazerLogoutAdquira(driver, wait);
    	    				mensagemResultadoAdquira = "0 Pedidos encontrados";
    	    			}
    					
    					
    				} else {
    					
    					fazerLogoutAdquira(driver, wait);
    					mensagemResultadoAdquira = "0 Pedidos encontrados";
    				
    				}
    				
    			} else {
    				
    				fazerLogoutAdquira(driver, wait);
    				mensagemResultadoAdquira = "0 Pedidos encontrados";
    				
    			}
    			
    		} else {
    			
    			fazerLogoutAdquira(driver, wait);
    			mensagemResultadoAdquira = "0 Pedidos encontrados";
    			
    		}
    		
    		// Gravo em um arquivo os Pedidos que tiveram problemas nas regras de preenchimento
    		if (listaPedidosComErrosNasRegraDePreenchimentoNoSharePoint != null && !listaPedidosComErrosNasRegraDePreenchimentoNoSharePoint.isEmpty()) {
    			gravarArquivo(subdiretorioRelatorioFinal2, "Pedidos com Erros nas Regras de Preenchimento no Sharepoint" + " " + dataAtual, ".txt", listaPedidosComErrosNasRegraDePreenchimentoNoSharePoint, "");
    		}
    		
    		gravarArquivo(diretorioLogs, "Resultado Adquira" + " " + dataAtual, ".txt", "", mensagemResultadoAdquira);
    		
    		// Grava na tabela Tb_Historico_Execucao_Robos o servi�o, data e hora e status da execu��o
    		inserirStatusExecucaoNoBanco("Adquira", dataAtualPlanilhaFinal, mensagemResultadoAdquira);
    		
    		//mensagemSucesso();
    		
    		System.out.println("Fim: " + new SimpleDateFormat("dd_MM_yyyy HH_mm_ss").format(new Date()));
    		
    		if (driver != null) {
    			driver.quit();
    		}
    		
    		mataProcessosGoogle();
    		mataProcessosFirefox();
			
		} catch (Exception e) {
			contadorExecutaAutomacaoAdquiraSharepoint ++;
			// Executo ate 5 vezes se der erro no executaAutomacaoAdquiraSharepoint
			if (contadorExecutaAutomacaoAdquiraSharepoint <= 5) {
				
				System.out.println("Deu erro no metodo executaAutomacaoAdquiraSharepoint, tentativa de acerto: " + contadorExecutaAutomacaoAdquiraSharepoint);
				executaAutomacaoAdquiraSharepoint(driver, usuario, senha);
			
			} else {
				throw new Exception("Ocorreu um erro no m�todo executaAutomacaoAdquiraSharepoint: " + e);
		    }

		}
    	
    }
    
	   public static void mataProcessosFirefox() throws IOException, SQLException, InterruptedException{
		   
		   // Mata o Firefox
		   Runtime.getRuntime().exec(Util.getValor("caminho.matar.firefox"));
		   Thread.sleep(3000);
	   
	   }

	   public static void mataProcessosGoogle() throws IOException, SQLException, InterruptedException{
			  
		   // Mata o Google
		   // Viegas
		   Runtime.getRuntime().exec(Util.getValor("caminho.matar.google"));
		   Thread.sleep(3000);
		   
		   // Mata o chromedriver
		   Runtime.getRuntime().exec(Util.getValor("caminho.matar.chromedriver"));
		   Thread.sleep(3000);
	   
	   }
	   
	   public static void criaDiretorioTemp(){
	    	
		   	String str1 = "echo %temp%";
		   	String command = "C:\\WINDOWS\\system32\\cmd.exe /y /c " + str1;
		    	
		   	try {
		    		
		   		Process processo = Runtime.getRuntime().exec(command);
		   		String line;
		   		String caminhoTemp = "";
		    		
		   		//pega o retorno do processo
		   		BufferedReader stdInput = new BufferedReader(new 
		   				InputStreamReader(processo.getInputStream()));
		    		
		   		//printa o retorno
		   		while ((line = stdInput.readLine()) != null) {
		   			caminhoTemp = line;
		   		}
		    		
		   		criaDiretorio(caminhoTemp);
		    		
		   	} catch (Exception e) {
		   		System.out.println("Deu erro na cria��o do diret�rio Temp: " + e.getMessage());
		   	}

	   }
    
    public static void criaDiretorio(String caminhoDiretorio){
        File diretorio = new File(caminhoDiretorio);
        if (!diretorio.exists()) {
        	diretorio.mkdirs();
        }
    }
    
    public static void moverArquivosEntreDiretorios(WebDriver driver, WebDriverWait wait, JavascriptExecutor js, String caminhoArquivoOrigem, String caminhoDiretorioDestino, String usuario, String senha) throws Exception{
    	
    	boolean sucesso = true;
    	File arquivoOrigem = new File(caminhoArquivoOrigem);
        File diretorioDestino = new File(caminhoDiretorioDestino);
        if (arquivoOrigem.exists() && diretorioDestino.exists()) {
        	sucesso = arquivoOrigem.renameTo(new File(diretorioDestino, arquivoOrigem.getName()));
        }
        
        if (!sucesso) {
        	contadorErrosMoverArquivos++;
        	
            // Tento mover o arquivo por at� 20 vezes
            if (contadorErrosMoverArquivos <= 20) {
            	
				System.out.println("Deu erro no m�todo moverArquivosEntreDiretorios, tentativa de acerto: " + contadorErrosMoverArquivos);
				// Est� dando erro de logout no servidor
				// O bot�o de logout est� ficando escondido
				// ent�o retirarei o logout e o login por enquanto
				fazerLogoutAdquira(driver, wait);
				fazerLoginAdquira(driver, wait, js, usuario, senha);
				//acessarPaginaInicial(driver, wait);
				fazerDownlodRelatorioPorPeriodo(driver, wait, js, usuario, senha);
				moverArquivosEntreDiretorios(driver, wait, js, caminhoArquivoOrigem, caminhoDiretorioDestino, usuario, senha);
            
            } else {
            	throw new Exception("Ocorreu um erro no momento de mover o relat�rio " + caminhoArquivoOrigem + " para " + caminhoDiretorioDestino);
            }
        	
        }
        
    }
    
	  public static void apagaArquivosDiretorioDeRelatorios(String caminhoDiretorio) throws Exception{
	    	
	    	boolean sucesso = false;
	        File diretorio = new File(caminhoDiretorio);
	        if (diretorio.exists() && diretorio.isDirectory()) {
	        	sucesso = true;
	        	
	        	//lista os nomes dos arquivos
				String arquivos [] = diretorio.list();
				
				if (arquivos != null && arquivos.length > 0) {
					
					for (String item : arquivos){
						
						File arquivo = new File(caminhoDiretorio + "/" + item);
						// Se existirem arquivos, os deleto
						if (arquivo.exists() && arquivo.isFile()) {
							arquivo.delete();
						}

					}
				}
	        	
	        }
	        
	        if (!sucesso) {
	        	throw new Exception("Nao existe o diretorio: " + caminhoDiretorio);
	        }
	        
	    }
	  
	  public static void apagaDiretoriosDeRelatorios(String caminhoDiretorio) throws Exception{
	    	
	    	boolean sucesso = false;
	        File diretorio = new File(caminhoDiretorio);
	        Date dataAtual = new Date();
	        Calendar cal = Calendar.getInstance();
	        cal.setTime(dataAtual);
	        cal.add(Calendar.DATE, -7);
	        Date dataAntes7Dias = cal.getTime();
	        
	        if (diretorio.exists() && diretorio.isDirectory()) {
	        	sucesso = true;
	        	
	        	//lista os nomes dos diret�rios
				String itens [] = diretorio.list();
				
				if (itens != null && itens.length > 0) {
					
					for (String item : itens){
						
						File pasta = new File(caminhoDiretorio + "/" + item);
						
						if (pasta.exists() && pasta.isDirectory()) {
							
							Long dataModificacaoPasta =  FileUtils.lastModified(pasta);
							Date dataModificacaoPasta2 = new Date(dataModificacaoPasta);
							
							// Se existirem diret�rios com a data anterior � data de 7 dias atr�s, os deleto
							if (dataModificacaoPasta2.before(dataAntes7Dias)) {
								FileUtils.deleteQuietly(pasta);
							}
							
						}

					}

				}
	        	
	        }
	        
	        if (!sucesso) {
	        	throw new Exception("Nao existe o diretorio: " + caminhoDiretorio);
	        }
	        
	    }

	  public static boolean existeArquivo(String caminhoArquivo) throws Exception{
	    	
	    	boolean existeArquivo = false;
	        File arquivo = new File(caminhoArquivo);
	        if (arquivo.exists() && !arquivo.isDirectory()) {
	        	existeArquivo = true;
	        }
			return existeArquivo;
	    }
	  
	  public static void abreAtualizaPlanilha(String caminhoArquivo) throws Exception{
			
		  
		    System.out.println("Abrindo a planilha..."); 
			// Executa batch que abre a planilha
		    Runtime.getRuntime().exec(caminhoArquivo);
			
			// Tempo para abrir a planilha 
			Thread.sleep(10000);
			
	        try {
	            
	        	Robot robot = new Robot();
	            // Comando que pressiona as teclas CTRL + ALT + F5 que
	            // que � usado para atualizar a planilha
	            robot.keyPress(KeyEvent.VK_CONTROL);
	            robot.keyPress(KeyEvent.VK_ALT);
	            robot.keyPress(KeyEvent.VK_F5);
	            robot.keyRelease(KeyEvent.VK_CONTROL);
	            robot.keyRelease(KeyEvent.VK_ALT);
	            robot.keyRelease(KeyEvent.VK_F5);
	            Thread.sleep(20000);
	            
	           // Comando que pressiona as teclas CTRL + S que
	           // que � usado para salvar a planilha
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
	            
	        } catch (Exception ex) {
	            throw new Exception("Erro ao digitar comandos para atualizar a planilha", ex);
	        } 
			
		}
	  
	   public static void inserePedidosNoBanco() throws IOException, SQLException{
			if (listaPedidosNaoFaturados != null && listaPedidosNaoFaturados.size() > 0) {
				int contador = 0;
				List<Pedido> listaPedidosDoBanco = new ArrayList<Pedido>();
				for (Pedido pedidosFaturadosNaoFaturados : listaPedidosNaoFaturados) {
					
					// Converto os valores de null para espa�o em branco para gravar branco no banco e n�o dar problema
					// no futuro relat�rio do sharepoint da Accenture que teremos
					// Esse relat�rio de sharepoint da Accenture ser� gerado atrav�s do banco
					Util.converteValorNullParaEspacoEmBranco(pedidosFaturadosNaoFaturados);
					
					preencherMensagemDeErroNoPedido(pedidosFaturadosNaoFaturados);

					AdquiraDao adquiraDao = new AdquiraDao();
					listaPedidosDoBanco = adquiraDao.recuperaPedidos(pedidosFaturadosNaoFaturados);
					boolean inserePedidoNoBanco = true;
					
					// Recupero pedidos do banco para n�o inserir de novo pedidos que tenham as mesmas mensagens do campo Erro_No_Pedido
					// e que tamb�m n�o possuam o mesmo n�mero de Contract Number
					if (listaPedidosDoBanco != null && !listaPedidosDoBanco.isEmpty()) {
						
						for (Pedido pedidoDoBanco : listaPedidosDoBanco) {
							
							boolean condicaoParaNaoInserirNoBanco = (pedidoDoBanco != null
									 								&& pedidoDoBanco.getNumero().trim().equals(pedidosFaturadosNaoFaturados.getNumero().trim())
   								 								    && pedidoDoBanco.getNumeroContractNumber().trim().equals(pedidosFaturadosNaoFaturados.getContractNumber().getNumero().trim())
																	&& pedidoDoBanco.getMensagemDeErroNoPedido().trim().equals(pedidosFaturadosNaoFaturados.getMensagemDeErroNoPedido().trim()));
							if (condicaoParaNaoInserirNoBanco) {
								inserePedidoNoBanco = false;
							}
							
						}						
	
					}
					
					if (inserePedidoNoBanco) {
						AdquiraDao adquiraDao2 = new AdquiraDao();
						adquiraDao2.inserirPedido(pedidosFaturadosNaoFaturados);
						contador++;
						System.out.println("Inseri o pedido de n�mero: " + contador + " de um total de " + listaPedidosNaoFaturados.size());
					}
					
				}
				
			}

	   }
	   
	   public static void inserirStatusExecucaoNoBanco(String servico, String dataHora, String status) throws IOException, SQLException{
		  
		   HistoricoExecucaoDao historicoExecucaoDao = new HistoricoExecucaoDao();
		   historicoExecucaoDao.inserirStatusExecucao(servico, dataHora, status);
	   
	   }

    // Fazer login no SharePoint
    public static void fazerLoginSharepoint(WebDriver driver, WebDriverWait wait) throws InterruptedException, IOException  {
		
    	driver.manage().window().maximize();
		driver.get(Util.getValor("url.sharepoint"));
		Thread.sleep(1000);
		
		// Abrindo a URl
		driver.manage().window().maximize();
		driver.get(Util.getValor("url.sharepoint"));
		Thread.sleep(1000);
        
		// Preenchendo dados do e-mail da Accenture
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("i0116")));
        WebElement emailAccenture = driver.findElement(By.id("i0116"));
        emailAccenture.sendKeys(Util.getValor("emailAccenture"));
        Thread.sleep(1000);
        
		// Clicar no bot�o Avan�ar
        wait.until(ExpectedConditions.elementToBeClickable(By.id("idSIButton9"))).click();
        
		// Preenchendo dados da senha da Accenture
        wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("passwordInput")));
        WebElement senhaAccenture = driver.findElement(By.id("passwordInput"));
        senhaAccenture.sendKeys(Util.getValor("senhaAccenture"));
        Thread.sleep(1000);
        
		// Clicar no bot�o de Sign in para fazer o login
        wait.until(ExpectedConditions.elementToBeClickable(By.id("submitButton"))).click();
        Thread.sleep(1000);
        
		// Clicar no bot�o de Ignorar para lembrar o dispositivo
        wait.until(ExpectedConditions.elementToBeClickable(By.id("vipSkipBtn"))).click();
        
		// Clicar no bot�o Sim para continuar conectado
        wait.until(ExpectedConditions.elementToBeClickable(By.id("idSIButton9"))).click();
        Thread.sleep(3000);

    }
    
	public static void criarPedidosSharepoint(WebDriver driver, WebDriverWait wait, JavascriptExecutor js, List<Pedido> listaPedidos) throws Exception  { 
		
		// Abrindo a URl
		driver.manage().window().maximize();
		driver.get(Util.getValor("url.sharepoint"));
		Thread.sleep(1000);
        
        System.out.println("Inicio Sharepoint: " + new SimpleDateFormat("dd_MM_yyyy HH_mm_ss").format(new Date()));
        int cont = 0;
        
        contadorErros = 0;
	    for (Pedido pedido : listaPedidos) {
	    	
    		if (!pedido.isSalvoNoSharepoint() && pedido.isEncontrouPdfAnexo() && pedido.isContractNumberConforme()) {
    			
    			preencherCamposBiling(driver, wait, js, pedido);
    		}

	    	System.out.println("Contador de pedidos: " + cont++ + " de um total de " + listaPedidos.size() );
	    	
        }
	    
	    System.out.println("Fim Sharepoint: " + new SimpleDateFormat("dd_MM_yyyy HH_mm_ss").format(new Date()));

	}
	
    public static void verificaSePedidoExisteNoSharePoint(WebDriver driver, WebDriverWait wait, JavascriptExecutor js, Pedido pedido) throws InterruptedException, IOException  { 

    	try {
    		
    		WebDriverWait waitPedidoExistenteSharePoint = new WebDriverWait(driver,3);
    		// Clicar no campo de busca do pedido
    		wait.until(ExpectedConditions.elementToBeClickable(By.id("inplaceSearchDiv_WPQ1_lsinput"))).click();
    		
			// Campo Contrato SAP
			String idCampoBusca = "inplaceSearchDiv_WPQ1_lsinput";
			wait.until(ExpectedConditions.elementToBeClickable(By.id(idCampoBusca))).click();
			Thread.sleep(1000);
			WebElement campoBusca  = driver.findElement(By.id(idCampoBusca));
			campoBusca.clear();
			campoBusca.sendKeys(pedido.getNumero());
			Thread.sleep(1000);
			campoBusca.sendKeys(Keys.RETURN);
			Thread.sleep(1000);
			
			// Se existe o link Status, ent�o a busca retornou resultados
			try {
				waitPedidoExistenteSharePoint.until( ExpectedConditions.visibilityOfElementLocated(By.xpath("(//a[contains(text(),'Status')])[2]")));
				System.out.println("Achou pedido: " + pedido.getNumero() + " no sharepoint");
				pedido.setSalvoNoSharepoint(true);
				Thread.sleep(1000);
			} catch (Exception e) {
				System.out.println("N�o achou pedido: " + pedido.getNumero() + " no sharepoint");
				Thread.sleep(1000);
			}
    		
		} catch (Exception e) {
			System.out.println("Passei no erro interno do verificaPedidoExistente: " + pedido.getNumero());
			
		}

    }	
	
    public static void clickNoCampoCaps( WebDriverWait wait, String idCap) {
    	// Estava dando o seguinte erro: org.openqa.selenium.StaleElementReferenceException: stale element reference: element is not attached to the page document
    	// Sempre que voc� enfrentar esse problema, apenas defina o elemento da web mais uma vez acima da linha em que voc� est� obtendo um erro
    	try {
    		wait.until(ExpectedConditions.elementToBeClickable(By.id(idCap))).click();
    	}
    	catch(org.openqa.selenium.StaleElementReferenceException ex) {
    		wait.until(ExpectedConditions.elementToBeClickable(By.id(idCap))).click();
    	}
    }
    
    public static void clickNoCampoObservacao( WebDriverWait wait, String idObservacao) {
		// Estava dando o seguinte erro: org.openqa.selenium.StaleElementReferenceException: stale element reference: element is not attached to the page document
		// Sempre que voc� enfrentar esse problema, apenas defina o elemento da web mais uma vez acima da linha em que voc� est� obtendo um erro
		try {
			wait.until(ExpectedConditions.elementToBeClickable(By.id(idObservacao))).click();
		}
		catch(org.openqa.selenium.StaleElementReferenceException ex) {
			wait.until(ExpectedConditions.elementToBeClickable(By.id(idObservacao))).click();
		}
    }

    public static void preencherCamposBiling(WebDriver driver, WebDriverWait wait, JavascriptExecutor js, Pedido pedido) throws Exception  { 

    	try {
    		driver.get(Util.getValor("url.sharepoint"));
    		Thread.sleep(3000);
    		
    		// Link new item
    		String textoNewItem = "new item";
    		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//span [text()='"+textoNewItem+"']"))).click();
    		Thread.sleep(3000);
    		
    		// Bot�o para anexar arquivo
    		String textoBotaoAnexarArquivo = "Click here to attach a file";
    		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//span [text()='"+textoBotaoAnexarArquivo+"']"))).click();
    		
    		// Bot�o para escolher o arquivo para anexar
    		wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("FileAttachmentUpload"))).sendKeys(subdiretorioPdfsBaixados +"\\"+ pedido.getNumero() + ".pdf");
    		Thread.sleep(1000);
    		wait.until(ExpectedConditions.elementToBeClickable(By.id("DialogButton0"))).click();
    		Thread.sleep(3000);
    		
    		// Scrollar para baixo um pouco a tela
    		// js.executeScript("window.scrollBy(0,1000)");
    		
    		String inicioIdCampos = "ctl00_ctl33_g_de9b44a5_3d1f_4e84_8896_8c3974a46081_FormControl0_";
    		
    		// Campo Status Solicita��o
    		String idComboStatusSolicitacao = inicioIdCampos + "V1_I1_D1";
    		// Estava dando o seguinte erro: org.openqa.selenium.StaleElementReferenceException: stale element reference: element is not attached to the page document
    		// Sempre que voc� enfrentar esse problema, apenas defina o elemento da web mais uma vez acima da linha em que voc� est� obtendo um erro
    		try {
    			wait.until(ExpectedConditions.elementToBeClickable(By.id(idComboStatusSolicitacao))).click();
    		}
    		catch(org.openqa.selenium.StaleElementReferenceException ex) {
    			wait.until(ExpectedConditions.elementToBeClickable(By.id(idComboStatusSolicitacao))).click();
    		}
    		WebElement comboStatusSolicitacao = driver.findElement(By.id(idComboStatusSolicitacao));
    		// Elementos do combo
    		Select elementosComboStatusSolicitacao  = new Select(comboStatusSolicitacao);
    		String statusSolictacao = "Aberto";
    		boolean existeStatusSolicitacao = false;
    		// Verifico se existe o status no combo
    		// Se n�o existir, lan�o exce��o
    		int quantidadeElementosComboStatusSolicitacao  = elementosComboStatusSolicitacao.getOptions().size();
    		if (quantidadeElementosComboStatusSolicitacao > 0) {
    			for (WebElement elemento : elementosComboStatusSolicitacao.getOptions()) {
    				if (statusSolictacao.equalsIgnoreCase(elemento.getText().trim())) {
    					existeStatusSolicitacao = true;
    					break;
    				}
    			}
    		}
    		if (existeStatusSolicitacao) {
    			elementosComboStatusSolicitacao.selectByVisibleText(statusSolictacao);
    			Thread.sleep(5000);
    		} else {
    			throw new Exception("Regra de preenchimento inv�lida. O campo Status Solicita��o do Sharepoint n�o possui a op��o " + statusSolictacao + "\n");
    		}
    		
    		// Campo N� Ctro SAP TLF
    		String idComboNumeroContrato = inicioIdCampos + "V1_I1_D3";
    		// Estava dando o seguinte erro: org.openqa.selenium.StaleElementReferenceException: stale element reference: element is not attached to the page document
    		// Sempre que voc� enfrentar esse problema, apenas defina o elemento da web mais uma vez acima da linha em que voc� est� obtendo um erro
    		try {
    			wait.until(ExpectedConditions.elementToBeClickable(By.id(idComboNumeroContrato))).click();
    		}
    		catch(org.openqa.selenium.StaleElementReferenceException ex) {
    			wait.until(ExpectedConditions.elementToBeClickable(By.id(idComboNumeroContrato))).click();
    		}
    		WebElement comboNumeroContrato = driver.findElement(By.id(idComboNumeroContrato));
    		// Elementos do combo
    		Select elementosComboNumeroContrato  = new Select(comboNumeroContrato);
    		//String numeroContrato = "4100094911";
    		String numeroContrato = pedido.getContractNumber().getNumero();
    		boolean existeNumeroContrato = false;
    		// Verifico se existe o status no combo
    		// Se n�o existir, lan�o exce��o
    		int quantidadeElementosComboNumeroContrato  = elementosComboNumeroContrato.getOptions().size();
    		if (quantidadeElementosComboNumeroContrato > 0) {
    			for (WebElement elemento : elementosComboNumeroContrato.getOptions()) {
    				if (numeroContrato.equalsIgnoreCase(elemento.getText().trim())) {
    					existeNumeroContrato = true;
    					break;
    				}
    			}
    		}
    		if (existeNumeroContrato) {
    			elementosComboNumeroContrato.selectByVisibleText(numeroContrato);
    			Thread.sleep(5000);
    		} else {
    			throw new Exception("Regra de preenchimento inv�lida. O campo N� Ctro SAP TLF do Sharepoint n�o possui a op��o " + numeroContrato + "\n");
    		}
    		
    		// Campo Contrato (Projeto)
    		String idComboContratoProjeto = inicioIdCampos + "V1_I1_D4";
    		// Estava dando o seguinte erro: org.openqa.selenium.StaleElementReferenceException: stale element reference: element is not attached to the page document
    		// Sempre que voc� enfrentar esse problema, apenas defina o elemento da web mais uma vez acima da linha em que voc� est� obtendo um erro
    		try {
    			wait.until(ExpectedConditions.elementToBeClickable(By.id(idComboContratoProjeto))).click();
    		}
    		catch(org.openqa.selenium.StaleElementReferenceException ex) {
    			wait.until(ExpectedConditions.elementToBeClickable(By.id(idComboContratoProjeto))).click();
    		}
    		WebElement comboContratoProjeto = driver.findElement(By.id(idComboContratoProjeto));
    		// Elementos do combo
    		Select elementosComboContratoProjeto  = new Select(comboContratoProjeto);
    		//String contratoProjeto = "Digital Factory";
    		//elementosComboContratoProjeto.selectByVisibleText(contratoProjeto);
    		int quantidadeElementosComboContratoProjeto = elementosComboContratoProjeto.getOptions().size();
    		// Se tiver mais que um ou nenhum, lan�arei exce��o
    		// Estou comparando com 2 porque todos os combos t�m uma op��o em branco
    		// Tamb�m verifico se a segunda op��o que � a que tem valor n�o est� vazia ou em branco
    		if (quantidadeElementosComboContratoProjeto == 2 && elementosComboContratoProjeto.getOptions().get(1).getText() != null && !elementosComboContratoProjeto.getOptions().get(1).getText().isEmpty()) {
    			elementosComboContratoProjeto.selectByIndex(1);
    			Thread.sleep(5000);
    		} else {
    			throw new Exception("Regra de preenchimento inv�lida. O campo Contrato (Projeto) do Sharepoint possui nenhum ou mais de um valor" + "\n");
    		}
    		
    		// Campo N� Contrato ACN
    		String idComboNumeroContratoACN = inicioIdCampos + "V1_I1_D5";
    		// Estava dando o seguinte erro: org.openqa.selenium.StaleElementReferenceException: stale element reference: element is not attached to the page document
    		// Sempre que voc� enfrentar esse problema, apenas defina o elemento da web mais uma vez acima da linha em que voc� est� obtendo um erro
    		try {
    			wait.until(ExpectedConditions.elementToBeClickable(By.id(idComboNumeroContratoACN))).click();
    		}
    		catch(org.openqa.selenium.StaleElementReferenceException ex) {
    			wait.until(ExpectedConditions.elementToBeClickable(By.id(idComboNumeroContratoACN))).click();
    		}
    		WebElement comboNumeroContratoACN = driver.findElement(By.id(idComboNumeroContratoACN));
    		// Elementos do combo
    		Select elementosComboNumeroContratoACN  = new Select(comboNumeroContratoACN);
    		// Se tiver mais que um ou nenhum, lan�arei exce��o
    		// Estou comparando com 2 porque todos os combos t�m uma op��o em branco
    		// Tamb�m verifico se a segunda op��o que � a que tem valor n�o est� vazia ou em branco
    		int quantidadeElementosComboNumeroContratoACN = elementosComboNumeroContratoACN.getOptions().size();
    		// Estou comparando com 2 porque todos os combos t�m uma op��o em branco
    		// Tamb�m verifico se a segunda op��o que � a que tem valor n�o est� vazia ou em branco
    		if (quantidadeElementosComboNumeroContratoACN == 2 && elementosComboNumeroContratoACN.getOptions().get(1).getText() != null && !elementosComboNumeroContratoACN.getOptions().get(1).getText().isEmpty()) {
    			elementosComboNumeroContratoACN.selectByIndex(1);
    			Thread.sleep(5000);
    		} else {
    			throw new Exception("Regra de preenchimento inv�lida. O campo N� Contrato ACN do Sharepoint possui nenhum ou mais de um valor" + "\n");
    		}
    		
    		// Campo Service Group
    		String idComboServiceGroup = inicioIdCampos + "V1_I1_D6";
    		// Estava dando o seguinte erro: org.openqa.selenium.StaleElementReferenceException: stale element reference: element is not attached to the page document
    		// Sempre que voc� enfrentar esse problema, apenas defina o elemento da web mais uma vez acima da linha em que voc� est� obtendo um erro
    		try {
    			wait.until(ExpectedConditions.elementToBeClickable(By.id(idComboServiceGroup))).click();
    		}
    		catch(org.openqa.selenium.StaleElementReferenceException ex) {
    			wait.until(ExpectedConditions.elementToBeClickable(By.id(idComboServiceGroup))).click();
    		}
    		WebElement comboServiceGroup = driver.findElement(By.id(idComboServiceGroup));
    		// Elementos do combo
    		Select elementosComboServiceGroup  = new Select(comboServiceGroup);
    		int quantidadeElementosComboServiceGroup = elementosComboServiceGroup.getOptions().size();
    		// Se tiver um ou mais que um, seleciono o primeiro
    		// Estou comparando com 2 porque todos os combos t�m uma op��o em branco
    		// Tamb�m verifico se a segunda op��o que � a que tem valor n�o est� vazia ou em branco
    		String mensagemComboServiceGroup = "";
    		if (quantidadeElementosComboServiceGroup >=2 && elementosComboServiceGroup.getOptions().get(1).getText() != null && !elementosComboServiceGroup.getOptions().get(1).getText().isEmpty()) {
    			// Se existirem mais que uma op��o, informo no campo de Observa��o que selecionei o primeiro
    			if (quantidadeElementosComboServiceGroup >=3) {
    				mensagemComboServiceGroup = "O campo Service Group possui mais de um valor, neste caso o robo selecionou a primeira opcao do combo.";
    			}
    			//String serviceGroup = "Application Outsourcing (AO)";
    			elementosComboServiceGroup.selectByIndex(1);
    			Thread.sleep(5000);
    		} else {
    			throw new Exception("Regra de preenchimento invalida. O campo Service Group do Sharepoint possui nenhum valor" + "\n");
    		}
    		
    		// Campo WBS
    		String idComboWBS = inicioIdCampos + "V1_I1_D7";
    		// Estava dando o seguinte erro: org.openqa.selenium.StaleElementReferenceException: stale element reference: element is not attached to the page document
    		// Sempre que voc� enfrentar esse problema, apenas defina o elemento da web mais uma vez acima da linha em que voc� est� obtendo um erro
    		try {
    			wait.until(ExpectedConditions.elementToBeClickable(By.id(idComboWBS))).click();
    		}
    		catch(org.openqa.selenium.StaleElementReferenceException ex) {
    			wait.until(ExpectedConditions.elementToBeClickable(By.id(idComboWBS))).click();
    		}
    		WebElement comboWBS = driver.findElement(By.id(idComboWBS));
    		// Elementos do combo
    		Select elementosComboWBS  = new Select(comboWBS);
    		// String wbs = "AZY71001";
    		// elementosComboWBS.selectByVisibleText(wbs);
    		int quantidadeElementosComboWBS  = elementosComboWBS.getOptions().size();
    		// Se tiver mais que um ou nenhum, lan�arei exce��o
    		// Estou comparando com 2 porque todos os combos t�m uma op��o em branco
    		// Tamb�m verifico se a segunda op��o que � a que tem valor n�o est� vazia ou em branco
    		if (quantidadeElementosComboWBS == 2 && elementosComboWBS.getOptions().get(1).getText() != null && !elementosComboWBS.getOptions().get(1).getText().isEmpty()) {
    			elementosComboWBS.selectByIndex(1);
    			Thread.sleep(5000);
    		} else {
    			throw new Exception("Regra de preenchimento invalida. O campo WBS do Sharepoint possui nenhum ou mais de um valor" + "\n");
    		}
    		Thread.sleep(5000);
    		
    		// Campo Type
    		String idComboType = inicioIdCampos + "V1_I1_D8";
    		// Estava dando o seguinte erro: org.openqa.selenium.StaleElementReferenceException: stale element reference: element is not attached to the page document
    		// Sempre que voc� enfrentar esse problema, apenas defina o elemento da web mais uma vez acima da linha em que voc� est� obtendo um erro
    		try {
    			wait.until(ExpectedConditions.elementToBeClickable(By.id(idComboType))).click();
    		}
    		catch(org.openqa.selenium.StaleElementReferenceException ex) {
    			wait.until(ExpectedConditions.elementToBeClickable(By.id(idComboType))).click();
    		}
    		WebElement comboType = driver.findElement(By.id(idComboType));
    		// Elementos do combo
    		Select elementosComboType = new Select(comboType);
    		// Se tiver mais que um, deixa em branco.
    		// Se tiver um, seleciona ele 
    		int quantidadeElementosComboType = elementosComboType.getOptions().size();
    		// Estou comparando com 2 porque todos os combos t�m uma op��o em branco
    		if (quantidadeElementosComboType == 2) {
    			//String type = "AD";
    			elementosComboType.selectByIndex(1);
    			Thread.sleep(5000);
    		}
    		
    		// Campo Servi�o
    		String idComboServico = inicioIdCampos + "V1_I1_D9";
    		// Estava dando o seguinte erro: org.openqa.selenium.StaleElementReferenceException: stale element reference: element is not attached to the page document
    		// Sempre que voc� enfrentar esse problema, apenas defina o elemento da web mais uma vez acima da linha em que voc� est� obtendo um erro
    		try {
    			wait.until(ExpectedConditions.elementToBeClickable(By.id(idComboServico))).click();
    		}
    		catch(org.openqa.selenium.StaleElementReferenceException ex) {
    			wait.until(ExpectedConditions.elementToBeClickable(By.id(idComboServico))).click();
    		}
    		WebElement comboServico = driver.findElement(By.id(idComboServico));
    		// Elementos do combo
    		Select elementosComboServico  = new Select(comboServico);
    		// Se tiver mais que um, deixa em branco.
    		// Se tiver um, seleciona ele 
    		int quantidadeElementosComboServico  = elementosComboServico.getOptions().size();
    		// Estou comparando com 2 porque todos os combos t�m uma op��o em branco
    		if (quantidadeElementosComboServico == 2) {
    			//String servico = "Desenvolvimento - Capex";
    			elementosComboServico.selectByIndex(1);
    			Thread.sleep(5000);
    		}
    		
    		// Campo Raz�o Social
    		String idComboRazaoSocial = inicioIdCampos + "V1_I1_D10";
    		// Estava dando o seguinte erro: org.openqa.selenium.StaleElementReferenceException: stale element reference: element is not attached to the page document
    		// Sempre que voc� enfrentar esse problema, apenas defina o elemento da web mais uma vez acima da linha em que voc� est� obtendo um erro
    		try {
    			wait.until(ExpectedConditions.elementToBeClickable(By.id(idComboRazaoSocial))).click();
    		}
    		catch(org.openqa.selenium.StaleElementReferenceException ex) {
    			wait.until(ExpectedConditions.elementToBeClickable(By.id(idComboRazaoSocial))).click();
    		}
    		WebElement comboRazaoSocial = driver.findElement(By.id(idComboRazaoSocial));
    		// Elementos do combo
    		Select elementosComboRazaoSocial = new Select(comboRazaoSocial);
    		String razaoSocialSharepointTelefonica = "TELEFÔNICA BRASIL S/A";
    		String razaoSocialSharepointTerraNetworks = "TERRA NETWORKS BRASIL S.A.";
    		String razaoSocialAdquiraTelefonica = "Telefonica Brasil S.A";
    		String razaoSocialAdquiraTerraNetworks1 = "Terra Networks Brasil S.A";
    		String razaoSocialAdquiraTerraNetworks2 = "Terra Network Brasil LTDA";
    		
    		String razaoSocialAdquira = pedido.getComprador();
    		String razaoSocial = "";
    		
    		if (razaoSocialAdquiraTelefonica.equalsIgnoreCase(razaoSocialAdquira)) {
    			razaoSocial = razaoSocialSharepointTelefonica;
    		} else if (razaoSocialAdquiraTerraNetworks1.equalsIgnoreCase(razaoSocialAdquira) || razaoSocialAdquiraTerraNetworks2.equalsIgnoreCase(razaoSocialAdquira)) {
    			razaoSocial = razaoSocialSharepointTerraNetworks;
    		} else {
    			throw new Exception("Regra de preenchimento invalida. A seguinte raz�o social do Adquira esta diferente do Sharepoint : " + razaoSocial + "\n");
    		}
    		
    		elementosComboRazaoSocial.selectByVisibleText(razaoSocial);
    		//elementosComboRazaoSocial.selectByVisibleText("TELEFÔNICA BRASIL S/A");
    		Thread.sleep(5000);
    		
    		// Campo CNPJ
    		String idComboCNPJ = inicioIdCampos + "V1_I1_D11";
    		// Estava dando o seguinte erro: org.openqa.selenium.StaleElementReferenceException: stale element reference: element is not attached to the page document
    		// Sempre que voc� enfrentar esse problema, apenas defina o elemento da web mais uma vez acima da linha em que voc� est� obtendo um erro
    		try {
    			wait.until(ExpectedConditions.elementToBeClickable(By.id(idComboCNPJ))).click();
    		}
    		catch(org.openqa.selenium.StaleElementReferenceException ex) {
    			wait.until(ExpectedConditions.elementToBeClickable(By.id(idComboCNPJ))).click();
    		}
    		WebElement comboCNPJ = driver.findElement(By.id(idComboCNPJ));
    		// Elementos do combo
    		Select elementosComboCNPJ = new Select(comboCNPJ);
    		//String cnpj = "02.558.157/0001-62";
    		String cnpj = pedido.getCnpjCliente();
    		boolean existeCNPJ = false;
    		// Verifico se existe o status no combo
    		// Se n�o existir, lan�o exce��o
    		int quantidadeElementosComboCNPJ  = elementosComboCNPJ.getOptions().size();
    		if (quantidadeElementosComboCNPJ > 0) {
    			for (WebElement elemento : elementosComboCNPJ.getOptions()) {
    				if (cnpj.equalsIgnoreCase(elemento.getText().trim())) {
    					existeCNPJ = true;
    					break;
    				}
    			}
    		}
    		if (existeCNPJ) {
    			elementosComboCNPJ.selectByVisibleText(cnpj);
    			Thread.sleep(5000);
    		} else {
    			throw new Exception("Regra de preenchimento invalida. O campo CNPJ do Sharepoint nao possui a opcao " + cnpj + "\n");
    		}
    		
    		// Campo Endere�o
    		String idComboEndereco = inicioIdCampos + "V1_I1_D12";
    		// Estava dando o seguinte erro: org.openqa.selenium.StaleElementReferenceException: stale element reference: element is not attached to the page document
    		// Sempre que voc� enfrentar esse problema, apenas defina o elemento da web mais uma vez acima da linha em que voc� est� obtendo um erro
    		try {
    			wait.until(ExpectedConditions.elementToBeClickable(By.id(idComboEndereco))).click();
    		}
    		catch(org.openqa.selenium.StaleElementReferenceException ex) {
    			wait.until(ExpectedConditions.elementToBeClickable(By.id(idComboEndereco))).click();
    		}
    		WebElement comboEndereco = driver.findElement(By.id(idComboEndereco));
    		// Elementos do combo
    		Select elementosComboEndereco = new Select(comboEndereco);
    		// String endereco = "Av. Engenheiro Luiz Carlos Berrini, 1376 - Cidade Mon��es - S�o Paulo - SP - CEP: 04571-936";
    		// elementosComboEndereco.selectByVisibleText(endereco);
    		int quantidadeElementosComboEndereco  = elementosComboEndereco.getOptions().size();
    		// Se tiver mais que um ou nenhum, lan�arei exce��o
    		// Estou comparando com 2 porque todos os combos t�m uma op��o em branco
    		// Tamb�m verifico se a segunda op��o que � a que tem valor n�o est� vazia ou em branco
    		if (quantidadeElementosComboEndereco == 2 && elementosComboEndereco.getOptions().get(1).getText() != null && !elementosComboEndereco.getOptions().get(1).getText().isEmpty()) {
    			elementosComboEndereco.selectByIndex(1);
    			Thread.sleep(5000);
    		} else {
    			throw new Exception("Regra de preenchimento invalida. O campo Endereco do Sharepoint possui nenhum ou mais de um valor" + "\n");
    		}
    		
    		// Campo N� Cliente
    		String idComboNumeroCliente = inicioIdCampos + "V1_I1_D13";
    		// Estava dando o seguinte erro: org.openqa.selenium.StaleElementReferenceException: stale element reference: element is not attached to the page document
    		// Sempre que voc� enfrentar esse problema, apenas defina o elemento da web mais uma vez acima da linha em que voc� est� obtendo um erro
    		try {
    			wait.until(ExpectedConditions.elementToBeClickable(By.id(idComboNumeroCliente))).click();
    		}
    		catch(org.openqa.selenium.StaleElementReferenceException ex) {
    			wait.until(ExpectedConditions.elementToBeClickable(By.id(idComboNumeroCliente))).click();
    		}
    		WebElement comboNumeroCliente = driver.findElement(By.id(idComboNumeroCliente));
    		// Elementos do combo
    		Select elementosComboNumeroCliente = new Select(comboNumeroCliente);
    		// String numeroCliente = "10004184";
    		// elementosComboNumeroCliente.selectByVisibleText(numeroCliente);
    		int quantidadeElementosComboNumeroCliente  = elementosComboNumeroCliente.getOptions().size();
    		// Se tiver mais que um ou nenhum, lan�arei exce��o
    		// Estou comparando com 2 porque todos os combos t�m uma op��o em branco
    		// Tamb�m verifico se a segunda op��o que � a que tem valor n�o est� vazia ou em branco
    		if (quantidadeElementosComboNumeroCliente == 2 && elementosComboNumeroCliente.getOptions().get(1).getText() != null && !elementosComboNumeroCliente.getOptions().get(1).getText().isEmpty()) {
    			elementosComboNumeroCliente.selectByIndex(1);
    			Thread.sleep(5000);
    		} else {
    			throw new Exception("Regra de preenchimento invalida. O campo Numero Cliente do Sharepoint possui nenhum ou mais de um valor" + "\n");
    		}
    		
    		// Campo Portal
    		String idComboPortal = inicioIdCampos + "V1_I1_D14";
    		// Estava dando o seguinte erro: org.openqa.selenium.StaleElementReferenceException: stale element reference: element is not attached to the page document
    		// Sempre que voc� enfrentar esse problema, apenas defina o elemento da web mais uma vez acima da linha em que voc� est� obtendo um erro
    		try {
    			wait.until(ExpectedConditions.elementToBeClickable(By.id(idComboPortal))).click();
    		}
    		catch(org.openqa.selenium.StaleElementReferenceException ex) {
    			wait.until(ExpectedConditions.elementToBeClickable(By.id(idComboPortal))).click();
    		}
    		WebElement comboPortal = driver.findElement(By.id(idComboPortal));
    		// Elementos do combo
    		Select elementosComboPortal = new Select(comboPortal);
    		// Sempre ser� esse valor, pois � de onde o rob� est� buscando as informa��es
    		String portal = "Adquira";
    		boolean existePortal = false;
    		// Verifico se existe o status no combo
    		// Se n�o existir, lan�o exce��o
    		int quantidadeElementosComboPortal  = elementosComboPortal.getOptions().size();
    		if (quantidadeElementosComboPortal > 0) {
    			for (WebElement elemento : elementosComboPortal.getOptions()) {
    				if (portal.equalsIgnoreCase(elemento.getText().trim())) {
    					existePortal = true;
    					break;
    				}
    			}
    		}
    		if (existePortal) {
    			elementosComboPortal.selectByVisibleText(portal);
    			Thread.sleep(5000);
    		} else {
    			throw new Exception("Regra de preenchimento invalida. O campo Portal do Sharepoint nao possui a opcao " + portal + "\n");
    		}
    		
    		// Campo Imposto
    		// O campo imposto j� aparece preenchido com a informa��o default Taxa 2
    		/*
     	        String idComboImposto = inicioIdCampos + "V1_I1_CB15_textBox";
     	        // Estava dando o seguinte erro: org.openqa.selenium.StaleElementReferenceException: stale element reference: element is not attached to the page document
     	        // Sempre que voc� enfrentar esse problema, apenas defina o elemento da web mais uma vez acima da linha em que voc� est� obtendo um erro
     	        try {
     	        	wait.until(ExpectedConditions.elementToBeClickable(By.id(idComboImposto))).click();
     	        }
     	        catch(org.openqa.selenium.StaleElementReferenceException ex) {
     	        	wait.until(ExpectedConditions.elementToBeClickable(By.id(idComboImposto))).click();
     	        }
     	        WebElement comboImposto = driver.findElement(By.id(idComboImposto));
     	        // Elementos do combo
     	        Select elementosComboImposto = new Select(comboImposto);
     	        String imposto = "Taxa 2";
     	        elementosComboImposto.selectByVisibleText(imposto);
     	        Thread.sleep(5000);
    		 */
    		
    		// Campo Org. Vendas
    		String idComboOrgVendas = inicioIdCampos + "V1_I1_D16";
    		// Estava dando o seguinte erro: org.openqa.selenium.StaleElementReferenceException: stale element reference: element is not attached to the page document
    		// Sempre que voc� enfrentar esse problema, apenas defina o elemento da web mais uma vez acima da linha em que voc� est� obtendo um erro
    		try {
    			wait.until(ExpectedConditions.elementToBeClickable(By.id(idComboOrgVendas))).click();
    		}
    		catch(org.openqa.selenium.StaleElementReferenceException ex) {
    			wait.until(ExpectedConditions.elementToBeClickable(By.id(idComboOrgVendas))).click();
    		}
    		WebElement comboOrgVendas = driver.findElement(By.id(idComboOrgVendas));
    		// Elementos do combo
    		Select elementosComboOrgVendas = new Select(comboOrgVendas);
    		// String orgVendas = "1500 - San Pablo";
    		// elementosComboOrgVendas.selectByVisibleText(orgVendas);
    		int quantidadeElementosComboOrgVendas  = elementosComboOrgVendas.getOptions().size();
    		// Se tiver mais que um ou nenhum, lan�arei exce��o
    		// Estou comparando com 2 porque todos os combos t�m uma op��o em branco
    		// Tamb�m verifico se a segunda op��o que � a que tem valor n�o est� vazia ou em branco
    		if (quantidadeElementosComboOrgVendas == 2 && elementosComboOrgVendas.getOptions().get(1).getText() != null && !elementosComboOrgVendas.getOptions().get(1).getText().isEmpty()) {
    			elementosComboOrgVendas.selectByIndex(1);
    			Thread.sleep(5000);
    		} else {
    			throw new Exception("Regra de preenchimento invalida. O campo Org. Vendas do Sharepoint possui nenhum ou mais de um valor" + "\n");
    		}
    		
    		// Campo C�digo Empresa
    		String idComboCodigoEmpresa = inicioIdCampos + "V1_I1_D17";
    		// Estava dando o seguinte erro: org.openqa.selenium.StaleElementReferenceException: stale element reference: element is not attached to the page document
    		// Sempre que voc� enfrentar esse problema, apenas defina o elemento da web mais uma vez acima da linha em que voc� est� obtendo um erro
    		try {
    			wait.until(ExpectedConditions.elementToBeClickable(By.id(idComboCodigoEmpresa))).click();
    		}
    		catch(org.openqa.selenium.StaleElementReferenceException ex) {
    			wait.until(ExpectedConditions.elementToBeClickable(By.id(idComboCodigoEmpresa))).click();
    		}
    		WebElement comboCodigoEmpresa = driver.findElement(By.id(idComboCodigoEmpresa));
    		// Elementos do combo
    		Select elementosComboCodigoEmpresa = new Select(comboCodigoEmpresa);
    		// String codigoEmpresa = "1500 - San Pablo";
    		// elementosComboCodigoEmpresa.selectByVisibleText(codigoEmpresa);
    		int quantidadeElementosComboCodigoEmpresa  = elementosComboCodigoEmpresa.getOptions().size();
    		// Se tiver mais que um ou nenhum, lan�arei exce��o
    		// Estou comparando com 2 porque todos os combos t�m uma op��o em branco
    		// Tamb�m verifico se a segunda op��o que � a que tem valor n�o est� vazia ou em branco
    		if (quantidadeElementosComboCodigoEmpresa == 2 && elementosComboCodigoEmpresa.getOptions().get(1).getText() != null && !elementosComboCodigoEmpresa.getOptions().get(1).getText().isEmpty()) {
    			elementosComboCodigoEmpresa.selectByIndex(1);
    			Thread.sleep(5000);
    		} else {
    			throw new Exception("Regra de preenchimento invalida. O campo Codigo Empresa do Sharepoint possui nenhum ou mais de um valor" + "\n");
    		}
    		
    		// Campo C�digo Material
    		String idComboCodigoMaterial = inicioIdCampos + "V1_I1_D18";
    		// Estava dando o seguinte erro: org.openqa.selenium.StaleElementReferenceException: stale element reference: element is not attached to the page document
    		// Sempre que voc� enfrentar esse problema, apenas defina o elemento da web mais uma vez acima da linha em que voc� est� obtendo um erro
    		try {
    			wait.until(ExpectedConditions.elementToBeClickable(By.id(idComboCodigoMaterial))).click();
    		}
    		catch(org.openqa.selenium.StaleElementReferenceException ex) {
    			wait.until(ExpectedConditions.elementToBeClickable(By.id(idComboCodigoMaterial))).click();
    		}
    		WebElement comboCodigoMaterial = driver.findElement(By.id(idComboCodigoMaterial));
    		// Elementos do combo
    		Select elementosComboCodigoMaterial = new Select(comboCodigoMaterial);
    		// String codigoMaterial = "J0104";
    		// elementosComboCodigoMaterial.selectByVisibleText(codigoMaterial);
    		int quantidadeElementosComboCodigoMaterial  = elementosComboCodigoMaterial.getOptions().size();
    		// Se tiver mais que um ou nenhum, lan�arei exce��o
    		// Estou comparando com 2 porque todos os combos t�m uma op��o em branco
    		// Tamb�m verifico se a segunda op��o que � a que tem valor n�o est� vazia ou em branco
    		if (quantidadeElementosComboCodigoMaterial == 2 && elementosComboCodigoMaterial.getOptions().get(1).getText() != null && !elementosComboCodigoMaterial.getOptions().get(1).getText().isEmpty()) {
    			elementosComboCodigoMaterial.selectByIndex(1);
    			Thread.sleep(5000);
    		} else {
    			throw new Exception("Regra de preenchimento invalida. O campo Codigo Material do Sharepoint possui nenhum ou mais de um valor" + "\n");
    		}
    		
    		// Campo ISS
    		String idComboISS = inicioIdCampos + "V1_I1_D19";
    		// Estava dando o seguinte erro: org.openqa.selenium.StaleElementReferenceException: stale element reference: element is not attached to the page document
    		// Sempre que voc� enfrentar esse problema, apenas defina o elemento da web mais uma vez acima da linha em que voc� est� obtendo um erro
    		try {
    			wait.until(ExpectedConditions.elementToBeClickable(By.id(idComboISS))).click();
    		}
    		catch(org.openqa.selenium.StaleElementReferenceException ex) {
    			wait.until(ExpectedConditions.elementToBeClickable(By.id(idComboISS))).click();
    		}
    		WebElement comboISS = driver.findElement(By.id(idComboISS));
    		// Elementos do combo
    		Select elementosComboISS = new Select(comboISS);
    		// String iss = "0.029";
    		// elementosComboISS.selectByVisibleText(iss);
    		int quantidadeElementosComboISS  = elementosComboISS.getOptions().size();
    		// Se tiver mais que um ou nenhum, lan�arei exce��o
    		// Estou comparando com 2 porque todos os combos t�m uma op��o em branco
    		// Tamb�m verifico se a segunda op��o que � a que tem valor n�o est� vazia ou em branco
    		if (quantidadeElementosComboISS == 2 && elementosComboISS.getOptions().get(1).getText() != null && !elementosComboISS.getOptions().get(1).getText().isEmpty()) {
    			elementosComboISS.selectByIndex(1);
    			Thread.sleep(5000);
    		} else {
    			throw new Exception("Regra de preenchimento invalida. O campo ISS do Sharepoint possui nenhum ou mais de um valor" + "\n");
    		}
    		
    		// Campo PIS
    		String idComboPIS = inicioIdCampos + "V1_I1_D20";
    		// Estava dando o seguinte erro: org.openqa.selenium.StaleElementReferenceException: stale element reference: element is not attached to the page document
    		// Sempre que voc� enfrentar esse problema, apenas defina o elemento da web mais uma vez acima da linha em que voc� est� obtendo um erro
    		try {
    			wait.until(ExpectedConditions.elementToBeClickable(By.id(idComboPIS))).click();
    		}
    		catch(org.openqa.selenium.StaleElementReferenceException ex) {
    			wait.until(ExpectedConditions.elementToBeClickable(By.id(idComboPIS))).click();
    		}
    		WebElement comboPIS = driver.findElement(By.id(idComboPIS));
    		// Elementos do combo
    		Select elementosComboPIS = new Select(comboPIS);
    		// String pis = "0.0065";
    		// elementosComboPIS.selectByVisibleText(pis);
    		int quantidadeElementosComboPIS  = elementosComboPIS.getOptions().size();
    		// Se tiver mais que um ou nenhum, lan�arei exce��o
    		// Estou comparando com 2 porque todos os combos t�m uma op��o em branco
    		// Tamb�m verifico se a segunda op��o que � a que tem valor n�o est� vazia ou em branco
    		if (quantidadeElementosComboPIS == 2 && elementosComboPIS.getOptions().get(1).getText() != null && !elementosComboPIS.getOptions().get(1).getText().isEmpty()) {
    			elementosComboPIS.selectByIndex(1);
    			Thread.sleep(5000);
    		} else {
    			throw new Exception("Regra de preenchimento invalida. O campo PIS do Sharepoint possui nenhum ou mais de um valor" + "\n");
    		}
    		
    		// Campo Cofins
    		String idComboCofins = inicioIdCampos + "V1_I1_D21";
    		// Estava dando o seguinte erro: org.openqa.selenium.StaleElementReferenceException: stale element reference: element is not attached to the page document
    		// Sempre que voc� enfrentar esse problema, apenas defina o elemento da web mais uma vez acima da linha em que voc� est� obtendo um erro
    		try {
    			wait.until(ExpectedConditions.elementToBeClickable(By.id(idComboCofins))).click();
    		}
    		catch(org.openqa.selenium.StaleElementReferenceException ex) {
    			wait.until(ExpectedConditions.elementToBeClickable(By.id(idComboCofins))).click();
    		}
    		WebElement comboCofins = driver.findElement(By.id(idComboCofins));
    		// Elementos do combo
    		Select elementosComboCofins = new Select(comboCofins);
    		// String cofins = "0.03";
    		// elementosComboCofins.selectByVisibleText(cofins);
    		int quantidadeElementosComboCofins  = elementosComboCofins.getOptions().size();
    		// Se tiver mais que um ou nenhum, lan�arei exce��o
    		// Estou comparando com 2 porque todos os combos t�m uma op��o em branco
    		// Tamb�m verifico se a segunda op��o que � a que tem valor n�o est� vazia ou em branco
    		if (quantidadeElementosComboCofins == 2 && elementosComboCofins.getOptions().get(1).getText() != null && !elementosComboCofins.getOptions().get(1).getText().isEmpty()) {
    			elementosComboCofins.selectByIndex(1);
    			Thread.sleep(5000);
    		} else {
    			throw new Exception("Regra de preenchimento invalida. O campo Cofins do Sharepoint possui nenhum ou mais de um valor" + "\n");
    		}
    		
    		// Campo C�digo Taxa
    		String idComboCodigoTaxa = inicioIdCampos + "V1_I1_D22";
    		// Estava dando o seguinte erro: org.openqa.selenium.StaleElementReferenceException: stale element reference: element is not attached to the page document
    		// Sempre que voc� enfrentar esse problema, apenas defina o elemento da web mais uma vez acima da linha em que voc� est� obtendo um erro
    		try {
    			wait.until(ExpectedConditions.elementToBeClickable(By.id(idComboCodigoTaxa))).click();
    		}
    		catch(org.openqa.selenium.StaleElementReferenceException ex) {
    			wait.until(ExpectedConditions.elementToBeClickable(By.id(idComboCodigoTaxa))).click();
    		}
    		WebElement comboCodigoTaxa = driver.findElement(By.id(idComboCodigoTaxa));
    		// Elementos do combo
    		Select elementosComboCodigoTaxa = new Select(comboCodigoTaxa);
    		// String codigoTaxa = "H0";
    		// elementosComboCodigoTaxa.selectByVisibleText(codigoTaxa);
    		int quantidadeElementosComboCodigoTaxa  = elementosComboCodigoTaxa.getOptions().size();
    		// Se tiver mais que um ou nenhum, lan�arei exce��o
    		// Estou comparando com 2 porque todos os combos t�m uma op��o em branco
    		// Tamb�m verifico se a segunda op��o que � a que tem valor n�o est� vazia ou em branco
    		if (quantidadeElementosComboCodigoTaxa == 2 && elementosComboCodigoTaxa.getOptions().get(1).getText() != null && !elementosComboCodigoTaxa.getOptions().get(1).getText().isEmpty()) {
    			elementosComboCodigoTaxa.selectByIndex(1);
    			Thread.sleep(5000);
    		} else {
    			throw new Exception("Regra de preenchimento invalida. O campo Codigo Taxa do Sharepoint possui nenhum ou mais de um valor" + "\n");
    		}
    		
    		// Campo Total VAT
    		// Campo em branco, n�o permite preenchimento
    		
    		// Campo PO/Pedido
    		String idPoPedido = inicioIdCampos + "V1_I1_T24";
    		// Estava dando o seguinte erro: org.openqa.selenium.StaleElementReferenceException: stale element reference: element is not attached to the page document
    		// Sempre que voc� enfrentar esse problema, apenas defina o elemento da web mais uma vez acima da linha em que voc� est� obtendo um erro
    		try {
    			wait.until(ExpectedConditions.elementToBeClickable(By.id(idPoPedido))).click();
    		}
    		catch(org.openqa.selenium.StaleElementReferenceException ex) {
    			wait.until(ExpectedConditions.elementToBeClickable(By.id(idPoPedido))).click();
    		}
    		//String inputTextPoPedido = "123";
    		String inputTextPoPedido = pedido.getNumero();
    		WebElement poPedido = driver.findElement(By.id(idPoPedido));
    		String comandoJsPoPedido = "arguments[0].setAttribute('value','"+inputTextPoPedido+"')";
    		
    		try {
    			js.executeScript(comandoJsPoPedido, poPedido); 
    			Thread.sleep(500);
    		} catch (Exception e) {
    			System.out.println("Passei: " + "poPedido");
    			poPedido.sendKeys(inputTextPoPedido);
    			Thread.sleep(500);
    		}
    		Thread.sleep(5000);
    		
    		// Campo Recebimento PO
    		String idRecebimentoPO = inicioIdCampos + "V1_I1_T27";
    		// Estava dando o seguinte erro: org.openqa.selenium.StaleElementReferenceException: stale element reference: element is not attached to the page document
    		// Sempre que voc� enfrentar esse problema, apenas defina o elemento da web mais uma vez acima da linha em que voc� est� obtendo um erro
    		try {
    			wait.until(ExpectedConditions.elementToBeClickable(By.id(idRecebimentoPO))).click();
    		}
    		catch(org.openqa.selenium.StaleElementReferenceException ex) {
    			wait.until(ExpectedConditions.elementToBeClickable(By.id(idRecebimentoPO))).click();
    		}
    		WebElement recebimentoPO = driver.findElement(By.id(idRecebimentoPO));
    		String comandoJsRecebimentoPO = "arguments[0].setAttribute('value','"+dataAtualSharepoint+"')";
    		
    		try {
    			js.executeScript(comandoJsRecebimentoPO, recebimentoPO); 
    			Thread.sleep(500);
    		} catch (Exception e) {
    			System.out.println("Passei: " + "recebimentoPO");
    			recebimentoPO.sendKeys(dataAtualSharepoint);
    			Thread.sleep(500);
    		}
    		Thread.sleep(5000);
    		
    		// Campo Prazo de Pgto
    		String idComboPrazoDePgto = inicioIdCampos + "V1_I1_D29";
    		// Estava dando o seguinte erro: org.openqa.selenium.StaleElementReferenceException: stale element reference: element is not attached to the page document
    		// Sempre que voc� enfrentar esse problema, apenas defina o elemento da web mais uma vez acima da linha em que voc� est� obtendo um erro
    		try {
    			wait.until(ExpectedConditions.elementToBeClickable(By.id(idComboPrazoDePgto))).click();
    		}
    		catch(org.openqa.selenium.StaleElementReferenceException ex) {
    			wait.until(ExpectedConditions.elementToBeClickable(By.id(idComboPrazoDePgto))).click();
    		}
    		WebElement comboPrazoDePgto = driver.findElement(By.id(idComboPrazoDePgto));
    		// Elementos do combo
    		Select elementosComboPrazoDePgto  = new Select(comboPrazoDePgto);
    		// String prazoDePgto  = "75";
    		// elementosComboPrazoDePgto.selectByVisibleText(prazoDePgto);
    		int quantidadeElementosComboPrazoDePgto  = elementosComboPrazoDePgto.getOptions().size();
    		// Se tiver mais que um ou nenhum, lan�arei exce��o
    		// Estou comparando com 2 porque todos os combos t�m uma op��o em branco
    		// Tamb�m verifico se a segunda op��o que � a que tem valor n�o est� vazia ou em branco
    		if (quantidadeElementosComboPrazoDePgto == 2 && elementosComboPrazoDePgto.getOptions().get(1).getText() != null && !elementosComboPrazoDePgto.getOptions().get(1).getText().isEmpty()) {
    			elementosComboPrazoDePgto.selectByIndex(1);
    			Thread.sleep(5000);
    		} else {
    			throw new Exception("Regra de preenchimento invalida. O campo Prazo de Pgto do Sharepoint possui nenhum ou mais de um valor" + "\n");
    		}
    		
    		// Campo CAP
    		String idCap = inicioIdCampos + "V1_I1_T30";
    		clickNoCampoCaps(wait, idCap);
    		//String inputTextCap = "CAP";
    		String inputTextCap = recuperaCap(pedido);
    		boolean preencheuCapNaObservacao = false;
    		
    		if (inputTextCap.length() <= 255) {
    			
    			preencheCampoCap(driver, wait, js, idCap, inputTextCap);
    			
    		} else if (inputTextCap.length() > 255 && inputTextCap.length() <= 510) {
    			
    			// O campo CAP s� aceita 255 caracteres, ent�o vou verificar o tamanho e se passar, insiro o restante no campo Observa��o
    			String primeiraParteCap = inputTextCap.substring(0, 255);
    			preencheCampoCap(driver, wait, js, idCap, primeiraParteCap);
    			
    			String segundaParteCap = inputTextCap.substring(255, inputTextCap.length());
    			String idObservacao = inicioIdCampos + "V1_I1_T32";
    			preencheuCapNaObservacao = true;
    			preencheCampoObservacao(driver, wait, js, idObservacao, segundaParteCap);
    			pedido.setObservacaoSharepoint("Os CAPs deste pedido excederam 255 caracteres.Parte destes CAPs estao no campo CAP e outra parte est� no campo Observacao do Sharepoint. " + "CAP completo: " + inputTextCap);
    			
    		} else {
    			// Se o Cap for maior que 510, insiro 255 no Cap e os outros 255 no campo Observa��o.
    			// Como ele � maior que 510 perderemos o restante da informa��o, ent�o vou mostrar o Cap completo na base de dados
    			String primeiraParteCap = inputTextCap.substring(0, 255);
    			preencheCampoCap(driver, wait, js, idCap, primeiraParteCap);
    			
    			String segundaParteCap = inputTextCap.substring(255, 510);
    			String idObservacao = inicioIdCampos + "V1_I1_T32";
    			preencheuCapNaObservacao = true;
    			preencheCampoObservacao(driver, wait, js, idObservacao, segundaParteCap);
    			pedido.setObservacaoSharepoint("Os CAPs deste pedido excederam 510 caracteres.Parte destes CAPs estao no campo CAP e outra parte esta no campo Observacao do Sharepoint. " + "CAP completo: " + inputTextCap);
    			
    		}
    		
    		// Campo OI
    		// N�o precisa preencher
    		/*
    			String idOi = inicioIdCampos + "V1_I1_T31";
     	        // Estava dando o seguinte erro: org.openqa.selenium.StaleElementReferenceException: stale element reference: element is not attached to the page document
     	        // Sempre que voc� enfrentar esse problema, apenas defina o elemento da web mais uma vez acima da linha em que voc� est� obtendo um erro
     	        try {
     	        	wait.until(ExpectedConditions.elementToBeClickable(By.id(idOi))).click();
     	        }
     	        catch(org.openqa.selenium.StaleElementReferenceException ex) {
     	        	wait.until(ExpectedConditions.elementToBeClickable(By.id(idOi))).click();
     	        }
    			String inputTextOi = "Oi";
    			WebElement oi = driver.findElement(By.id(idOi));
    			String comandoJsOi = "arguments[0].setAttribute('value','"+inputTextOi+"')";
    			
    			try {
    				js.executeScript(comandoJsOi, oi); 
    				Thread.sleep(500);
    			} catch (Exception e) {
    				System.out.println("Passei: " + "Oi");
    				oi.sendKeys(inputTextOi);
    				Thread.sleep(500);
    			}
    			Thread.sleep(5000);
    		 */
    		
    		// Campo Observa��o
    		String idObservacao = inicioIdCampos + "V1_I1_T32";
    		// N�o sei porque, mas mesmo preenchendo o campo CAPS o sharepoint estava dando erro quando eu tentava salvar.
    		// O erro � como se o campo CAPS estivesse vazio.
    		// Ent�o percebi que se eu clicasse em outro campo e depois clicasse no campo CAPS de novo e novamente clicasse em outro campo, o erro sumia
    		clickNoCampoObservacao(wait, idObservacao);
    		Thread.sleep(1000);
    		clickNoCampoCaps(wait, idCap);
    		Thread.sleep(1000);
    		clickNoCampoObservacao(wait, idObservacao);
    		Thread.sleep(1000);
    		String inputTextObservacao = "";
    		if (mensagemComboServiceGroup != null && !mensagemComboServiceGroup.isEmpty()) {
    			inputTextObservacao = mensagemComboServiceGroup;
    		}
    		
    		if (!preencheuCapNaObservacao) {
    			pedido.setObservacaoSharepoint(inputTextObservacao);
    			preencheCampoObservacao(driver, wait, js, idObservacao, inputTextObservacao);
    		}
    		
    		// Campo Item
    		List<Integer> listaItemAdquira = pedido.getListaItem();
    		if (listaItemAdquira != null && !listaItemAdquira.isEmpty()) {
    			
    			for (Integer itemAdquira : listaItemAdquira) {
    				
    				String idItem = "//input[@value='" + itemAdquira + "']";
    				
    				try {
    					// Caso o n�mero de itens do pedido ultrapasse os itens suportados pelo Sharepoint, entra no catch para n�o sair do preenchimento do resto do pedido
    					WebDriverWait waitItem = new WebDriverWait(driver, 5);
    					waitItem.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(idItem)));
    				} catch (Exception e) {
    					continue;
    				}
    				
    				WebElement item = driver.findElement(By.xpath(idItem));
    				if (!item.isSelected ()) {					
    					// Estava dando o seguinte erro: org.openqa.selenium.StaleElementReferenceException: stale element reference: element is not attached to the page document
    					// Sempre que voc� enfrentar esse problema, apenas defina o elemento da web mais uma vez acima da linha em que voc� est� obtendo um erro
    					try {
    						wait.until(ExpectedConditions.elementToBeClickable(By.xpath(idItem))).click();
    						Thread.sleep(1000);
    					}
    					catch(org.openqa.selenium.StaleElementReferenceException ex) {
    						wait.until(ExpectedConditions.elementToBeClickable(By.xpath(idItem))).click();
    						Thread.sleep(1000);
    					}
    					
    				}
    				
    			}
    			
    		}
    		Thread.sleep(5000);
    		
    		// Campo Nota Fiscal
    		// N�o precisa preencher
    		/*
    			String idNotaFiscal = inicioIdCampos + "V1_I1_T33";
     	        // Estava dando o seguinte erro: org.openqa.selenium.StaleElementReferenceException: stale element reference: element is not attached to the page document
     	        // Sempre que voc� enfrentar esse problema, apenas defina o elemento da web mais uma vez acima da linha em que voc� est� obtendo um erro
     	        try {
     	        	wait.until(ExpectedConditions.elementToBeClickable(By.id(idNotaFiscal))).click();
     	        }
     	        catch(org.openqa.selenium.StaleElementReferenceException ex) {
     	        	wait.until(ExpectedConditions.elementToBeClickable(By.id(idNotaFiscal))).click();
     	        }
    			String inputTextNotaFiscal = "Nota Fiscal";
    			WebElement notaFiscal = driver.findElement(By.id(idNotaFiscal));
    			String comandoJsNotaFiscal = "arguments[0].setAttribute('value','"+inputTextNotaFiscal+"')";
    			
    			try {
    				js.executeScript(comandoJsNotaFiscal, notaFiscal); 
    				Thread.sleep(500);
    			} catch (Exception e) {
    				System.out.println("Passei: " + "Oi");
    				notaFiscal.sendKeys(inputTextNotaFiscal);
    				Thread.sleep(500);
    			}
    			Thread.sleep(5000);
    		 */
    		
    		// Campo DMR
    		// N�o precisa preencher
    		/*
    			String idDmr = inicioIdCampos + "V1_I1_T34";
     	        // Estava dando o seguinte erro: org.openqa.selenium.StaleElementReferenceException: stale element reference: element is not attached to the page document
     	        // Sempre que voc� enfrentar esse problema, apenas defina o elemento da web mais uma vez acima da linha em que voc� est� obtendo um erro
     	        try {
     	        	wait.until(ExpectedConditions.elementToBeClickable(By.id(idDmr))).click();
     	        }
     	        catch(org.openqa.selenium.StaleElementReferenceException ex) {
     	        	wait.until(ExpectedConditions.elementToBeClickable(By.id(idDmr))).click();
     	        }
    			String inputTextDmr = "DMR";
    			WebElement dmr = driver.findElement(By.id(idDmr));
    			String comandoJsDmr = "arguments[0].setAttribute('value','"+inputTextDmr+"')";
    			
    			try {
    				js.executeScript(comandoJsDmr, dmr); 
    				Thread.sleep(500);
    			} catch (Exception e) {
    				System.out.println("Passei: " + "Oi");
    				dmr.sendKeys(inputTextDmr);
    				Thread.sleep(500);
    			}
    			Thread.sleep(5000);
    		 */
    		
    		// Campo Data da Nota
    		// N�o vamos preencher por enquanto
    		/*
    			String idDataDaNota = inicioIdCampos + "V1_I1_T35";
     	        // Estava dando o seguinte erro: org.openqa.selenium.StaleElementReferenceException: stale element reference: element is not attached to the page document
     	        // Sempre que voc� enfrentar esse problema, apenas defina o elemento da web mais uma vez acima da linha em que voc� est� obtendo um erro
     	        try {
     	        	wait.until(ExpectedConditions.elementToBeClickable(By.id(idDataDaNota))).click();
     	        }
     	        catch(org.openqa.selenium.StaleElementReferenceException ex) {
     	        	wait.until(ExpectedConditions.elementToBeClickable(By.id(idDataDaNota))).click();
     	        }
     	        String inputTextDataDaNota = dataDaNotaSharepoint(new Date());
     	        WebElement dataDaNota = driver.findElement(By.id(idDataDaNota));
    			String comandoJsDataDaNota = "arguments[0].setAttribute('value','"+inputTextDataDaNota+"')";
    			
    			try {
    				js.executeScript(comandoJsDataDaNota, dataDaNota); 
    				Thread.sleep(500);
    			} catch (Exception e) {
    				System.out.println("Passei: " + "recebimentoPO");
    				dataDaNota.sendKeys(inputTextDataDaNota);
    				Thread.sleep(500);
    			}
    			Thread.sleep(5000);
    		 */
    		
    		// Campo Valor NFE
    		String idValorNFE = inicioIdCampos + "V1_I1_T37";
    		// Estava dando o seguinte erro: org.openqa.selenium.StaleElementReferenceException: stale element reference: element is not attached to the page document
    		// Sempre que voc� enfrentar esse problema, apenas defina o elemento da web mais uma vez acima da linha em que voc� est� obtendo um erro
    		try {
    			wait.until(ExpectedConditions.elementToBeClickable(By.id(idValorNFE))).click();
    		}
    		catch(org.openqa.selenium.StaleElementReferenceException ex) {
    			wait.until(ExpectedConditions.elementToBeClickable(By.id(idValorNFE))).click();
    		}
    		String inputTextValorNFE = pedido.getValor();
    		WebElement valorNFE = driver.findElement(By.id(idValorNFE));
    		String comandoJsValorNFE = "arguments[0].setAttribute('value','"+inputTextValorNFE+"')";
    		
    		try {
    			js.executeScript(comandoJsValorNFE, valorNFE); 
    			Thread.sleep(500);
    		} catch (Exception e) {
    			System.out.println("Passei: " + "valorNFE");
    			valorNFE.sendKeys(inputTextValorNFE);
    			Thread.sleep(500);
    		}
    		
    		// Campo Pgto. Previsto
    		// N�o precisa preencher
    		
    		// Campo Data do Pgto.
    		// N�o precisa preencher
    		/*
    			String idDataDoPgto = inicioIdCampos + "V1_I1_T39";
     	        // Estava dando o seguinte erro: org.openqa.selenium.StaleElementReferenceException: stale element reference: element is not attached to the page document
     	        // Sempre que voc� enfrentar esse problema, apenas defina o elemento da web mais uma vez acima da linha em que voc� est� obtendo um erro
     	        try {
     	        	wait.until(ExpectedConditions.elementToBeClickable(By.id(idDataDoPgto))).click();
     	        }
     	        catch(org.openqa.selenium.StaleElementReferenceException ex) {
     	        	wait.until(ExpectedConditions.elementToBeClickable(By.id(idDataDoPgto))).click();
     	        }
     	        String inputTextDataDoPgto = dataAtualSharepoint;
     	        WebElement dataDoPgto = driver.findElement(By.id(idDataDoPgto));
    			String comandoJsDataDoPgto = "arguments[0].setAttribute('value','"+inputTextDataDoPgto+"')";
    			
    			try {
    				js.executeScript(comandoJsDataDoPgto, dataDoPgto); 
    				Thread.sleep(500);
    			} catch (Exception e) {
    				System.out.println("Passei: " + "recebimentoPO");
    				dataDoPgto.sendKeys(inputTextDataDoPgto);
    				Thread.sleep(500);
    			}
    			Thread.sleep(5000);
    		 */
    		
    		// Bot�o de Save
    		String textoBotaoSave = "Save";
    		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//span [text()='"+textoBotaoSave+"']"))).click();
    		pedido.setSalvoNoSharepoint(true);
    		
    		// Bot�o de Close
    		//String textoBotaoClose = "Close";
    		//wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//span [text()='"+textoBotaoClose+"']"))).click();
    		
    		Thread.sleep(3000);
    		
		} catch (Exception e) {
			// Se der alguma exce��o por conta das regras de preenchimento, gravo as mensagens em um arquivo e continuo o processamento dos demais
			String mensagemErro = e.toString();
			if (mensagemErro.contains("Regra de preenchimento invalida")) {
				// Usarei a informa��o de erro de preenchimento para mostrar no relat�rio final :)
				if (mensagemErro != null && !mensagemErro.isEmpty()) {
					
					if (mensagemErro.indexOf("java.lang.Exception") != -1) {
						mensagemErro = mensagemErro.replaceAll("java.lang.Exception: ", "");
					} 
					
					pedido.setMensagemDeErroNoPedido(mensagemErro);
					listaPedidosComErrosNasRegraDePreenchimentoNoSharePoint = listaPedidosComErrosNasRegraDePreenchimentoNoSharePoint + mensagemErro;
				}
				
				System.out.println(mensagemErro);
			} else {
				contadorErrosPreencherCamposBiling++;
				if (contadorErrosPreencherCamposBiling <= 30) {
					System.out.println("Passei no erro interno do preencherCamposBiling: " + pedido.getNumero() + " " + pedido.isSalvoNoSharepoint());
					preencherCamposBiling(driver, wait, js, pedido);
				} else {
					throw new Exception("Ocorreu um erro no cadastro do pedido no sharepoint: " + e);
				}

			}
			
		}
    	
    }	
    
    public static void preencheCampoCap(WebDriver driver, WebDriverWait wait, JavascriptExecutor js, String idCap, String textoCap) throws Exception  {
    	
    	WebElement cap = driver.findElement(By.id(idCap));
    	String comandoJsCap = "arguments[0].setAttribute('value','"+textoCap+"')";
    	
    	try {
    		js.executeScript(comandoJsCap, cap); 
    		Thread.sleep(500);
    	} catch (Exception e) {
    		System.out.println("Passei: " + "CAP");
    		cap.sendKeys(textoCap);
    		Thread.sleep(500);
    	}
    	Thread.sleep(5000);

    }
    
    public static void preencheCampoObservacao(WebDriver driver, WebDriverWait wait, JavascriptExecutor js, String idObservacao, String textoObservacao) throws Exception  {
    	
    	WebElement observacao = driver.findElement(By.id(idObservacao));
    	String comandoJsObservacao = "arguments[0].setAttribute('value','"+textoObservacao+"')";
    	
    	try {
    		js.executeScript(comandoJsObservacao, observacao); 
    		Thread.sleep(500);
    	} catch (Exception e) {
    		observacao.sendKeys(textoObservacao);
    		Thread.sleep(500);
    	}
    	Thread.sleep(5000);

    }
   
    // Op��o de Restablecer Filtros
    public static void restabelecerFiltros(WebDriverWait wait) throws InterruptedException, IOException  {
       	// Se for utilizar o m�todo buscaAvancadaComUrl, o id da Op��o de Restablecer Filtros ser� TID_CONTENTPANEL_supplier_SupplierWelcomePanel_BUTTON_RESET_FILTERS
    	// Se for utilizar o m�todo buscaAvancada, o id da Op��o de Restablecer Filtros ser� TID_CONTENTPANEL_supplier_SupplierAdvancedSearchPanel_BUTTON_RESET_FILTERS
    	try {
    		Thread.sleep(1000);
    		wait.until(ExpectedConditions.elementToBeClickable(By.id("TID_CONTENTPANEL_supplier_SupplierAdvancedSearchPanel_BUTTON_RESET_FILTERS"))).click();
		} catch (Exception e) {
			System.out.println("Nao achou o id TID_CONTENTPANEL_supplier_SupplierAdvancedSearchPanel_BUTTON_RESET_FILTERS.");
			System.out.println("Tentando o TID_CONTENTPANEL_supplier_SupplierWelcomePanel_BUTTON_RESET_FILTERS.");
			wait.until(ExpectedConditions.elementToBeClickable(By.id("TID_CONTENTPANEL_supplier_SupplierWelcomePanel_BUTTON_RESET_FILTERS"))).click();
		}
    	
    	Thread.sleep(2000);
    }
    
    // Bot�o Aplicar Filtros
    public static void aplicarFiltros(WebDriver driver, WebDriverWait wait, JavascriptExecutor js, String contractNumber) throws Exception  {
       	// Se for utilizar o m�todo buscaAvancadaComUrl, o id do Bot�o Aplicar Filtros ser� TID_CONTENTPANEL_supplier_SupplierWelcomePanel_BUTTON_APPLY_FILTERS
    	// Se for utilizar o m�todo buscaAvancada, o id do Bot�o Aplicar Filtros ser� TID_CONTENTPANEL_supplier_SupplierAdvancedSearchPanel_BUTTON_APPLY_FILTERS
		try {
			
			Thread.sleep(1000);
			wait.until(ExpectedConditions.elementToBeClickable(By.id("TID_CONTENTPANEL_supplier_SupplierAdvancedSearchPanel_BUTTON_APPLY_FILTERS"))).click();
		
		} catch (Exception e2) {
			
			try {
				Thread.sleep(1000);
				wait.until(ExpectedConditions.elementToBeClickable(By.id("TID_CONTENTPANEL_supplier_SupplierWelcomePanel_BUTTON_APPLY_FILTERS"))).click();
			} catch (Exception e3) {
				throw new Exception("Deu erro no botao TID_CONTENTPANEL_supplier_SupplierWelcomePanel_BUTTON_APPLY_FILTERS: " + e3);
			}
			
		}
		Thread.sleep(2000);
    }
    
    // Op��o de Aceitar Cookies
    public static void aceitarCookies(WebDriverWait wait) throws InterruptedException, IOException  {
    	wait.until(ExpectedConditions.elementToBeClickable(By.id("TID_DIV_COOKIESDISCLAIMERBAR_BUTTON_ACCEPT"))).click();
    	Thread.sleep(3000);
    }
    
    // Poup-up Tela Inicial
    public static void popUpTelaInicial(WebDriverWait wait) throws InterruptedException, IOException  {
    	// Marca op��o para n�o voltar a ver
    	wait.until(ExpectedConditions.elementToBeClickable(By.id("TID_CONTENTPANEL_supplier_SupplierWelcomePanel_DIALOGWINDOW_carouselDialog_CHECKBOX_doNotShowAgain"))).click();
    	Thread.sleep(1000);
    	
    	// Fecha pop-up
    	wait.until(ExpectedConditions.elementToBeClickable(By.className("v-window-closebox"))).click();
    	Thread.sleep(1000);

    }

    
    // Op��o de Busca Avan�ada
    // Se for usar esse m�todo, alguns bot�es ter�o os ids alterados.Para saber quais s�o os ids, procure nos coment�rios por buscaAvancadaComUrl
    // Se for usar o m�todo buscaAvancada, a ordem de execu��o dos try catch abaixo dos coment�rios deve ser invertida
    public static void buscaAvancadaComUrl(WebDriver driver, WebDriverWait wait) throws InterruptedException, IOException  {
    	driver.get(Util.getValor("url.adquira.pesquisa.avancada"));
    	Thread.sleep(5000);
    }
    
    // Op��o de Busca Avan�ada
    // Se for usar esse m�todo, alguns bot�es ter�o os ids alterados.Para saber quais s�o os ids, procure nos coment�rios por buscaAvancada
    // Se for usar o m�todo buscaAvancadaComUrl, a ordem de execu��o dos try catch abaixo dos coment�rios deve ser invertida
    public static void buscaAvancada(WebDriverWait wait) throws InterruptedException, IOException  {
    	wait.until(ExpectedConditions.elementToBeClickable(By.id("TID_SEARCHBAR_BUTTON_ADVANCED_SEARCH"))).click();
    	Thread.sleep(3000);
    }

    // Op��o de Ver Notifica��es
    public static void verNotificacoes(WebDriverWait wait) throws InterruptedException, IOException  {
    	wait.until(ExpectedConditions.elementToBeClickable(By.id("TID_NOTIFICATIONBAR_BUTTON_CLOSE_ICON"))).click();
    	Thread.sleep(1000);
    }
    
    
    // Op��o de Ver Notifica��es
    public static void fazerLogoutAdquiraDepoisLoginDepoisBuscaAvancadaDepoisRestabeleceFiltros(WebDriver driver, WebDriverWait wait, JavascriptExecutor js, String usuario, String senha) throws Exception  {
		
		// Est� dando erro de logout no servidor
		// O bot�o de logout est� ficando escondido
		// ent�o retirarei o logout e o login por enquanto
		fazerLogoutAdquira(driver, wait);
		
    	fazerLoginAdquira(driver, wait, js, usuario, senha);
    	
    	//acessarPaginaInicial(driver, wait);
	    
		// Op��o de Busca Avan�ada
		buscaAvancadaComUrl(driver, wait);
    	
		// Op��o de Restabelecer Filtros
		restabelecerFiltros(wait);
    	
    }
    
    public static void encontrarContractNumberNoPdfDoPedido(WebDriver driver, WebDriverWait wait, JavascriptExecutor js, Pedido pedidoNaoFaturado, String subdiretorioPdfsBaixados2  )  throws IOException, Exception {
    	
    	
    	String numeroContractNumber = Pdf.retornaNumeroContractNumber(subdiretorioPdfsBaixados2 + "/" + pedidoNaoFaturado.getNumero() + ".pdf");
    	
    	pedidoNaoFaturado.setContractNumberConforme(false);
    	
    	if (numeroContractNumber != null && !numeroContractNumber.isEmpty()) {
    		numeroContractNumber = numeroContractNumber.trim();
    		
    		if (listaContractNumbers != null && !listaContractNumbers.isEmpty()) {
    			
    			for (ContractNumber contractNumber : listaContractNumbers) {
    				
    				if (numeroContractNumber.equals(contractNumber.getNumero().trim())) {
    					pedidoNaoFaturado.setContractNumber(contractNumber);
    					pedidoNaoFaturado.setContractNumberConforme(true);
    					break;
    				}
    				
    			}
    			
    		}
    	
    	} else {
    		numeroContractNumber = "-";
    	}

    	if (!pedidoNaoFaturado.isContractNumberConforme()) {
    		
    		ContractNumber contractNumber = new ContractNumber();
    		contractNumber.setContrato("-");
    		contractNumber.setFrente("-");
    		contractNumber.setNumero(numeroContractNumber);
    		contractNumber.setWbs("-");
    		pedidoNaoFaturado.setContractNumber(contractNumber);
    		
    	}
    	
    }
    
    public static void fazerDownlodPdfPedidoMoveArquivosEDescompacta(WebDriver driver, WebDriverWait wait, JavascriptExecutor js, Pedido pedidoNaoFaturado, String subdiretorioPdfsBaixados2, String usuario, String senha  ) throws IOException, Exception {
		
		try {
			
			// Deleto arquivos que existirem no diret�rio de relat�rios
			apagaArquivosDiretorioDeRelatorios(Util.getValor("caminho.download.relatorios"));
			
			// Baixa os pdfs de cada pedido
			fazerDownlodPdfPedido(driver, wait, js, pedidoNaoFaturado);
			Thread.sleep(1000);
			
			if (pedidoNaoFaturado.isEncontrouPdfAnexo()) {
				
				//Move o pdf baixado do diret�rio relatorios para o diret�rio correto
				moverArquivosEntreDiretorios2(Util.getValor("caminho.download.relatorios") + "\\" + nomeZipBaixado, subdiretorioPdfsBaixados);
				Thread.sleep(1000);
				
				// Descompacta o arquivo e renomeia o pdf com o numero do pedido
				descompactaArquivoZip(subdiretorioPdfsBaixados2 + "/" + nomeZipBaixado, subdiretorioPdfsBaixados2 + "/" + pedidoNaoFaturado.getNumero() + ".pdf");
				Thread.sleep(1000);
				
			}
			
			
		} catch (Exception e) {
			contadorErrosMoverArquivos ++;
			// Executo at� 50 vezes se der erro no aplicarFiltros
			if (contadorErrosMoverArquivos <= 50) {
				System.out.println("Deu erro no metodo fazerDownlodPdfPedidoMoveArquivosEDescompacta, tentativa de acerto: " + contadorErrosMoverArquivos);
				
				// Est� dando erro de logout no servidor
				// O bot�o de logout est� ficando escondido
				// ent�o retirarei o logout e o login por enquanto
				fazerLogoutAdquira(driver, wait);
				
				fazerLoginAdquira(driver, wait, js, usuario, senha);
				
				//acessarPaginaInicial(driver, wait);
				
				fazerDownlodPdfPedidoMoveArquivosEDescompacta(driver, wait, js, pedidoNaoFaturado, subdiretorioPdfsBaixados2, usuario, senha);

			}
			//Aqui n�o vou colocar o else para dar um throw new Exception porque percebi que existem pedidos repetidos.
			//Quando temos pedidos repetidos, o programa n�o est� deixando sobrescrever arquivos iguais (� o que eu achei vendo 
			//uma das execu��es, ou seja, n�o debuguei), ent�o cai no catch acima.
			//Por�m ao fim das 50 tentativas, ele sai e vai para o pr�ximo pedido, o que n�o aconteceria se tivesse o throw new Exception
		}
    	
    }
    
   public static void moverArquivosEntreDiretorios2(String caminhoArquivoOrigem, String caminhoDiretorioDestino) throws Exception{
    	
    	boolean sucesso = true;
    	File arquivoOrigem = new File(caminhoArquivoOrigem);
        File diretorioDestino = new File(caminhoDiretorioDestino);
        if (arquivoOrigem.exists() && diretorioDestino.exists()) {
        	sucesso = arquivoOrigem.renameTo(new File(diretorioDestino, arquivoOrigem.getName()));
        }
        
        if (!sucesso) {
        	throw new Exception("Ocorreu um erro no momento de mover o relatorio " + caminhoArquivoOrigem + " para " + caminhoDiretorioDestino);
        	
        }
        
    }


    public static void fazerDownlodRelatorioPorPeriodo(WebDriver driver, WebDriverWait wait, JavascriptExecutor js, String usuario, String senha) throws Exception  { 
    	
    	try {
    		
		// Deleto arquivos que existirem no diret�rio de relat�rios
		apagaArquivosDiretorioDeRelatorios(Util.getValor("caminho.download.relatorios"));
    	
    	extracaoPossuiPedidos = true;
    	
    	driver.get(Util.getValor("url.adquira.pesquisa.pedidos.por.periodo"));
    	
    	// Clicando no item PEDIDOS RECIBIDOS
    	//String idPedidosRecibidos = "//*[@id=\"TID_TOPMENU\"]/div[1]/div/span[2]";
    	//wait.until(ExpectedConditions.elementToBeClickable(By.id(idPedidosRecibidos))).click();
    	
    	// Clicando no item CONSULTA Y GESTI�N
    	//String idConsultaGestion = "//*[@id=\"marketplacecontrol-1122975566-overlays\"]/div[2]/div/div/span[1]";
    	//wait.until(ExpectedConditions.elementToBeClickable(By.id(idConsultaGestion))).click();
    	
    	// Esperando aparecer o link DESCARGAR LISTADO
    	String idDescargarListado = "TID_CONTENTPANEL_supplier_SupplierPurchaseOrdersListPanel_BUTTON_DOWNLOAD_LIST";
    	wait.until(ExpectedConditions.visibilityOfElementLocated(By.id(idDescargarListado)));
    	
    	// Esperando aparecer o texto Items por p�gina:
    	String textoItemsPorPagina = "Items por página:";
		// Aguarda o surgimento da palavra Items por p�gina:
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div [text()='"+textoItemsPorPagina+"']")));
		Thread.sleep(1000);
		
		// Clicar no item do menu DESDE
		//String idDesde = "//*[@id=\"TID_CONTENTPANEL_supplier_SupplierPurchaseOrdersListPanel\"]/div[3]/div[1]/div/div[1]/div/div[4]/div[1]";
		String idDesde = "//*[@id=\"TID_CONTENTPANEL_supplier_SupplierPurchaseOrdersListPanel\"]/div[3]/div[1]/div/div[1]/div/div[5]/div[1]";
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath(idDesde))).click();
		Thread.sleep(1000);
		
		// Clicar no item �LTIMA SEMANA
		String idUltimaSemana = "//*[@id=\"_DateRangeFilterPopupButton_Button_LAST_WEEK\"]";
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath(idUltimaSemana))).click();
		Thread.sleep(1000);

		try {
			// Se for utilizar o m�todo buscaAvancadaComUrl, o id da a��o Clicar no link Descargar Listado ser� TID_CONTENTPANEL_supplier_SupplierWelcomePanel_CONTENTPANEL_supplier_SupplierPurchaseOrdersListPanel_BUTTON_DOWNLOAD_LIST
			// Se for utilizar o m�todo buscaAvancada, o id da a��o Clicar no link Descargar Listado ser� TID_CONTENTPANEL_supplier_SupplierAdvancedSearchPanel_CONTENTPANEL_supplier_SupplierPurchaseOrdersListPanel_BUTTON_DOWNLOAD_LIST
			// Para este caso, creio que a regra acima n�o se aplica, ou seja, o id sempre ser� TID_CONTENTPANEL_supplier_SupplierAdvancedSearchPanel_CONTENTPANEL_supplier_SupplierPurchaseOrdersListPanel_BUTTON_DOWNLOAD_LIST
			// independente se o m�todo for o buscaAvancadaComUrl ou o buscaAvancada
			
			//driver.findElement(By.id(idDescargarListado));
			
			// Verifico se existe o link Descargar Listado
			// Caso o link Descargar Listado n�o apare�a, � porque provavelmente n�o existem pedidos para o per�odo selecionado
			// Neste caso, armazeno os contract numbers em um array
			WebDriverWait waitIdDescargarListado = new WebDriverWait(driver, 10);
			waitIdDescargarListado.until(ExpectedConditions.visibilityOfElementLocated(By.id(idDescargarListado)));
			Thread.sleep(3000);
			
		} catch (Exception e) {
			System.out.println("Nao foram encontrados pedidos");
			extracaoPossuiPedidos = false;
		}
		
		// Filtro do menu por per�odo
		
		
		if (extracaoPossuiPedidos) {	
			
			// Aguarda o surgimento da palavra Items por p�gina:
			wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div [text()='"+textoItemsPorPagina+"']")));
			
			// Clicar no link Descargar Listado
			wait.until(ExpectedConditions.elementToBeClickable(By.id(idDescargarListado))).click();
			
			// Op��o de Incluir L�neas
		    String idTextoIncluirLineas= "//span[@id='pid-includeLines']/label";
		    wait.until(ExpectedConditions.elementToBeClickable(By.xpath(idTextoIncluirLineas))).click();
		    Thread.sleep(3000);
		    
			// Clicar no bot�o Generar Fichero
			// Se for utilizar o m�todo buscaAvancadaComUrl, o bot�o Generar Ficheiro ter� o id TID_CONTENTPANEL_supplier_SupplierWelcomePanel_BUTTON_GENERATE_FILE
			// Se for utilizar o m�todo buscaAvancada, o bot�o Generar Ficheiro ter� o id TID_CONTENTPANEL_supplier_SupplierAdvancedSearchPanel_BUTTON_GENERATE_FILE
		    WebDriverWait waitGenerarFichero = new WebDriverWait(driver, 5);
			try {                                                                                                                                                                                          
				waitGenerarFichero.until(ExpectedConditions.elementToBeClickable(By.id("TID_CONTENTPANEL_supplier_SupplierPurchaseOrdersListPanel_BUTTON_GENERATE_FILE"))).click();
			} catch (Exception e2) {
				waitGenerarFichero.until(ExpectedConditions.elementToBeClickable(By.id("TID_CONTENTPANEL_supplier_SupplierAdvancedSearchPanel_BUTTON_GENERATE_FILE"))).click();
			}
			Thread.sleep(2000);
			
			// Clicar no bot�o Descargar
			// Se for utilizar o m�todo buscaAvancadaComUrl, o bot�o Descargar ter� o id TID_CONTENTPANEL_supplier_SupplierWelcomePanel_DIALOGWINDOW_DIALOG_FILE_DOWNLOAD_BUTTON_DOWNLOAD
			// Se for utilizar o m�todo buscaAvancada, o bot�o Descargar ter� o id TID_CONTENTPANEL_supplier_SupplierAdvancedSearchPanel_DIALOGWINDOW_DIALOG_FILE_DOWNLOAD_BUTTON_DOWNLOAD
			 WebDriverWait waitbotaoDescargar = new WebDriverWait(driver, 5);
			try {
				waitbotaoDescargar.until(ExpectedConditions.elementToBeClickable(By.id("TID_CONTENTPANEL_supplier_SupplierPurchaseOrdersListPanel_DIALOGWINDOW_DIALOG_FILE_DOWNLOAD_BUTTON_DOWNLOAD"))).click();
			} catch (Exception e2) {
				waitbotaoDescargar.until(ExpectedConditions.elementToBeClickable(By.id("TID_CONTENTPANEL_supplier_SupplierAdvancedSearchPanel_DIALOGWINDOW_DIALOG_FILE_DOWNLOAD_BUTTON_DOWNLOAD"))).click();
			}
			Thread.sleep(2000);
			
			// Pego o nome do arquivo
			// Se for utilizar o m�todo buscaAvancadaComUrl, o id do nome do arquivo ter� no meio do nome a palavra TID_CONTENTPANEL_supplier_SupplierWelcomePanel_DIALOGWINDOW_DIALOG_FILE_DOWNLOAD
			// Se for utilizar o m�todo buscaAvancada, o id do nome do arquivo ter� no meio do nome a palavra TID_CONTENTPANEL_supplier_SupplierAdvancedSearchPanel_DIALOGWINDOW_DIALOG_FILE_DOWNLOAD
			String idNomeArquivo = "";
			WebDriverWait waitNomeArquivo = new WebDriverWait(driver, 5);
			try {
				idNomeArquivo = "//*[@id=\"TID_CONTENTPANEL_supplier_SupplierPurchaseOrdersListPanel_DIALOGWINDOW_DIALOG_FILE_DOWNLOAD\"]/div/div/div[3]/div/div/div/div/div[2]/div[1]/span[2]";
				waitNomeArquivo.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(idNomeArquivo)));
			} catch (Exception e2) {
				idNomeArquivo = "//div[@id='TID_CONTENTPANEL_supplier_SupplierAdvancedSearchPanel_DIALOGWINDOW_DIALOG_FILE_DOWNLOAD']/div/div/div[3]/div/div/div/div/div[2]/div/span[2]";
				waitNomeArquivo.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(idNomeArquivo)));
			}
			
			// Pego o nome do arquivo
			WebElement textoNomeArquivo = driver.findElement(By.xpath(idNomeArquivo));
			
			String caracterNomeArquivo = " ";
			if ("Chrome".equals(Util.getValor("navegador"))) {
				caracterNomeArquivo = "_";
			} else if ("Firefox".equals(Util.getValor("navegador"))) {
				caracterNomeArquivo = " ";
			}
			
			nomeRelatorioBaixado = textoNomeArquivo.getText().replaceAll(":", caracterNomeArquivo);
			
			// Se for utilizar o m�todo buscaAvancadaComUrl, o bot�o Cancelar ter� o id TID_CONTENTPANEL_supplier_SupplierWelcomePanel_DIALOGWINDOW_DIALOG_FILE_DOWNLOAD_BUTTON_CANCEL
			// Se for utilizar o m�todo buscaAvancada, o bot�o Cancelar ter� o id TID_CONTENTPANEL_supplier_SupplierAdvancedSearchPanel_DIALOGWINDOW_DIALOG_FILE_DOWNLOAD_BUTTON_CANCEL
			// Clicar no bot�o Cancelar
			WebDriverWait waitbotaoCancelar = new WebDriverWait(driver, 5);
			try {
				waitbotaoCancelar.until(ExpectedConditions.elementToBeClickable(By.id("TID_CONTENTPANEL_supplier_SupplierPurchaseOrdersListPanel_DIALOGWINDOW_DIALOG_FILE_DOWNLOAD_BUTTON_CANCEL"))).click();
			} catch (Exception e2) {
				waitbotaoCancelar.until(ExpectedConditions.elementToBeClickable(By.id("TID_CONTENTPANEL_supplier_SupplierAdvancedSearchPanel_DIALOGWINDOW_DIALOG_FILE_DOWNLOAD_BUTTON_CANCEL"))).click();
			}
			Thread.sleep(2000);
			
		}
		
		//voltarParaHome(driver, wait, js);
		
		// Est� dando erro de logout no servidor
		// O bot�o de logout est� ficando escondido
		// ent�o retirarei o logout e o login por enquanto
		//fazerLogoutAdquira(driver, wait);
		
		//acessarPaginaInicial(driver, wait);
		
		} catch (Exception e2) {
			contadorfazerDownlodRelatorioPorPeriodo ++;
			// Executo ate 100 vezes se der erro no fazerDownlodRelatorioPorPeriodo
			if (contadorfazerDownlodRelatorioPorPeriodo <= 100) {
				
				System.out.println("Deu erro no metodo fazerDownlodRelatorioPorPeriodo, tentativa de acerto: " + contadorfazerDownlodRelatorioPorPeriodo);
				// Est� dando erro de logout no servidor
				// O bot�o de logout est� ficando escondido
				// ent�o retirarei o logout e o login por enquanto
				fazerLogoutAdquira(driver, wait);
				fazerLoginAdquira(driver, wait, js, usuario, senha);
				//acessarPaginaInicial(driver, wait);
				fazerDownlodRelatorioPorPeriodo(driver, wait, js, usuario, senha);
			
			} else {
	        	throw new Exception("Ocorreu um erro no momento de fazer o download do relatorio por periodo: " + e2);
	        }

		
		}
    	
    }
    
    public static void voltarParaHome (WebDriver driver, WebDriverWait wait, JavascriptExecutor js) throws Exception {
    	
		js.executeScript("javascript:history.back()");
		
		// Aguarda aparecer texto na home
		// Se esse texto n�o aparecer na home, � porque o Adquira est� muito lento e n�o conseguiu carregar as informa��es
		// e neste caso ele provavelmente n�o conseguir� seguir adiante
	    String textoFaturasEmitidasPorMes = "FACTURAS EMITIDAS POR MES";
	    wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div [text()='"+textoFaturasEmitidasPorMes+"']")));
	    
    	// Aguarda aparecer texto na home
    	//String textoPosicionGlobal = "Posición global";
    	//wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//span[contains(.,'Posición global')]")));
	    
	    fecharMensagemVerNotificacoes(driver);

    }
    
    public static void aplicarFiltro (WebDriver driver, WebDriverWait wait, JavascriptExecutor js, String contractNumber, String usuario, String senha ) throws Exception {
    	
    	preencherCamposFiltro(driver, wait, contractNumber);
    	
    	try {
			
    		// Bot�o Aplicar
        	// Se for utilizar o m�todo buscaAvancadaComUrl, o bot�o Aplicar ter� o id TID_CONTENTPANEL_supplier_SupplierWelcomePanel_BUTTON_APPLY
        	// Se for utilizar o m�todo buscaAvancada, o bot�o Aplicar ter� o id TID_CONTENTPANEL_supplier_SupplierAdvancedSearchPanel_BUTTON_APPLY
    		WebDriverWait waitbotaoAplicar = new WebDriverWait(driver, 5);
    		try {
    			waitbotaoAplicar.until(ExpectedConditions.elementToBeClickable(By.id("TID_CONTENTPANEL_supplier_SupplierWelcomePanel_BUTTON_APPLY"))).click();
    		} catch (Exception e2) {
    			waitbotaoAplicar.until(ExpectedConditions.elementToBeClickable(By.id("TID_CONTENTPANEL_supplier_SupplierAdvancedSearchPanel_BUTTON_APPLY"))).click();
    		}
    		Thread.sleep(5000);

    	} catch (Exception e) {
    		
    		contadorErros++;
    		// Executo at� 20 vezes se der erro no preenchemento dos campos do filtro
    		if (contadorErros <= 20) {
    			
    			System.out.println("Deu erro no preenchimento dos filtros");
    			// Caso ocorra algum problema na automa��o onde o campo do Nome n�o seja preenchido com a palavra contract number,
    			// clico no bot�o Aceptar da mensagem de erro que ocorre
            	// Se for utilizar o m�todo buscaAvancadaComUrl, o bot�o Aplicar ter� o id TID_CONTENTPANEL_supplier_SupplierWelcomePanel_BUTTON_ACCEPT
            	// Se for utilizar o m�todo buscaAvancada, o bot�o Aplicar ter� o id TID_CONTENTPANEL_supplier_SupplierAdvancedSearchPanel_BUTTON_ACCEPT
    			WebDriverWait waitbotaoAceptar = new WebDriverWait(driver, 5);
    			try {
    				waitbotaoAceptar.until(ExpectedConditions.elementToBeClickable(By.id("TID_CONTENTPANEL_supplier_SupplierWelcomePanel_BUTTON_ACCEPT"))).click();
				} catch (Exception e2) {
					waitbotaoAceptar.until(ExpectedConditions.elementToBeClickable(By.id("TID_CONTENTPANEL_supplier_SupplierAdvancedSearchPanel_BUTTON_ACCEPT"))).click();
				}
    			Thread.sleep(3000);
    			
    			// Cancelo o modal de filtro
            	// Se for utilizar o m�todo buscaAvancadaComUrl, o bot�o para cancelar o modal ter� o id TID_CONTENTPANEL_supplier_SupplierWelcomePanel_BUTTON_CANCEL
            	// Se for utilizar o m�todo buscaAvancada, o bot�o para cancelar o modal ter� o id TID_CONTENTPANEL_supplier_SupplierAdvancedSearchPanel_BUTTON_CANCEL
    			WebDriverWait waitbotaoCancel = new WebDriverWait(driver, 5);
    			try {
    				waitbotaoCancel.until(ExpectedConditions.elementToBeClickable(By.id("TID_CONTENTPANEL_supplier_SupplierWelcomePanel_BUTTON_CANCEL"))).click();
				} catch (Exception e2) {
					waitbotaoCancel.until(ExpectedConditions.elementToBeClickable(By.id("TID_CONTENTPANEL_supplier_SupplierAdvancedSearchPanel_BUTTON_CANCEL"))).click();
				}
    			
    			Thread.sleep(3000);
    			
    			// Fa�o o logout, depois o login depois entro na busca avan�ada e por fim restabele�o filtros
    			// Dessa forma, descobri que o campo Nome volta a funcionar preenchendo a palavra contract number
    			fazerLogoutAdquiraDepoisLoginDepoisBuscaAvancadaDepoisRestabeleceFiltros(driver, wait, js, usuario, senha);
    			
    			// Op��o de Datos Adicionales
            	// Se for utilizar o m�todo buscaAvancadaComUrl, a Op��o de Datos Adicionales ter� o id TID_CONTENTPANEL_supplier_SupplierWelcomePanel_CUSTOMFIELD_EXTRINSICDATA_ICONBUTTON_ICON_EDIT
            	// Se for utilizar o m�todo buscaAvancada, a Op��o de Datos Adicionales ter� o id TID_CONTENTPANEL_supplier_SupplierAdvancedSearchPanel_CUSTOMFIELD_EXTRINSICDATA_ICONBUTTON_ICON_EDIT
    			WebDriverWait waitbotaoDatosAdicionales = new WebDriverWait(driver, 5);
    			try {
    				waitbotaoDatosAdicionales.until(ExpectedConditions.elementToBeClickable(By.id("TID_CONTENTPANEL_supplier_SupplierWelcomePanel_CUSTOMFIELD_EXTRINSICDATA_ICONBUTTON_ICON_EDIT"))).click();
				} catch (Exception e2) {
					waitbotaoDatosAdicionales.until(ExpectedConditions.elementToBeClickable(By.id("TID_CONTENTPANEL_supplier_SupplierAdvancedSearchPanel_CUSTOMFIELD_EXTRINSICDATA_ICONBUTTON_ICON_EDIT"))).click();
				}
    			Thread.sleep(3000);
    			aplicarFiltro(driver, wait, js, contractNumber, usuario, senha);
    		
    		}
    		
		}
    	
    }
    
    // Preenche o campo Nombre digitando a palavra e apertando o ENTER
    public static void preencherCamposFiltro (WebDriver driver, WebDriverWait wait, String contractNumber) throws InterruptedException {
    	
    	// Campo Nombre
    	String idClasseDoCampoNombre = "v-filterselect-input";
    	wait.until(ExpectedConditions.visibilityOfElementLocated(By.className(idClasseDoCampoNombre)));
    	WebElement campoNombre = driver.findElement(By.className(idClasseDoCampoNombre));
    	campoNombre.clear();
/*    	campoNombre.sendKeys("CONTRACT");
    	Thread.sleep(2000);
    	campoNombre.sendKeys(" ");
    	Thread.sleep(2000);
    	campoNombre.sendKeys("NUMBER");
    	Thread.sleep(2000);*/
    	
    	campoNombre.sendKeys("C");
    	Thread.sleep(1000);
    	campoNombre.sendKeys("O");
    	Thread.sleep(1000);
    	campoNombre.sendKeys("N");
    	Thread.sleep(1000);
    	campoNombre.sendKeys("T");
    	Thread.sleep(1000);
    	campoNombre.sendKeys("R");
    	Thread.sleep(1000);
    	campoNombre.sendKeys("A");
    	Thread.sleep(1000);
    	campoNombre.sendKeys("C");
    	Thread.sleep(1000);
    	campoNombre.sendKeys("T");
    	Thread.sleep(1000);
    	campoNombre.sendKeys(" ");
    	Thread.sleep(2000);
    	campoNombre.sendKeys("N");
    	Thread.sleep(1000);
       	campoNombre.sendKeys("U");
    	Thread.sleep(1000);
       	campoNombre.sendKeys("M");
    	Thread.sleep(1000);
       	campoNombre.sendKeys("B");
    	Thread.sleep(1000);
       	campoNombre.sendKeys("E");
    	Thread.sleep(1000);
       	campoNombre.sendKeys("R");
    	Thread.sleep(1000);
    	
    	// Aperta o ENTER para selecionar a op��o
    	campoNombre.sendKeys(Keys.RETURN); 
    	Thread.sleep(3000);
    	
    	// Campo Valor
    	String idCampoValor = "pid-value";
    	wait.until(ExpectedConditions.visibilityOfElementLocated(By.id(idCampoValor)));
    	WebElement campoValor = driver.findElement(By.id(idCampoValor));
    	campoValor.clear();
    	campoValor.sendKeys(contractNumber);
    	Thread.sleep(3000);
    	
		// Bot�o A�adir
    	// Se for utilizar o m�todo buscaAvancadaComUrl, o bot�o A�adir ter� o id TID_CONTENTPANEL_supplier_SupplierWelcomePanel_BUTTON_ADD
    	// Se for utilizar o m�todo buscaAvancada, o bot�o A�adir ter� o id TID_CONTENTPANEL_supplier_SupplierAdvancedSearchPanel_BUTTON_ADD
    	WebDriverWait waitbotaoAnadir = new WebDriverWait(driver, 5);
		try {
			waitbotaoAnadir.until(ExpectedConditions.elementToBeClickable(By.id("TID_CONTENTPANEL_supplier_SupplierWelcomePanel_BUTTON_ADD"))).click();
		} catch (Exception e2) {
			waitbotaoAnadir.until(ExpectedConditions.elementToBeClickable(By.id("TID_CONTENTPANEL_supplier_SupplierAdvancedSearchPanel_BUTTON_ADD"))).click();
		}
		Thread.sleep(2000);
	
    	
    }
    
    // Preenche o campo Nombre selecionando a palavra desejada em um combo de op��es
    public static void preencherCamposFiltro2 (WebDriver driver, WebDriverWait wait, String contractNumber) throws InterruptedException {
    	
    	// Abre as op��es do combo de pesquisa
    	String idCampoNombre = "//div[@id='TID_CONTENTPANEL_supplier_SupplierAdvancedSearchPanel_COMBOBOX_NAME']/input";
    	wait.until(ExpectedConditions.elementToBeClickable(By.xpath(idCampoNombre))).click();
    	WebElement campoNombre = driver.findElement(By.xpath(idCampoNombre));
    	campoNombre.clear();
    	campoNombre.sendKeys(Keys.DOWN);
    	
    	// Clica no sinal de mais at� encontrar a op��o de CONTRACT NUMBER
    	String idSinalMais = "//div[@id='VAADIN_COMBOBOX_OPTIONLIST']/div/div[3]";
    	
    	for (int i = 0; i < 3; i++) {
    		wait.until(ExpectedConditions.elementToBeClickable(By.xpath(idSinalMais))).click();
    		Thread.sleep(2000);
    	}
    	
    	// Seleciona a op��o de Contract Number
    	wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//div[@id='VAADIN_COMBOBOX_OPTIONLIST']/div/div[2]/table/tbody/tr/td/span"))).click();
    	Thread.sleep(2000);
    	
    	// Campo Valor
    	String idCampoValor = "TID_CONTENTPANEL_supplier_SupplierAdvancedSearchPanel_TEXTFIELD_value";
    	wait.until(ExpectedConditions.visibilityOfElementLocated(By.id(idCampoValor)));
    	WebElement campoValor = driver.findElement(By.id(idCampoValor));
    	campoValor.clear();
    	campoValor.sendKeys(contractNumber);
    	Thread.sleep(1000);
    	
		// Bot�o A�adir
		wait.until(ExpectedConditions.elementToBeClickable(By.id("TID_CONTENTPANEL_supplier_SupplierAdvancedSearchPanel_BUTTON_ADD"))).click();
		Thread.sleep(2000);
    	
    }
    
    
    public static void fazerDownlodPdfPedido(WebDriver driver, WebDriverWait wait, JavascriptExecutor js, Pedido pedido  ) throws InterruptedException {
    	
       	// Aguardando o campo de B�squeda Avanzada aparecer 
        wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("TID_SEARCHBAR_BUTTON_ADVANCED_SEARCH")));
        
        // Campo de Busca
    	String idCampoBusca = "TID_SEARCHBAR_TEXTFIELD";
    	wait.until(ExpectedConditions.visibilityOfElementLocated(By.id(idCampoBusca)));
    	WebElement campoBusca = driver.findElement(By.id(idCampoBusca));
    	campoBusca.clear();
    	campoBusca.sendKeys(pedido.getNumero());
    	Thread.sleep(1000);
    	
    	// Aperta o ENTER para selecionar a op��o
    	campoBusca.sendKeys(Keys.RETURN); 
    	Thread.sleep(2000);
    	
		// Clica no pedido retornado pela busca
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//div[@id='TID_CONTENTPANEL_supplier_SupplierSearchResultsPanel']/div[3]/div/div[2]/div/div/div[2]/span[2]"))).click();
		Thread.sleep(2000);
		
		// Se o pedido n�o tiver pdf para baixar, seto o atributo encontrouPdfAnexo como false
		try {
			
			// Clica no link descargar pdf
			wait.until(ExpectedConditions.elementToBeClickable(By.id("TID_CONTENTPANEL_supplier_SupplierPurchaseOrdersListPanel_LIST_DETAIL_PANEL_SupplierPurchaseOrderDetailPanel_BUTTON_DOWNLOAD_PDF"))).click();
			Thread.sleep(2000);
			
			// Pego o nome do arquivo
			String idNomeArquivo = "//div[@id='TID_CONTENTPANEL_supplier_SupplierPurchaseOrdersListPanel_DIALOGWINDOW_DIALOG_FILE_DOWNLOAD']/div/div/div[3]/div/div/div/div/div[2]/div/span[2]";
			wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(idNomeArquivo)));
			WebElement textoNomeArquivo = driver.findElement(By.xpath(idNomeArquivo));
			nomeZipBaixado = textoNomeArquivo.getText().replaceAll("pdf", "zip");
			
			// Clica no bot�o descargar
			wait.until(ExpectedConditions.elementToBeClickable(By.id("TID_CONTENTPANEL_supplier_SupplierPurchaseOrdersListPanel_DIALOGWINDOW_DIALOG_FILE_DOWNLOAD_BUTTON_DOWNLOAD"))).click();
			Thread.sleep(2000);
			
			// Clicar no bot�o Cancelar
			wait.until(ExpectedConditions.elementToBeClickable(By.id("TID_CONTENTPANEL_supplier_SupplierPurchaseOrdersListPanel_DIALOGWINDOW_DIALOG_FILE_DOWNLOAD_BUTTON_CANCEL"))).click();
			Thread.sleep(2000);
			
			pedido.setEncontrouPdfAnexo(true);
			
		} catch (Exception e) {
			pedido.setEncontrouPdfAnexo(false);
		}
		
		
    }
    
    public static void recuperaContractNumbersSharepoint(WebDriver driver, WebDriverWait wait) throws Exception {
    	
    	try {
    		
    		// Abrindo a URl
    		// Lista do sharepoint que tem os contract numbers ativos
    		driver.manage().window().maximize();
    		driver.get(Util.getValor("url.contract.numbers.sharepoint"));
    		// Espera aparecer o texto new item
    		String textoNewItem = "new item";
    		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//span [text()='"+textoNewItem+"']")));
    		Thread.sleep(3000);
    		// Click no link de Status para abrir as linhas de contract numbers
    		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//body[1]/form[1]/div[12]/div[1]/div[2]/div[2]/div[3]/div[1]/div[1]/div[1]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/table[2]/tbody[2]/tr[1]/td[1]/a[1]"))).click();
    		Thread.sleep(3000);
    		// Quantidade de linhas de contract numbers
    		String qtdLinhasContractNumbers =  wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("/html/body/form/div[12]/div/div[2]/div[2]/div[3]/div[1]/div/div/div/table/tbody/tr/td/table[2]/tbody[2]/tr/td/span"))).getText().trim();
    		// Retirando os caracteres que n�o forem n�meros
    		String somenteNumerosQtdLinhasContractNumbers = qtdLinhasContractNumbers.replaceAll("[^0-9]", "");
    		int quantidadeLinhasContractNumbers = Integer.parseInt(somenteNumerosQtdLinhasContractNumbers);
    		
    		if (quantidadeLinhasContractNumbers > 0) {
    			
    			for (int i = 1; i <= quantidadeLinhasContractNumbers; i ++) {
    				
    				String idProjeto     = "/html/body/form/div[12]/div/div[2]/div[2]/div[3]/div[1]/div/div/div/table/tbody/tr/td/table[2]/tbody[3]/tr[" + i + "]/td[2]";
    				String idFrente      = "/html/body/form/div[12]/div/div[2]/div[2]/div[3]/div[1]/div/div/div/table/tbody/tr/td/table[2]/tbody[3]/tr[" + i + "]/td[3]";
    				String idContratoSap = "/html/body/form/div[12]/div/div[2]/div[2]/div[3]/div[1]/div/div/div/table/tbody/tr/td/table[2]/tbody[3]/tr[" + i + "]/td[4]"; 
    				
    				String projeto = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(idProjeto))).getText();
    				String frente = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(idFrente))).getText();
    				String contratoSap = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(idContratoSap))).getText();
    				
    				ContractNumber contractNumber = new ContractNumber();
    				
    				// Contrato
    				if (projeto != null && !projeto.isEmpty()) {
    					contractNumber.setContrato(projeto.trim());
    				}
    				// Frente
    				if (frente != null && !frente.isEmpty()) {
    					contractNumber.setFrente(frente.trim());
    				}
    				// Contrato Sap
    				if (contratoSap != null && !contratoSap.isEmpty()) {
    					contractNumber.setNumero(contratoSap.trim());
    				}
    				// WBS
    				contractNumber.setWbs("-");
    				
    				//System.out.println(contractNumber.getContrato() + " " +  contractNumber.getFrente() + " " + contractNumber.getNumero());
    				
    				listaNumerosContractNumbersDistintos.add(contractNumber.getNumero());
    				listaContractNumbersTemporaria.add(contractNumber);
    				
    			}
    			
    		}
    		 
		} catch (Exception e) {
			contadorErrosRecuperaContractNumbersSharepoint ++;
	            Thread.sleep(3000);
	            
	            // Tento fazer por at� 10 vezes
	            if (contadorErrosRecuperaContractNumbersSharepoint <= 10) {
	            	
					System.out.println("Deu erro no metodo recuperaContractNumbersSharepoint, tentativa de acerto: " + contadorErrosRecuperaContractNumbersSharepoint);
					recuperaContractNumbersSharepoint(driver, wait);
	            
	            } else {
	         	   throw new Exception("Erro no metodo recuperaContractNumbersSharepoint: " + e);
	            }
			}
    	
    }
    
    public static void recuperaPedidosSharepoint(WebDriver driver, WebDriverWait wait) throws Exception {
    	
    	try {
    		
    		// Abrindo a URl
    		// Lista do sharepoint que tem os pedidos
    		driver.manage().window().maximize();
    		driver.get(Util.getValor("url.pedidos.sharepoint"));
    		// Espera aparecer o texto new item
    		String textoNewItem = "new item";
    		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//span [text()='"+textoNewItem+"']")));
    		Thread.sleep(3000);
    		
    		// Click no link Automacao para abrir as linhas de pedidos
    		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[contains(text(),'Agrupamento do Robô')]"))).click();
    		Thread.sleep(3000);
    		// Quantidade de linhas de pedidos
    		String qtdLinhasPedidos =  wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("/html/body/form/div[12]/div/div[2]/div[2]/div[3]/div[1]/div/div/div/table/tbody/tr/td/table[2]/tbody[2]/tr/td/span"))).getText().trim();
    		// Retirando os caracteres que n�o forem n�meros
    		String somenteNumerosQtdLinhasPedidos = qtdLinhasPedidos.replaceAll("[^0-9]", "");
    		int quantidadeLinhasPedidos = Integer.parseInt(somenteNumerosQtdLinhasPedidos);
    		
    		if (quantidadeLinhasPedidos > 0) {
    			
    			for (int i = 1; i <= quantidadeLinhasPedidos; i ++) {
    				String idNumeroPedido = "/html/body/form/div[12]/div/div[2]/div[2]/div[3]/div[1]/div/div/div/table/tbody/tr/td/table[2]/tbody[3]/tr[" + i + "]/td[2]/div/a";
    				
    				String numeroPedido = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(idNumeroPedido))).getText();
    				
    				Pedido pedido = new Pedido();
    				
    				// Pedido
    				if (numeroPedido != null && !numeroPedido.isEmpty()) {
    					pedido.setNumero(numeroPedido.trim());
    				}
    				
    				listaPedidosFaturados.add(pedido);
    				
    			}
    			
    		}
    		
    	} catch (Exception e) {
			contadorErrosRecuperaPedidosSharepoint ++;
	            Thread.sleep(3000);
	            
	            // Tento fazer por at� 10 vezes
	            if (contadorErrosRecuperaPedidosSharepoint <= 10) {
	            	
					System.out.println("Deu erro no metodo recuperaPedidosSharepoint, tentativa de acerto: " + contadorErrosRecuperaPedidosSharepoint);
					recuperaPedidosSharepoint(driver, wait);
	            
	            } else {
	         	   throw new Exception("Erro no metodo recuperaPedidosSharepoint: " + e);
	            }
			}
		
    }
    
    public static void fazerLoginAdquira(WebDriver driver, WebDriverWait wait, JavascriptExecutor js, String usuario, String senha) throws Exception {
    	
    	try {
    		
   			// Abrindo a URl do Adquira
   			driver.manage().window().maximize();
   			driver.get(Util.getValor("url.adquira"));
   			Thread.sleep(2000);
            
       		// Preenchendo os dados de login
    		String idLogin = "//input[@id='campo-usuario']";
    		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(idLogin)));
    		WebElement username = driver.findElement(By.xpath(idLogin));
    		username.sendKeys(usuario);
    		Thread.sleep(1000);
    		
            // Preenchendo os dados da senha
    		String idSenha = "//input[@id='campo-contrasena']";
    		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(idSenha)));
            WebElement password = driver.findElement(By.xpath(idSenha));
            password.sendKeys(senha);
            Thread.sleep(1000);

            // Fazendo login
            WebElement botaoLogin = driver.findElement(By.xpath("//input[@id='boton-enviar']"));
            js.executeScript("arguments[0].click()", botaoLogin);
            
            Thread.sleep(10000);
            
    		String textoFaturasEmitidasPorMes = "FACTURAS EMITIDAS POR MES";
    		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div [text()='"+textoFaturasEmitidasPorMes+"']")));
            
        	fecharMensagemAceitarCookies(driver);
        	
        	fecharMensagemVerNotificacoes(driver);
        	
        	fecharPopUpTelaInicial(driver);

		} catch (Exception e) {
			contadorErrosLogin ++;
	            Thread.sleep(3000);
	            
	            // Tento fazer o login por at� 10 vezes
	            if (contadorErrosLogin <= 10) {
	            	
					System.out.println("Deu erro no metodo fazerLogin, tentativa de acerto: " + contadorErrosLogin);
					fazerLogoutAdquira(driver, wait);
					fazerLoginAdquira(driver, wait, js, usuario, senha);
	            
	            } else {
	         	   throw new Exception("Erro no Login: " + e);
	            }
			}
    	
    }
    
    public static void acessarPaginaInicial(WebDriver driver, WebDriverWait wait) throws Exception {
    	
    	//fechandoAbaEabrindoNova(driver);
    	
		// Abrindo a URl do Adquira
		driver.manage().window().maximize();
		driver.get(Util.getValor("url.adquira"));
		Thread.sleep(2000);
		
		fecharMensagemAceitarCookies(driver);
		
		fecharMensagemVerNotificacoes(driver);

		String textoFaturasEmitidasPorMes = "FACTURAS EMITIDAS POR MES";
	    try {
	    	wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div [text()='"+textoFaturasEmitidasPorMes+"']")));
		} catch (Exception e) {
			throw new Exception(e);
		}
	    
	    
	    
	    //wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div [text()='"+textoFaturasEmitidasPorMes+"']")));
	    // Se esse texto n�o aparecer na home, � porque o Adquira est� muito lento e n�o conseguiu carregar as informa��es
	    // e neste caso ele provavelmente n�o conseguir� seguir adiante
	    // Aguarda aparecer texto na home
	    /*
	    try {
	    	wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div [text()='"+textoFaturasEmitidasPorMes+"']")));
		} catch (Exception e) {
			wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//span[contains(.,'Posici�n globala')]")));
		}
    	*/
    	
    	
    	//fechandoAbaEabrindoNova(driver);
    	
    }
    
    public static void fecharMensagemAceitarCookies(WebDriver driver) throws Exception {
    	
	    // Op��o de Aceitar Cookies
	    // Se aparecer a op��o de aceitar cookies, clico no aceitar
	    // Aguardo at� 10 segundos para a op��o aparecer
	    // Se ela n�o aparecer dar� erro, da� sigo adiante
    	try {
    		
    		WebDriverWait waitAceitarCookies = new WebDriverWait(driver, 10);
    		aceitarCookies(waitAceitarCookies);
		
    	} catch (Exception e) {
    		System.out.println("Deu erro na opcao de aceitar cookies");
		}
    	
    }
    
    public static void fecharMensagemVerNotificacoes(WebDriver driver) throws Exception {
    	
	    // Op��o de Ver Notifica��es
	    // Se aparecer a op��o de ver notifica��es, clico no bot�o de fechar
	    // Aguardo at� 10 segundos para a op��o aparecer
	    // Se ela n�o aparecer dar� erro, da� sigo adiante
    	try {
    		
    		WebDriverWait waitVerNotificacoes = new WebDriverWait(driver, 10);
    		verNotificacoes(waitVerNotificacoes);
		
    	} catch (Exception e) {
    		System.out.println("Deu erro na opcao de ver notificacoes");
		}
    	
    }
    
    public static void fecharPopUpTelaInicial(WebDriver driver) throws Exception {
    	
	    // Se aparecer um pop-up com um informativo
	    // Aguardo at� 10 segundos para a op��o aparecer
	    // Se ela n�o aparecer dar� erro, da� sigo adiante
    	try {
    		
    		WebDriverWait waitFecharPopUpTelaInicial = new WebDriverWait(driver, 5);
    		popUpTelaInicial(waitFecharPopUpTelaInicial);
		
    	} catch (Exception e) {
    		System.out.println("Deu erro na opcao de fechar pop-up da tela inicial");
		}
    	
    }


    
	@SuppressWarnings({ "resource" })
	public static void lerPlanilhaContractNumbers(String planilha) throws Exception {

		try {
		   FileInputStream arquivo = new FileInputStream(new File(
				   planilha));
		
		   OPCPackage pkg = OPCPackage.open(new File(planilha));
		
		   XSSFWorkbook workbook = new XSSFWorkbook(pkg);
		   
		   XSSFSheet sheetRelatorio = workbook.getSheetAt(0);
		   
		   // Uso o DataFormatter para deixar todos os campos como String, inclusive
		   // os que tem n�meros
		   DataFormatter formatter = new DataFormatter();
		   for (int i=0; i <= sheetRelatorio.getLastRowNum(); i++) {
		       Row row = sheetRelatorio.getRow(i);
		       
		       if (row != null) {
		    	   
		    	   if (row.getRowNum() == 0) {
		    		   continue;
		    	   }
		    	   
		           Cell contrato = row.getCell(0);
		           Cell frente = row.getCell(1);
		           Cell contratoSap = row.getCell(2);
		           
		           boolean contratoPossuiValor = contrato != null && contrato.toString() != null && !contrato.toString().isEmpty();
		           boolean frentePossuiValor = frente != null && frente.toString() != null && !frente.toString().isEmpty();
		           boolean contratoSapPossuiValor = contratoSap != null && contratoSap.toString() != null && !contratoSap.toString().isEmpty();
		           // Pode ser que existam linhas que aparentemente est�o vazias, mas possuem conte�do em branco
		           if (contratoPossuiValor && frentePossuiValor && contratoSapPossuiValor) {
		        	   
		        	   ContractNumber contractNumber = new ContractNumber();
		        	   
		        	   // Contrato
		        	   if (formatter.formatCellValue(row.getCell(0)) != null && !formatter.formatCellValue(row.getCell(0)).isEmpty()) {
		        		   contractNumber.setContrato(formatter.formatCellValue(row.getCell(0)).trim());
		        	   }
		        	   // Frente
		        	   if (formatter.formatCellValue(row.getCell(1)) != null && !formatter.formatCellValue(row.getCell(1)).isEmpty()) {
		        		   contractNumber.setFrente(formatter.formatCellValue(row.getCell(1)).trim());
		        	   }
		        	   // Contrato Sap
		        	   if (formatter.formatCellValue(row.getCell(2)) != null && !formatter.formatCellValue(row.getCell(2)).isEmpty()) {
		        		   contractNumber.setNumero(formatter.formatCellValue(row.getCell(2)).trim());
		        	   }
		        	   // WBS
		        	   if (formatter.formatCellValue(row.getCell(3)) != null && !formatter.formatCellValue(row.getCell(3)).isEmpty()) {
		        		   contractNumber.setWbs(formatter.formatCellValue(row.getCell(3)).trim());
		        	   }

		        	   // Status_Contrato
		        	   // O rob� ir� processar os contract numbers que possu�rem status Ativo no campo Status_Contrato da planilha Rel_Projeto.xlsx"
		        	   String status = formatter.formatCellValue(row.getCell(11));
		        	   if (status != null && !status.isEmpty() && "Ativo".equals(status)) {
		        		   listaNumerosContractNumbersDistintos.add(contractNumber.getNumero());
		        		   listaContractNumbersTemporaria.add(contractNumber);
		        	   }
		        	   
		           }
		    	   
		       }
		       
		   }
		   
		   arquivo.close();
		
			} catch (FileNotFoundException e) {
			   e.printStackTrace();
			   System.out.println("Arquivo Excel de relatorio nao encontrado!");
			   throw new Exception("Arquivo Excel de relatorio nao encontrado!");
			}
		
			if (listaContractNumbersTemporaria.size() == 0) {
			   throw new Exception("Lista de contract numbers esta vazia");
			}
			
		}
	
	public static void criaListaContractNumbersDistintos() throws Exception {
		
		if (listaNumerosContractNumbersDistintos != null && !listaNumerosContractNumbersDistintos.isEmpty()) {
			
			for (String numeroContractNumberDistinto : listaNumerosContractNumbersDistintos) {
				
				if (numeroContractNumberDistinto != null && !numeroContractNumberDistinto.isEmpty()) {
					
					if (listaContractNumbersTemporaria != null && !listaContractNumbersTemporaria.isEmpty()) {
						
						for (ContractNumber contractNumber : listaContractNumbersTemporaria) {
							
							if (contractNumber != null) {
								
								if (numeroContractNumberDistinto.equals(contractNumber.getNumero())) {
									listaContractNumbers.add(contractNumber);
									break;
								}
							}
							
						}
					}

				}
				
			}
			
		}
		
	}
	
    @SuppressWarnings("resource")
	public static void lerRelatorioExcel(WebDriver driver, WebDriverWait wait, JavascriptExecutor js, String relatorio, String subdiretorioRelatoriosBaixados, String usuario, String senha) throws Exception {

        try {
               FileInputStream arquivo = new FileInputStream(new File(
            		   relatorio));
               
               File arquivoExcel = new File(relatorio);
               
               if (arquivoExcel.exists() && arquivoExcel.isFile() && arquivoExcel.length() > 0) {
            	   
            	   OPCPackage pkg = OPCPackage.open(new File(relatorio));
            	   
            	   XSSFWorkbook workbook = new XSSFWorkbook(pkg);
            	   
            	   XSSFSheet sheetPedidos =  workbook.getSheetAt(0);
            	   
            	   Iterator<Row> rowIterator = sheetPedidos.iterator();
            	   
            	   Pedido pedido = null;
            	   String azul = "FF99CCFF";
            	   String verde = "FFCCFFCC";
            	   
            	   List<Pedido> listaPedidosDeLinhaAzul = new ArrayList<Pedido>();
            	   List<Pedido> listaPedidosDeLinhaVerde = new ArrayList<Pedido>();
            	   
            	   while (rowIterator.hasNext()) {
            		   
            		   Row row = rowIterator.next();
            		   
            		   if (row.getRowNum() == 0 || row.getRowNum() == 1 || row.getRowNum() == 2) {
            			   continue;
            		   }
            		   
            		   Iterator<Cell> cellIterator = row.cellIterator();
        			   pedido = new Pedido();
        			   boolean isAzul = false;
        			   boolean isVerde = false;
        			   
        			   while (cellIterator.hasNext()) {
        				   Cell cell = cellIterator.next();
        				   
        				   XSSFCellStyle cellStyle = (XSSFCellStyle) cell.getCellStyle();
        				   XSSFColor color = cellStyle.getFillForegroundColorColor();
        				   
        				   // Na planilha baixada do Adquira, um pedido � composto por informa��es que est�o na linha azul e na linha verde.
        				   // Separo essas informa��es em duas listas para depois criar uma lista �nica

        				   if (cell.getColumnIndex() == 0 && azul.equals(((XSSFColor)color).getARGBHex()) || cell.getColumnIndex() == 1 && azul.equals(((XSSFColor)color).getARGBHex()) ||
        					   cell.getColumnIndex() == 2 && azul.equals(((XSSFColor)color).getARGBHex()) || cell.getColumnIndex() == 4 && azul.equals(((XSSFColor)color).getARGBHex()) ||
        					   cell.getColumnIndex() == 5 && azul.equals(((XSSFColor)color).getARGBHex()) || cell.getColumnIndex() == 6 && azul.equals(((XSSFColor)color).getARGBHex()) ||
        					   cell.getColumnIndex() == 11 && azul.equals(((XSSFColor)color).getARGBHex())) {
        					   
        					   isAzul = true;
        					   
        					   switch (cell.getColumnIndex()) {
        					   case 0:
        						   pedido.setNumero(cell.getStringCellValue());
        						   break;
        					   case 1:
        						   pedido.setComprador(cell.getStringCellValue());
        						   break;
        					   case 2:
        						   pedido.setCnpjCliente(cell.getStringCellValue());
        						   break;       
        					   case 4:
        						   pedido.setData(formatarDataPedido(cell.getStringCellValue()));
        						   break;
        					   case 5:
        						   pedido.setValor(formatarValorPedido(cell.getStringCellValue()));
        						   break;
        					   case 6:
        						   pedido.setEstado(cell.getStringCellValue());
        						   break;
        					   case 11:
        						   pedido.setPrazoPagamento(cell.getStringCellValue());
        						   break;	   
        					   }
        					   
        				   } else if (cell.getColumnIndex() == 0 && verde.equals(((XSSFColor)color).getARGBHex()) || cell.getColumnIndex() == 1 && verde.equals(((XSSFColor)color).getARGBHex()) ||
        						      cell.getColumnIndex() == 2 && verde.equals(((XSSFColor)color).getARGBHex())) {
        					   
        					   isVerde = true;
        					   
        					   switch (cell.getColumnIndex()) {
        					   case 0:
        						   pedido.setNumero(cell.getStringCellValue());
        						   break;
        					   case 1:
        						   pedido.setItem(Integer.valueOf(cell.getStringCellValue()));
        						   break;  
        					   case 2:
        						   pedido.setCap(cell.getStringCellValue());
        						   break;       
        						   
        					   }
        					   
        				   }
        				   
        			   }
        			   
		        	   // Data Extra��o
		        	    pedido.setDataExtracao(dataAtualPlanilhaFinal);
            		   
        			   if (isAzul) {
        				   
        				   listaPedidosDeLinhaAzul.add(pedido);
        			   
        			   } else if (isVerde) {
        				   
        				   listaPedidosDeLinhaVerde.add(pedido);
        			   }
            		   
            	   }
            	   arquivo.close();
            	   
            	   // Unifica os pedidos de linha azul com os pedidos de linha verde na lista listaPedidos
            	   unificaPedidosdeLinhaAzulComPedidosDeLinhaVerde(listaPedidosDeLinhaAzul, listaPedidosDeLinhaVerde, listaPedidos);
               }

		} catch (Exception e) {
     	   contadorErroslerRelatorioExcel ++;
            Thread.sleep(3000);
            System.out.println("Arquivo Excel nao encontrado! Tentando resolver, tentativa de numero: " + contadorErroslerRelatorioExcel);
            
            // Tento ler o arquivo por at� 20 vezes
            if (contadorErroslerRelatorioExcel <= 20) {
            	
				System.out.println("Deu erro no metodo lerRelatorioExcel, tentativa de acerto: " + contadorErroslerRelatorioExcel);
				// Est� dando erro de logout no servidor
				// O bot�o de logout est� ficando escondido
				// ent�o retirarei o logout e o login por enquanto
				fazerLogoutAdquira(driver, wait);
				fazerLoginAdquira(driver, wait, js, usuario, senha);
				//acessarPaginaInicial(driver, wait);
				fazerDownlodRelatorioPorPeriodo(driver, wait, js, usuario, senha);
				moverArquivosEntreDiretorios(driver, wait, js, Util.getValor("caminho.download.relatorios") + "\\" + nomeRelatorioBaixado, subdiretorioRelatoriosBaixados, usuario, senha);
         	    lerRelatorioExcel(driver, wait, js, relatorio, subdiretorioRelatoriosBaixados, usuario, senha);
            
            } else {
         	   throw new Exception("Arquivo Excel nao encontrado! : " + e);
            }
		}
        
        if (listaPedidos.size() == 0) {
        	   throw new Exception("Lista de pedidos esta vazia");
        }
        
  }
    
    public static void unificaPedidosdeLinhaAzulComPedidosDeLinhaVerde(List<Pedido> listaPedidosDeLinhaAzul, List<Pedido> listaPedidosDeLinhaVerde, List<Pedido> listaPedidos) throws Exception  { 
    	
    	if (listaPedidosDeLinhaAzul != null && !listaPedidosDeLinhaAzul.isEmpty()) {
    		
    		for (Pedido pedidoAzul : listaPedidosDeLinhaAzul) {
    			
    			List<Integer> listaItem = new ArrayList<Integer>();

    			List<String> listaCap = new ArrayList<String>();
    			
    			if (pedidoAzul != null) {
    				
    				if (listaPedidosDeLinhaVerde != null && !listaPedidosDeLinhaVerde.isEmpty()) {
    					
    					for (Pedido pedidoVerde : listaPedidosDeLinhaVerde) {
    						
    						if (pedidoVerde != null) {
    							
    							if (pedidoAzul.getNumero().trim().equals(pedidoVerde.getNumero().trim())) {
    								
    								listaItem.add(pedidoVerde.getItem());
    								listaCap.add(pedidoVerde.getCap());
    								
    							}
    							
    						}
    						
    					}
    					
    				}
    				
    				pedidoAzul.setListaItem(listaItem);
    				pedidoAzul.setListaCap(listaCap);

    				if (!"CANCELADO".equalsIgnoreCase(pedidoAzul.getEstado()) && pedidoAzul.getPrazoPagamento().contains("M75") ||
    					!"CANCELADO".equalsIgnoreCase(pedidoAzul.getEstado()) && pedidoAzul.getPrazoPagamento().contains("ME75") ||
    					!"CANCELADO".equalsIgnoreCase(pedidoAzul.getEstado()) && pedidoAzul.getPrazoPagamento().contains("M60")) {
    				   
    				   listaPedidos.add(pedidoAzul);
    			   }

    			}
    			
			}
    		
    	}
    	
    }
    
    public static boolean isPedidoComContractNumberInvalido(Pedido pedido) throws Exception  { 
    	
    	boolean isPedidoComContractNumberInvalido = false;
    	String[] contractNumbers = Util.getValor("contract.numbers.ignorados.adquira").split(",");
    	
    	if (contractNumbers.length > 0) {
    		
    		for (String contractNumber : contractNumbers) {
    			
    			if (contractNumber != null && !contractNumber.isEmpty()) {
    				
    				if (contractNumber.trim().equals(pedido.getContractNumber().getNumero().trim())) {
    					
    					isPedidoComContractNumberInvalido =  true;
    					break;
    					
    				}
    				
    			}
    			
    		}
    		
    	}
    	    	
    	return isPedidoComContractNumberInvalido;
    	
    }
    
	@SuppressWarnings({ "resourc" })
	public static void lerPlanilhaPedidosFaturados(String planilha) throws Exception {

		try {
		   FileInputStream arquivo = new FileInputStream(new File(
				   planilha));
		
		   OPCPackage pkg = OPCPackage.open(new File(planilha));
		
		   XSSFWorkbook workbook = new XSSFWorkbook(pkg);
		   
		   XSSFSheet sheetPedidosFaturados = workbook.getSheet("Raw_Data");
		   
		   // Uso o DataFormatter para deixar todos os campos como String, inclusive
		   // os que tem n�meros
		   DataFormatter formatter = new DataFormatter();
		   
		   for (int i=0; i <= sheetPedidosFaturados.getLastRowNum(); i++) {
		       Row row = sheetPedidosFaturados.getRow(i);
		       
		       if (row != null) {
		    	   
		    	   Pedido pedido = new Pedido();
		    	   
		    	   boolean ignorarLinhas = row.getRowNum() == 0; 
		    	   
		    	   if (ignorarLinhas) {
		    		   continue;
		    	   }
		    	   
		           Cell poPedido = row.getCell(14);
		           boolean poPedidoPossuioValor = poPedido != null && poPedido.toString() != null && !poPedido.toString().isEmpty();
		
		           // Pode ser que existam linhas que aparentemente est�o vazias, mas possuem conte�do em branco
		           if (poPedidoPossuioValor) {
		        	   
		        	   // PO/Pedido
		        	   if (formatter.formatCellValue(poPedido) != null && !formatter.formatCellValue(poPedido).isEmpty()) {
		        		   pedido.setNumero(formatter.formatCellValue(poPedido).trim());
		        	   }

		            	listaPedidosFaturados.add(pedido);
		        	   
		           }
		    	   
		       }
		       
		   }
		   
		   arquivo.close();
		
			} catch (FileNotFoundException e) {
			   e.printStackTrace();
			   System.out.println("Arquivo Excel de pedidos faturados nao encontrado!");
			   throw new Exception("Arquivo Excel de pedidos faturados nao encontrado!");
			}
		
			if (listaPedidosFaturados.size() == 0) {
			   throw new Exception("Lista de pedidos faturados esta vazia");
			}
			
		}
    
    @SuppressWarnings({ "deprecation", "resource" })
	public static void criarRelatorioFinal(String relatorioPedidos) throws Exception {
    	  
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheetPedidos = workbook.createSheet("Pedidos");
        
        String contrato = "Contrato";
        String frente = "Frente";
        String contratoSap = "Contrato SAP";
        String wbs = "WBS";
        String numeroPedido = "Número Pedido";
        String dataPedido = "Data Pedido";
        String valorPedido = "Valor Pedido";
        String cnpjCliente = "Cnpj Cliente";
        String comprador = "Comprador";
        String faturado = "Faturado";
        String salvoNoSharepoint = "Salvo no Sharepoint";
        String observacaoSharepoint = "Campo Observacao no Sharepoint";
        String dataExtracao = "Data Extracao";
        String errosNoPedido = "Erros no Pedido";

        sheetPedidos.setColumnWidth(0, contrato.length() * 380);
        sheetPedidos.setColumnWidth(1, frente.length() * 450);
        sheetPedidos.setColumnWidth(2, contratoSap.length() * 260);
        sheetPedidos.setColumnWidth(3, wbs.length() * 900);
        sheetPedidos.setColumnWidth(4, numeroPedido.length() * 280);
        sheetPedidos.setColumnWidth(5, dataPedido.length() * 260);
        sheetPedidos.setColumnWidth(6, valorPedido.length() * 250);
        sheetPedidos.setColumnWidth(7, cnpjCliente.length() * 380);
        sheetPedidos.setColumnWidth(8, comprador.length() * 500);
        sheetPedidos.setColumnWidth(9, faturado.length() * 380);
        sheetPedidos.setColumnWidth(10, salvoNoSharepoint.length() * 280);
        sheetPedidos.setColumnWidth(11, observacaoSharepoint.length() * 380);
        sheetPedidos.setColumnWidth(12, dataExtracao.length() * 380);
        sheetPedidos.setColumnWidth(13, errosNoPedido.length() * 280);

        int rownum = 0;
        int cellnumero = 0;
        // Configura��o do Style da c�lula
        XSSFCellStyle styleNegrito = workbook.createCellStyle();
        // Negrito
        XSSFFont font = workbook.createFont();
        font.setBold(true);
        // C�lula com fundo cinza
        XSSFColor myColor = new XSSFColor(Color.LIGHT_GRAY);
        styleNegrito.setFont(font);
        styleNegrito.setFillForegroundColor(myColor);
        styleNegrito.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        
        // Cor cinza da c�lula
        //XSSFCellStyle styleCinza = workbook.createCellStyle();
        //XSSFColor myColor = new XSSFColor(Color.LIGHT_GRAY);
        //styleCinza.setFillForegroundColor(myColor);
        //style.setFillBackgroundColor(myColor);
        //styleCinza.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        
        // Cabe�alho
        Row row = sheetPedidos.createRow(rownum++);
        Cell cellCabecalhoContrato = row.createCell(cellnumero++);
        cellCabecalhoContrato.setCellValue(contrato);
        cellCabecalhoContrato.setCellStyle(styleNegrito);
        
        Cell cellCabecalhoFrente = row.createCell(cellnumero++);
        cellCabecalhoFrente.setCellValue(frente);
        cellCabecalhoFrente.setCellStyle(styleNegrito);
        
        Cell cellCabecalhoContratoSap = row.createCell(cellnumero++);
        cellCabecalhoContratoSap.setCellValue(contratoSap);
        cellCabecalhoContratoSap.setCellStyle(styleNegrito);
        
        Cell cellCabecalhoWbs = row.createCell(cellnumero++);
        cellCabecalhoWbs.setCellValue(wbs);
        cellCabecalhoWbs.setCellStyle(styleNegrito);

        Cell cellCabecalhoNumeroPedido = row.createCell(cellnumero++);
        cellCabecalhoNumeroPedido.setCellValue(numeroPedido);
        cellCabecalhoNumeroPedido.setCellStyle(styleNegrito);
        
        Cell cellCabecalhoData = row.createCell(cellnumero++);
        cellCabecalhoData.setCellValue(dataPedido);
        cellCabecalhoData.setCellStyle(styleNegrito);
        
        Cell cellCabecalhoValor = row.createCell(cellnumero++);
        cellCabecalhoValor.setCellValue(valorPedido);
        cellCabecalhoValor.setCellStyle(styleNegrito);

        Cell cellCabecalhoCnpjCliente = row.createCell(cellnumero++);
        cellCabecalhoCnpjCliente.setCellValue(cnpjCliente);
        cellCabecalhoCnpjCliente.setCellStyle(styleNegrito);
        
        Cell cellCabecalhoComprador = row.createCell(cellnumero++);
        cellCabecalhoComprador.setCellValue(comprador);
        cellCabecalhoComprador.setCellStyle(styleNegrito);
        
        Cell cellCabecalhoFaturado = row.createCell(cellnumero++);
        cellCabecalhoFaturado.setCellValue(faturado);
        cellCabecalhoFaturado.setCellStyle(styleNegrito);
        
        Cell cellCabecalhoSalvoNoSharepoint = row.createCell(cellnumero++);
        cellCabecalhoSalvoNoSharepoint.setCellValue(salvoNoSharepoint);
        cellCabecalhoSalvoNoSharepoint.setCellStyle(styleNegrito);
        
        Cell cellCabecalhoObservacaoSharepoint = row.createCell(cellnumero++);
        cellCabecalhoObservacaoSharepoint.setCellValue(observacaoSharepoint);
        cellCabecalhoObservacaoSharepoint.setCellStyle(styleNegrito);
        
        Cell cellCabecalhoDataExtracao = row.createCell(cellnumero++);
        cellCabecalhoDataExtracao.setCellValue(dataExtracao);
        cellCabecalhoDataExtracao.setCellStyle(styleNegrito);
        
        Cell cellCabecalhoErrosNoPedido = row.createCell(cellnumero++);
        cellCabecalhoErrosNoPedido.setCellValue(errosNoPedido);
        cellCabecalhoErrosNoPedido.setCellStyle(styleNegrito);

        
        for (Pedido pedidosNaoFaturados : listaPedidosNaoFaturados) {
        	
            row = sheetPedidos.createRow(rownum++);
            int cellnum = 0;
            
            Cell cellContrato = row.createCell(cellnum++);
            cellContrato.setCellValue(pedidosNaoFaturados.getContractNumber().getContrato());
            
            Cell cellFrente = row.createCell(cellnum++);
            cellFrente.setCellValue(pedidosNaoFaturados.getContractNumber().getFrente());
            
            Cell cellContratoSap = row.createCell(cellnum++);
            cellContratoSap.setCellValue(pedidosNaoFaturados.getContractNumber().getNumero());
            
            Cell cellContractWbs = row.createCell(cellnum++);
            cellContractWbs.setCellValue(pedidosNaoFaturados.getContractNumber().getWbs());
            
            Cell cellNumeroPedido = row.createCell(cellnum++);
            cellNumeroPedido.setCellValue(pedidosNaoFaturados.getNumero());

            Cell cellData = row.createCell(cellnum++);
            cellData.setCellValue(pedidosNaoFaturados.getData());
            
            Cell cellValor = row.createCell(cellnum++);
            cellValor.setCellValue(pedidosNaoFaturados.getValor());

            Cell cellCnpjCliente = row.createCell(cellnum++);
            cellCnpjCliente.setCellValue(pedidosNaoFaturados.getCnpjCliente());
            
            Cell cellComprador = row.createCell(cellnum++);
            cellComprador.setCellValue(pedidosNaoFaturados.getComprador());
            
            Cell cellFaturado = row.createCell(cellnum++);
            cellFaturado.setCellValue( pedidosNaoFaturados.isFaturado() ? "Sim" : "Nao" );
            
            Cell cellSalvoNoSharepoint = row.createCell(cellnum++);
            // Se o pedido n�o est� faturado, mostro a informa��o se foi salvo no sharepoint
            if (!pedidosNaoFaturados.isFaturado()) {
            	cellSalvoNoSharepoint.setCellValue( pedidosNaoFaturados.isSalvoNoSharepoint() ? "Sim" : "Nao" );
            	// Se o pedido est� faturado, subentende-se que est� salvo no sharepoint
            } else {
            	cellSalvoNoSharepoint.setCellValue("Sim");
            }
            
            Cell cellObservacaoSharepoint = row.createCell(cellnum++);
            cellObservacaoSharepoint.setCellValue(pedidosNaoFaturados.getObservacaoSharepoint());
            
            Cell cellDataExtracao = row.createCell(cellnum++);
            cellDataExtracao.setCellValue(pedidosNaoFaturados.getDataExtracao());
            
            Cell cellErrosNoPedido = row.createCell(cellnum++);
            preencherMensagemDeErroNoPedido(pedidosNaoFaturados);
            cellErrosNoPedido.setCellValue(pedidosNaoFaturados.getMensagemDeErroNoPedido());
            
        }
          
        try {
            FileOutputStream out = 
                    new FileOutputStream(new File(relatorioPedidos));
            workbook.write(out);
            out.close();
            System.out.println("Arquivo Excel criado com sucesso!");
              
        } catch (FileNotFoundException e) {
            e.printStackTrace();
               System.out.println("Arquivo nao encontrado!");
        } catch (IOException e) {
            e.printStackTrace();
               System.out.println("Erro na edicao do arquivo!");
        }

  }
    
	@SuppressWarnings({ "resource" })
	public static int retornaUltimaLinhaPlanilhaRelatorioIncremental(String relatorioIncremental) throws Exception {

		int ultimaLinhaPlanilhaRelatorioIncremental = 0;

		try {
		   FileInputStream arquivo = new FileInputStream(new File(
				   relatorioIncremental));
		
		   OPCPackage pkg = OPCPackage.open(new File(relatorioIncremental));
		
		   XSSFWorkbook workbook = new XSSFWorkbook(pkg);
		   
		   XSSFSheet sheetRelatorioIncremental = workbook.getSheet("Pedidos");
		   
		   
		   for (int i=0; i <= sheetRelatorioIncremental.getLastRowNum(); i++) {
		       Row row = sheetRelatorioIncremental.getRow(i);
		       
		       if (row != null) {
		    	   
		    	   Cell contrato = row.getCell(0);
		    	   Cell frente = row.getCell(1);
		    	   Cell contratoSap = row.getCell(2);

		    	   boolean contratoPossuiValor = contrato != null && contrato.toString() != null && !contrato.toString().isEmpty();
		    	   boolean frentePossuiValor = frente != null && frente.toString() != null && !frente.toString().isEmpty();
		    	   boolean contratoSapPossuiValor = contratoSap != null && contratoSap.toString() != null && !contratoSap.toString().isEmpty();

		    	   boolean ignorarLinhas = contratoPossuiValor && frentePossuiValor && contratoSapPossuiValor;
		    	   
		    	   // Preciso pegar a �ltima linha com informa��o, ent�o ignoro as que tem informa��o
		    	   if (ignorarLinhas) {
		    		   ultimaLinhaPlanilhaRelatorioIncremental = i;
		    		   continue;		    		   
		    	   }
		    	   
		       }
		       
		   }
		   
		   arquivo.close();
		
			} catch (FileNotFoundException e) {
			   e.printStackTrace();
			   System.out.println("Arquivo Excel do relatorio incremental nao encontrado!");
			   throw new Exception("Arquivo Excel do relatorio incremental nao encontrado!");
			}
		
		return ultimaLinhaPlanilhaRelatorioIncremental;
		
		}
    
    @SuppressWarnings({ "resource" })
	public static void preenchePlanilhaRelatorioIncremental(String relatorioIncremental) throws Exception {
        
    	try  {
    		
    		XSSFWorkbook workbook;
        	FileInputStream file = new FileInputStream(new File(relatorioIncremental));
            
        	workbook = new XSSFWorkbook(file);
            XSSFSheet sheetRelatorioIncremental = workbook.getSheet("Pedidos");
            
            int rownum = retornaUltimaLinhaPlanilhaRelatorioIncremental(relatorioIncremental) + 1;

            for (Pedido pedidosFaturadosENaoFaturados : listaPedidos) {

            	int cellnum = 0;
            	
            	Row row = sheetRelatorioIncremental.createRow(rownum++);
                
                Cell cellContrato = row.createCell(cellnum++);
                cellContrato.setCellValue(pedidosFaturadosENaoFaturados.getContractNumber().getContrato());
                
                Cell cellFrente = row.createCell(cellnum++);
                cellFrente.setCellValue(pedidosFaturadosENaoFaturados.getContractNumber().getFrente());
                
                Cell cellContratoSap = row.createCell(cellnum++);
                cellContratoSap.setCellValue(pedidosFaturadosENaoFaturados.getContractNumber().getNumero());
                
                Cell cellContractWbs = row.createCell(cellnum++);
                cellContractWbs.setCellValue(pedidosFaturadosENaoFaturados.getContractNumber().getWbs());
                
                Cell cellNumeroPedido = row.createCell(cellnum++);
                cellNumeroPedido.setCellValue(pedidosFaturadosENaoFaturados.getNumero());

                Cell cellData = row.createCell(cellnum++);
                cellData.setCellValue(pedidosFaturadosENaoFaturados.getData());
                
                Cell cellValor = row.createCell(cellnum++);
                cellValor.setCellValue(pedidosFaturadosENaoFaturados.getValor());

                Cell cellCnpjCliente = row.createCell(cellnum++);
                cellCnpjCliente.setCellValue(pedidosFaturadosENaoFaturados.getCnpjCliente());
                
                Cell cellComprador = row.createCell(cellnum++);
                cellComprador.setCellValue(pedidosFaturadosENaoFaturados.getComprador());
                
                Cell cellFaturado = row.createCell(cellnum++);
                cellFaturado.setCellValue( pedidosFaturadosENaoFaturados.isFaturado() ? "Sim" : "Nao" );
                
                Cell cellSalvoNoSharepoint = row.createCell(cellnum++);
                // Se o pedido n�o est� faturado, mostro a informa��o se foi salvo no sharepoint
                if (!pedidosFaturadosENaoFaturados.isFaturado()) {
                	cellSalvoNoSharepoint.setCellValue( pedidosFaturadosENaoFaturados.isSalvoNoSharepoint() ? "Sim" : "Nao" );
                	// Se o pedido est� faturado, subentende-se que est� salvo no sharepoint
                } else {
                	cellSalvoNoSharepoint.setCellValue("Sim");
                }
                
                Cell cellProblemaNaRegraDePreenchimentoNoSharepoint = row.createCell(cellnum++);
                cellProblemaNaRegraDePreenchimentoNoSharepoint.setCellValue(pedidosFaturadosENaoFaturados.getMensagemErroRegraPreenchimento());
                
                Cell cellObservacaoSharepoint = row.createCell(cellnum++);
                cellObservacaoSharepoint.setCellValue(pedidosFaturadosENaoFaturados.getObservacaoSharepoint());
                
                Cell cellDataExtracao = row.createCell(cellnum++);
                cellDataExtracao.setCellValue(pedidosFaturadosENaoFaturados.getDataExtracao());
                
            }
            
             file.close();

            FileOutputStream outFile = new FileOutputStream(new File(relatorioIncremental));
            workbook.write(outFile);
            outFile.close();
            System.out.println("Arquivo Excel editado com sucesso!");
	    
		} catch (FileNotFoundException e) {
			   System.out.println("Arquivo Excel do relatorio incremental nao encontrado!");
		} catch (IOException e) {
		        System.out.println("Erro na edicao do relatorio incremental nao encontrado!");
		}
	
    }

    
	@SuppressWarnings("resource")
	public static void descompactaArquivoZip(String arquivoZip, String arquivoPdf) {

		try {

			FileInputStream file = new FileInputStream(arquivoZip);

			ZipInputStream zis = new ZipInputStream(file);

			ZipEntry zipElement = zis.getNextEntry();

			while (zipElement != null) {

				try {
					FileOutputStream out = new FileOutputStream(new File(arquivoPdf));

					//byte[] buffer = new byte[Math.toIntExact(zipElement.getSize())];
					byte [] buffer = new byte[1024];

					int location;

					while ((location = zis.read(buffer)) != -1) {
						out.write(buffer, 0, location);
					}

					System.out.println("Arquivo zip descompactado com sucesso");

				} catch (FileNotFoundException e) {
					e.printStackTrace();
					System.out.println("Arquivo zip nao encontrado!");
				} catch (IOException e) {
					e.printStackTrace();
					System.out.println("Erro na edicao do arquivo!");
				}

				zipElement = zis.getNextEntry();
			}
		} catch (IOException e) {
			e.printStackTrace();
			throw new RuntimeException(e);
		}

	}

    
    
    public static String formatarValorPedido(String valorPedido) throws Exception {
    	
    	double valorPedidoDouble = 0;
    	String valorPedidoString = "";
    	
    	if (valorPedido != null && !valorPedido.isEmpty()) {
    		// Transformo, por exemplo, o valor BRL 1.094.435,93 em 1094435.93
    		valorPedido = valorPedido.replace("BRL", "").replaceAll("\\.", "").replaceAll("\\,", "\\.");
    		valorPedidoDouble = Double.valueOf(valorPedido);
    		valorPedidoString = String.format("%.2f", new BigDecimal(valorPedidoDouble));
    		
    		if (valorPedidoString != null && !valorPedidoString.isEmpty() && valorPedidoString.contains(",")) {
    			valorPedidoString = valorPedidoString.replace(",", ".");
    		}
    	}
    	
    	return valorPedidoString;
    	
    }
    
    public static String recuperaCap(Pedido pedido) throws Exception {
    	
    	String listaCap = "";
    	
    	if (pedido != null && pedido.getListaCap().size() > 0) {
    		
    		if (pedido.getListaCap() != null && !pedido.getListaCap().isEmpty()) {
    			
    			if (pedido.getListaCap().size() > 0) {
    				
    				for (String capAdquira : pedido.getListaCap()) {
    					
    					String cap = formataCap(capAdquira);
    					
    					if (capAdquira.equals(pedido.getListaCap().get(pedido.getListaCap().size()-1))) {
    						listaCap = listaCap + cap;
    					} else {
    						listaCap = listaCap + cap + ";";
    					}
    					
    				}

    			}
    		
    		}
    		
    		
    	}
		
    	return listaCap;

    }
    
    public static String formataCap(String cap) throws Exception {
    	
    	if (cap != null && !cap.isEmpty()) {
    		
    		if (cap.contains("CAP")) {
    			
    			if (cap.indexOf("CAP") != -1) {
    				
    				int indiceCap = cap.indexOf("CAP");

    				 String finalCap = cap.substring(indiceCap, cap.length());
    				 
    				 if (finalCap != null && !finalCap.isEmpty()) {
    					 
    					 // Retirando os caracteres que n�o forem n�meros
    					 String somenteNumeros = finalCap.replaceAll("[^0-9]", "");
    					 
    					 // Se depois da palavra CAP tivermos n�meros, retornaremos esse CAP seguido de n�meros
    					 // Se n�o, retornamos o CAP completo que vem da planilha do Adquira
    					 if (somenteNumeros != null && !somenteNumeros.isEmpty()) {
    						 
    						 cap = finalCap;
    					 
    					 }

    				 }
    				 
    			}
    			
    		}
    		
    	}
    	
    	return cap;
    	
    }
    
    public static String formatarDataPedido(String dataPedido) throws Exception {
    	
    	Date date = null;
    	try {
    		
    		// Formato que vem do excel baixado: 14-oct-2019 que � a data em espanhol
    		// O formato abaixo parou de funcionar e n�o sei porque kkk
    		// Estou ent�o convertendo o m�s para n�mero
    		// DateFormat formatter = new SimpleDateFormat("DD-MMM-YYYY", new Locale("es", "ES"));
    		
    		dataPedido = retornaDataComMesEmNumero(dataPedido);
    		DateFormat formatter = new SimpleDateFormat("dd-MM-yyyy");
    		
    		// Transformo em date
    		date = formatter.parse(dataPedido);
    		
    		// Converto a data para o formato novo
    		SimpleDateFormat formatoNovo = new SimpleDateFormat("dd/MM/yyyy"); 
    		dataPedido = formatoNovo.format(date);
    		
    	} catch (ParseException e) {            
    		throw e;
    	}
    	
		return dataPedido; 
    	
    }
    
    public static String retornaDataComMesEmNumero (String dataPedido) {
    	
    	String[] partesDataPedido = dataPedido.split("-");
    	String dia = partesDataPedido[0];
    	String mes = partesDataPedido[1];
    	String ano = partesDataPedido[2];
    	
		   switch (mes) {
		   case "ene":
			   mes = "01";
			   break;
		   case "feb":
			   mes = "02";
			   break;
		   case "mar":
			   mes = "03";
			   break;       
		   case "abr":
			   mes = "04";
			   break;
		   case "may":
			   mes = "05";
			   break;
		   case "jun":
			   mes = "06";
			   break;
		   case "jul":
			   mes = "07";
			   break;
		   case "ago":
			   mes = "08";
			   break;
		   case "sep":
			   mes = "09";
			   break;
		   case "oct":
			   mes = "10";
			   break;       
		   case "nov":
			   mes = "11";
			   break;
		   case "dic":
			   mes = "12";
			   break;

		   }

	    dataPedido = dia + "-" + mes + "-" + ano;
    	
    	return dataPedido;
    }
    
    public static String dataDaNotaSharepoint(Date dataInclusaoNota) throws Exception {
    	
    	//A regra para a data da nota no sharepoint � esta:
    	//	1) Adiciona-se 75 dias � data de inclus�o da nota.
    	//	2) A data da nota no sharepoint ser� de acordo com os seguintes dias de vencimento: 4, 12 e 22.
    	//	Ex: Nota inclu�da dia 15/10/20 + 75 dias = 29/12/2020. Neste cen�rio a data de vencimento da nota no sharepoint ser� 04/01/2021.  
    	
    	String dataDaNotaSharepoint = null;
    	// Adicionando 75 dias na data de inclus�o da nota
		Calendar cal = Calendar.getInstance(); 
		cal.setTime(dataInclusaoNota); 
		cal.add(Calendar.DATE, 75);
		Date dataInclusaoNotaAcrescidaDe75Dias = cal.getTime();
		
		// Pegando dia, m�s e ano da data de inclus�o acrescida de 75 dias
		GregorianCalendar calendar = new GregorianCalendar();
		calendar.setTime(dataInclusaoNotaAcrescidaDe75Dias);
		int dia = calendar.get(GregorianCalendar.DAY_OF_MONTH);
		int mes = calendar.get(GregorianCalendar.MONTH);
		int ano = calendar.get(GregorianCalendar.YEAR);
		
		// Regra dos dias de vencimento 4, 12 e 22
		int diaVencimento = 0;
		Calendar cal2 = Calendar.getInstance();
		
		if (dia >= 1 && dia <= 4) {
		
			diaVencimento = 4;
			cal2.set(ano, mes, diaVencimento); 
		
		} else if (dia > 4 && dia <= 12) {
			
			diaVencimento = 12;
			cal2.set(ano, mes, diaVencimento); 
		
		} else if (dia > 12 && dia <= 22) {
			
			diaVencimento = 22;
			cal2.set(ano, mes, diaVencimento); 
		
		} else if (dia > 22 && dia <= 31) {
			
			diaVencimento = 4;
			
			// Neste caso do if, a data ser� a pr�xima do m�s.
			// Mas se o m�s for dezembro, ent�o o pr�ximo m�s ser� janeiro e o ano ser� o pr�ximo.
			// Dezembro
			if (mes == 11) {
				mes = 0;
				ano = ano +1;
			} else {
				mes = mes +1;
			}
			cal2.set(ano, mes, diaVencimento); 
			
		}
		
		dataDaNotaSharepoint = new SimpleDateFormat("MM/dd/yyyy").format(cal2.getTime());
		
		return dataDaNotaSharepoint; 
    	
    }
    
    public static void fechandoAbaEabrindoNova(WebDriver driver) throws Exception{
    	
    	Thread.sleep(3000);
    	// Abrindo uma aba nova
    	((JavascriptExecutor) driver).executeScript("window.open()");
    	Thread.sleep(1000);
    	
    	// Fechando a primeira aba
    	driver.close();
    	
    	// Lista de abas abertas
    	// Teremos uma janela somente aberta
    	List<String> windowHandles = new ArrayList<>(driver.getWindowHandles());
    	
    	// Embora s� exista uma janela aberta nesse momento,
    	// preciso setar o drive para esta �nica janela
    	driver.switchTo().window(windowHandles.get(0));
    	
    }
    
    // Fazendo o logout
    public static void fazerLogoutAdquira(WebDriver driver, WebDriverWait wait) throws Exception{
    	
    	try {
    		
    		//buscaAvancadaComUrl(driver, wait);
    		
    		//fechandoAbaEabrindoNova(driver);
    		
    		//buscaAvancadaComUrl(driver, wait);
    		
    	    fecharMensagemAceitarCookies(driver);
        	
    	    fecharMensagemVerNotificacoes(driver);
    		
    		wait.until(ExpectedConditions.elementToBeClickable(By.id("TID_USERBAR_BUTTON_USER_MENU"))).click();
    		Thread.sleep(3000);
    		
    		String textoCerrarSession = "Cerrar sesión";
    		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//span [text()='"+textoCerrarSession+"']"))).click();
    		Thread.sleep(3000);
    		
		} catch (Exception e) {
			contadorErrosLogout ++;
            Thread.sleep(3000);
	            
	            // Tento fazer o logout por at� 10 vezes
	            if (contadorErrosLogout <= 10) {
	            	
					System.out.println("Deu erro no metodo fazerLogout, tentativa de acerto: " + contadorErrosLogout);
					fazerLogoutAdquira(driver, wait);
	            
	            } else {
	         	   throw new Exception("Erro no Logout: " + e);
	            }
			}
    	
    }
    
    public static void mensagemSucesso() throws IOException {
    	
        String caminho = Util.getValor("caminho.executavel.automacao.sharepoint");
        
        Object[] options = { "Sim", "Nao" };
        int i = JOptionPane.showOptionDialog(null, "Extracao dos pedidos no Adquira executada com sucesso! \n\n Gostaria de iniciar a insercao dos pedidos no Sharepoint?", "Automatizacao Adquira", JOptionPane.YES_NO_OPTION, JOptionPane.QUESTION_MESSAGE, null, options, options[0]);
        
        if (i == JOptionPane.YES_OPTION) {
        	
        	Process process = Runtime.getRuntime().exec(caminho);
        	
        } else {
        	
        	System.exit(0); 
        }
    	
    }
    
    
    public static void mensagemErro(String mensagem) throws IOException {
	    String erro = "";
	
	    // Cria um JFrame
	    JFrame frame = new JFrame("JOptionPane exemplo");
	
	    // Cria o JOptionPane por showMessageDialog
	    JOptionPane.showMessageDialog(frame,
	    		mensagem + erro + "", //mensagem
	        "Automatizacao Adquira", // titulo da janela 
	        JOptionPane.INFORMATION_MESSAGE);
	    System.exit(0);
    }
    
    public static void preencherMensagemDeErroNoPedido(Pedido pedido) throws IOException {
    	
    	String mensagemNaoEncontrouPdfAnexo = "PDF nao encontrado no Adquira";
    	String mensagemNaoEncontrouContractNumberNoPdf = "Contrato SAP nao encontrado no PDF do Adquira";
    	String mensagemNaoEncontrouContractNumberCorrespondenteComAListaDeContractNumbersDoSharepoint = "Contrato SAP nao cadastrado no SharePoint";
    	
    	if (pedido != null) {
    		
    		if (!pedido.isEncontrouPdfAnexo()) {
    			pedido.setMensagemDeErroNoPedido(mensagemNaoEncontrouPdfAnexo);
    			
    		} else {
    			
    			if (!pedido.isContractNumberConforme()) {
    				
    				if ("-".equals(pedido.getContractNumber().getNumero())) {
    					
    					pedido.setMensagemDeErroNoPedido(mensagemNaoEncontrouContractNumberNoPdf);
    					
    				} else {
    					
    					pedido.setMensagemDeErroNoPedido(mensagemNaoEncontrouContractNumberCorrespondenteComAListaDeContractNumbersDoSharepoint);
    				}
    				
    			}
    			
    			
    		}
    		
    	}
    	
    }
    
    // Propiedades do driver para abrir no IE, Chrome ou Firefox
    public static WebDriver getWebDriver() throws InterruptedException {
    	
    	WebDriver driver = null;
    		
            try {
				
            	if ("Chrome".equals(Util.getValor("navegador"))) {
				    
					File file = new File(Util.getValor("driver.Chrome.selenium"));
					System.setProperty(Util.getValor("propriedade.sistema.para.driver.Chrome.selenium"), file.getAbsolutePath());
				    DesiredCapabilities caps = DesiredCapabilities.chrome();
				    caps.setJavascriptEnabled(true);
				    caps.setCapability("ignoreZoomSetting", true);
				    caps.setCapability("nativeEvents",false);
				    ChromeOptions chromeOptions = new ChromeOptions(); 
				    Map<String, Object> chromePreferences = new HashMap<String, Object>();
					chromePreferences.put("profile.default_content_settings.popups", 0);
				    chromePreferences.put("download.default_directory",Util.getValor("caminho.download.relatorios"));
				    chromePreferences.put("browser.helperApps.neverAsk.saveToDisk", "text/plain, application/vnd.ms-excel, text/csv, text/comma-separated-values, application/octet-stream");
				    chromeOptions.setExperimentalOption("prefs", chromePreferences);
				    
				    // Argumento que faz com que o navegador use os dados do usu�rio salvos
				    // Com isso n�o ser� necess�rio digitar os dados de login no sharepoint, pois ele pegar� as informa��es do usu�rio salvas na m�quina
				    // Um ponto importante � que n�o poderemos ter mais de uma sess�o do Chrome aberta
				    // Outro ponto importante � que a op��o acima browser.helperApps.neverAsk.saveToDisk que permite que o browser salve um arquivo sem perguntar aonde salvar,
				    // n�o funcionar� por conta do trecho abaixo.
				    // Neste caso deveremos setar manualmente essa op��o no Chrome antes de rodar o rob�
				    // Ser� necess�rio fazer aparecer essa pasta no explorer do usu�rio
				    chromeOptions.addArguments("user-data-dir=" + Util.getValor("caminho.dados.usuario.Chrome"));
				    chromeOptions.addArguments("--lang=pt");

				    // Com essa op��o, o Selenium executa tudo sem mostrar o navegador
				    // Por�m no Adquira n�o funciona
				    //chromeOptions.addArguments("--headless");
				    
				    driver = new ChromeDriver(chromeOptions);
				    // Limpa o cache usando m�todo do driver
				    driver.manage().deleteAllCookies();
				    
				    
				} else if ("internetExplorer".equals(Util.getValor("navegador"))) {
				
					File file = new File(Util.getValor("driver.internetExplorer.selenium"));
					System.setProperty(Util.getValor("propriedade.sistema.para.driver.internetExplorer.selenium"), file.getAbsolutePath());
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
				    // Limpa o cache usando m�todo do driver
				    driver.manage().deleteAllCookies();
				
				} else if ("Firefox".equals(Util.getValor("navegador"))) {
					
					File file = new File(Util.getValor("driver.Firefox.selenium"));
					System.setProperty(Util.getValor("propriedade.binario.Firefox.selenium"),Util.getValor("binario.Firefox")); 
					System.setProperty(Util.getValor("propriedade.sistema.para.driver.Firefox.selenium"),file.getAbsolutePath());
     				File profileDirectory = new File(Util.getValor("caminho.dados.usuario.Firefox"));
     				FirefoxProfile fxProfile = new FirefoxProfile(profileDirectory);
				    fxProfile.setPreference("browser.download.folderList",2);
				    fxProfile.setPreference("browser.download.manager.showWhenStarting",false);
				    fxProfile.setPreference("browser.download.dir",Util.getValor("caminho.download.relatorios"));
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
				    // Limpa o cache usando m�todo do driver
				    driver.manage().deleteAllCookies();
				}
			
            } catch (IOException e) {
				System.out.println("Ocorreu um erro no metodo getWebDriver: " + e.getMessage());
			}
            
            return driver;
    } 
    
    public static WebDriver getHandleToWindow(String title, WebDriver driver){

        WebDriver popup = null;
        Set<String> windowIterator = driver.getWindowHandles();
        System.err.println("No of windows :  " + windowIterator.size());
        for (String s : windowIterator) {
          String windowHandle = s; 
          popup = driver.switchTo().window(windowHandle);
          System.out.println("Window Title : " + popup.getTitle());
          System.out.println("Window Url : " + popup.getCurrentUrl());
          if (popup.getTitle().equals(title) ){
              System.out.println("Selected Window Title : " + popup.getTitle());
              return popup;
          }

        }
          System.out.println("Window Title :" + popup.getTitle());
          System.out.println();
          return popup;
        }
    
    
    public static void gravarArquivo(String caminhoDiretorio, String nomeArquivo, String extensaoArquivo, String conteudoArquivo, String mensagem) throws IOException {
    	
    	String arquivo = caminhoDiretorio + "/" + nomeArquivo + extensaoArquivo; 
    	File file = new File(arquivo);
		BufferedWriter writer = new BufferedWriter(new FileWriter(file));
		writer.write(mensagem + conteudoArquivo);
		writer.newLine();
		//Criando o conte�do do arquivo
		writer.flush();
		//Fechando conex�o e escrita do arquivo.
		writer.close();
		
    }

}