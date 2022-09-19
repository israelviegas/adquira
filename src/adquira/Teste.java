package adquira;
import java.io.FileInputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Properties;
import java.util.TimeZone;
import  java.sql.Timestamp;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import  java.util.Date;   

import javax.swing.JFrame;
import javax.swing.JOptionPane;

import org.apache.poi.hpsf.Decimal;

public class Teste {

	public static void main(String[] args) throws Exception {
		
		String valorPedido = "BRL 10.354.954,69";
		//valorPedido = "BRL 10.954,69";
		
		String valorFormatado = formatarValorPedido(valorPedido);
		
		System.out.println(valorFormatado);
		
		//System.out.println(String.format("%.2f", new BigDecimal(valorFormatado)));
		
		Map<String, String> scripts = new HashMap<String, String>();
		scripts.put("scriptAs", "1");
		System.out.println(scripts.get("scriptAs"));
		
		String timesTamp = "1660092265210";
		
		long dataLong = Long.parseLong(timesTamp);
		
        Timestamp timestamp = new Timestamp(dataLong);  
        Date data = new Date(timestamp.getTime());  
        DateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss:SSS");
        df.setTimeZone(TimeZone.getTimeZone("GMT"));
        String dataLocal = df.format(data);
        
        System.out.println(dataLocal);    
		
		int teste = 0;
		
		teste = teste + 1;
		
		System.out.println(teste);
		
		boolean valor = false;
		
		System.out.println(valor);
		
		List<Pedido> listaPedidos = new ArrayList<Pedido>();
		List<Pedido> listaPedidosFaturados = new ArrayList<Pedido>();
		List<Pedido> listaPedidosNaoFaturados = new ArrayList<Pedido>();
		
		// Pedidos totais
		Pedido pedido1 = new Pedido();
		pedido1.setNumeroContractNumber("1");
		pedido1.setNumero("1");
		pedido1.setCampoNome("Telefonica Brasil S.A");
		pedido1.setCampoValor("02.558.157/0518-24");
		pedido1.setData("14-oct-2019");
		pedido1.setValor("123");
		listaPedidos.add(pedido1);
		
		Pedido pedido2 = new Pedido();
		pedido2.setNumeroContractNumber("1");
		pedido2.setNumero("2");
		pedido2.setCampoNome("CNPJ Centro");
		pedido2.setCampoValor("02.558.157/0518-25");
		pedido2.setData("14-oct-2019");
		pedido2.setValor("123");
		listaPedidos.add(pedido2);
		
		Pedido pedido3 = new Pedido();
		pedido3.setNumeroContractNumber("1");
		pedido3.setNumero("3");
		pedido3.setCampoNome("Telefonica Brasil S.A");
		pedido3.setCampoValor("02.558.157/0518-24");
		pedido3.setData("14-oct-2019");
		pedido3.setValor("123");
		listaPedidos.add(pedido3);
		
		listaPedidosNaoFaturados.addAll(listaPedidos);
		
		
		pedido3.setNumeroContractNumber("3");
		
		// Pedidos faturados
		Pedido pedido4 = new Pedido();
		pedido4.setNumeroContractNumber("1");
		pedido4.setNumero("1");
		pedido4.setCampoNome("CNPJ Centro");
		pedido4.setCampoValor("02.558.157/0518-25");
		pedido4.setData("14-oct-2019");
		pedido4.setValor("123");
		listaPedidosFaturados.add(pedido4);
		
		Pedido pedido5 = new Pedido();
		pedido5.setNumeroContractNumber("2");
		pedido5.setNumero("2");
		pedido5.setCampoNome("CNPJ Centro");
		pedido5.setCampoValor("02.558.157/0518-25");
		pedido5.setData("14-oct-2019");
		pedido5.setValor("123");
		listaPedidosFaturados.add(pedido5);
		
		
/*		for (int i = 0; i < listaPedidos.size(); i++) {
			
			for (int j = 0; j < listaPedidosFaturados.size(); j++) {
				

				
			}

			
		}
		*/
		
		
		
		  // Algo que deseja mostrar (aviso, mensagem de erro)
/*	    String erro = "Erro 404: não foi possível encontrar o batman";

	    // Cria um JFrame
	    JFrame frame = new JFrame("JOptionPane exemplo");

	    // Cria o JOptionPane por showMessageDialog
	    JOptionPane.showMessageDialog(frame,
	        "Houve um problema ao procurar o batman:\n\n '" + erro + "'.", //mensagem
	        "Erro 404", // titulo da janela 
	        JOptionPane.INFORMATION_MESSAGE);
	    System.exit(0);*/
		
		
	    String erro = "";

/*	    // Cria um JFrame
	    JFrame frame = new JFrame("JOptionPane exemplo");

	    // Cria o JOptionPane por showMessageDialog
	    JOptionPane.showMessageDialog(frame,
	        "Houve um problema na extração dos pedidos do Adquira\n" + erro + "", //mensagem
	        "Automatização Adquira", // titulo da janela 
	        JOptionPane.INFORMATION_MESSAGE);
	    System.exit(0);
	    
	    
	    String caminho = "C:\\Automacao Adquira\\executaveis\\gerar relatorio pedidos.bat";
	    
	    
	    Object[] options = { "Sim", "Não" };
	    int i = JOptionPane.showOptionDialog(null, "Extração dos pedidos do Adquira executada com sucesso! \n Gostaria de executar a caralha do outro programa?", "Saída", JOptionPane.YES_NO_OPTION, JOptionPane.QUESTION_MESSAGE, null, options, options[0]);
	   
	    if (i == JOptionPane.YES_OPTION) {
	    	
	    //	Process process = Runtime.getRuntime().exec(caminho);

	    } else {
	    	
	    	System.exit(0); 
	    }*/
	    
		
		for (Pedido pedido : listaPedidos) {

			for (Pedido pedidoFaturado : listaPedidosFaturados) {

				if (pedido.getNumero().equals(pedidoFaturado.getNumero())) {
					
					listaPedidosNaoFaturados.remove(pedido);
					
				}
	        	
	        }
        	
        }
		
		
		for (Pedido pedido : listaPedidos) {

			System.out.println("Número pedido: " + pedido.getNumero());
        	
        }
		
		
		for (Pedido pedido : listaPedidosNaoFaturados) {

			System.out.println("Número pedido não faturado: " + pedido.getNumero());
        	
        }

		
		// Teste de abertura de excel
		// Executa o túnel para poder acessar o Power Bi
		Runtime.getRuntime().exec(getValor("caminho.contract.numbers2"));
		
		
	}

	public static String getValor(String chave) throws IOException{
		Properties props = getProp();
		return (String)props.getProperty(chave);
	}
	
    // Arquivo de properties sendo usado fora do projeto
    public static Properties getProp() throws IOException {
        Properties props = new Properties();
       
        
        // String arquivoProperties = "C:/Automacao Adquira/configuracoes/gerar relatorio pedidos.properties"; 
        // Quando a integração do Sharepoibt estiver pronta, usar o arquivo propriedadesAdquira com Sharepoint.properties
         String arquivoProperties = "C:/Viegas/desenvolvimento/Selenium/arquivos propriedades/propriedadesAdquira.properties"; 
        
        FileInputStream file = new FileInputStream(arquivoProperties);
        props.load(file);
        return props;
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
    

}

