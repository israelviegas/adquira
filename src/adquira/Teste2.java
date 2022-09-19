package adquira;
import java.io.IOException;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.List;
import java.util.Locale;

import javax.swing.JFrame;
import javax.swing.JOptionPane;

public class Teste2 {

	public static void main(String[] args) throws Exception {

		
		String data = "15-oct-2020";
		
		String dataFinal = formatarDataPedido(data);
		
		
		System.out.println("Data Final formatada: " + dataFinal);
		
		System.out.println("Data Nota Final formatada: " + formataDataDaNota(new Date()));
		
    	String cap = "Presta��o de Servi�o - CAP47328";
    	
		System.out.println("formataCap: " + formataCap(cap));
		
		if ("Terra Networks Brasil S.A.".equalsIgnoreCase("TERRA NETWORKS BRASIL S.A.")) {
			System.out.println("Iguais");
		} else {
			System.out.println("Diferentes");
		}
		
		
		/*
		 * GregorianCalendar calendar = new GregorianCalendar();
		 * calendar.setTime(minhaData); //aqui voc� usa sua vari�vel que chamei de
		 * "minhaData" int dia = calendar.get(GregorianCalendar.DAY_OF_MONTH); int mes =
		 * calendar.get(GregorianCalendar.MONTH);
		 */
		

		for (int i = 0 ; i < 100 ; i++) {
			
			//Thread.sleep(1000);
			
		//	imprime(i);
			
		}
		
		String teste1 = "Primeira frase, ";
		String teste2 = "Segunda frase, ";
		
		String fraseFinal = teste1 + teste2;
		
		if (fraseFinal.lastIndexOf(", ") != -1) {

			fraseFinal = fraseFinal.substring(0, fraseFinal.lastIndexOf(", "));
			
		}
		
		String numeroFormatado = String.format("%04d", 8);
		
		System.out.println(numeroFormatado);
		
		
		 String.format("|%20d|", 93);
		
		System.out.println(fraseFinal);
		
		
		
		
		
}
	
	
	public static void  imprime (int i) throws InterruptedException {
		
		System.out.println("Passei aqui");
		
	}
	
	
	public static void  testeVariavel (int i) throws InterruptedException {
		
		if (i == 50) {
			
			
		}
		
	}
	
	
    public static String formataDataDaNota(Date dataInclusaoNota) throws Exception {
    	
    	//A regra para a data da nota no sharepoint � esta:
    	//	1) Adiciona-se 75 dias � data de inclus�o da nota.
    	//	2) A data da nota no sharepoint ser� de acordo com os seguintes dias de vencimento: 4, 12 e 22.
    	//	Ex: Nota inclu�da dia 15/10/20 + 75 dias = 29/12/2020. Neste cen�rio a data de vencimento da nota no sharepoint ser� 04/01/2021.  
    	
		DateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");
		String teste = "15/10/2020";
		
		// Transformo em date
		dataInclusaoNota = formatter.parse(teste);
    	
    	Date date = null;
    	String dataDaNota = null;
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
		// 21/06/2021
		
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
		
		System.out.println("Data com os dias corretos: " + cal2.getTime());
    	
		return dataDaNota; 
    	
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

	
}
