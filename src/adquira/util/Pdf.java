package adquira.util;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.pdfbox.text.PDFTextStripperByArea;

import java.io.File;

// File file = new File("C:\\Viegas\\desenvolvimento\\Selenium\\Adquira\\5101148472.pdf");

public class Pdf {
	
	 public static void main(String[] args) throws Exception {

		 	String pedido = "C:\\Viegas\\desenvolvimento\\Selenium\\Adquira\\relatorios\\2022_05_03 18_33_33\\pdfs baixados 2022_05_03 18_33_33\\5101160445.pdf";
		 	
		 	System.out.println(retornaNumeroContractNumber(pedido));
		 
	    }
	 
	 
	 public static String retornaConteudoPedido (String pedido) throws Exception {
		 
		 File file = new File(pedido);
		 
		 String pdfFileInText = "";

		 try {
			 
			 PDDocument document = PDDocument.load( file );
			 
			 if (!document.isEncrypted()) {
				 
				 PDFTextStripperByArea stripper = new PDFTextStripperByArea();
				 stripper.setSortByPosition(true);
				 
				 PDFTextStripper tStripper = new PDFTextStripper();
				 
				 pdfFileInText = tStripper.getText(document);
				 
			 }
			
		 } catch (Exception e) {
				throw new Exception("Ocorreu um erro no m�todo retornaConteudoPedido: " + e);
		}
		 
		 return pdfFileInText;
		 
	 }
	 
	 
	 public static String retornaNumeroContractNumber (String pedido) throws Exception {
		 
		 try {

			 String conteudoPedido = retornaConteudoPedido(pedido);
			 
			 String fraseContracNumber = "CONTRACT NUMBER";
			 
			 String somenteNumerosContractNumber = "";
			 
			 if (conteudoPedido != null && !conteudoPedido.isEmpty()) {
				 
				 if (conteudoPedido.lastIndexOf(fraseContracNumber) != -1) {
					 
					 int posicaoFraseContractNumber = conteudoPedido.lastIndexOf(fraseContracNumber);
					 
					 int tamanhoNumeroContractNumber = 11;
					 
					 int folga = 1;
					 
					 String contractNumber = conteudoPedido.substring(posicaoFraseContractNumber, posicaoFraseContractNumber + fraseContracNumber.length() + tamanhoNumeroContractNumber + folga);
					 
					 if (contractNumber != null && !contractNumber.isEmpty()) {
						 
						 somenteNumerosContractNumber = contractNumber.replaceAll("[^0-9]", "").trim();
						 
					 }
					 
				 }
				 
			 }
			 
			 return somenteNumerosContractNumber;
			
		} catch (Exception e) {
			throw new Exception("Ocorreu um erro no metodo retornaNumeroContractNumber: " + e);
		}
		 
		 
	 }	 

}
