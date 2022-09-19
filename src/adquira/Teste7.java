package adquira;

import java.io.BufferedReader;
import java.io.File;
import java.io.InputStreamReader;

public class Teste7 {

	public static void main(String[] args) {
		

	}
	
    public static void criaDiretorio(String caminhoDiretorio){
        File diretorio = new File(caminhoDiretorio);
        if (!diretorio.exists()) {
        	diretorio.mkdirs();
        }
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
    		
    		System.out.println("Deu erro na criação do diretório Temp: " + e.getMessage());
    	
    	}
    	
    }

}
