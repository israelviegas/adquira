package adquira;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

public class Teste3 {
	
	private static List<ContractNumber> listaContractNumbers = new ArrayList<ContractNumber>();
	private static List<String> listaRepetidos = new ArrayList<String>();
	private static List<String> listaDistintosFinal = new ArrayList<String>();
	
	public static void main(String[] args) {
		
		Collection lista = new ArrayList();
		lista.add("1");
		lista.add("2");
		lista.add("4");
		lista.add("5");
		lista.add("1");
		lista.add("1");
		
		lista = Collections.singleton(new HashSet(lista));
		
		for (Object object : lista) {
			System.out.println(object.toString());
		}
	
	
		
		listaRepetidos.add("1");
		listaRepetidos.add("2");
		listaRepetidos.add("4");
		listaRepetidos.add("5");
		listaRepetidos.add("1");
		listaRepetidos.add("1");
		listaRepetidos.add("2");
		listaRepetidos.add("4");
		listaRepetidos.add("5");
		listaRepetidos.add("3");
		listaRepetidos.add("3");

		for (String elementosRepetidos : listaRepetidos) {
			
			System.out.println("Elemento repetidos " + elementosRepetidos);
			
		}

		System.out.println(" ");
	
		Set<String> novaSet = new HashSet<String>();
		novaSet.add("1");
		novaSet.add("2");
		novaSet.add("3");
		novaSet.add("4");
		novaSet.add("5");
		novaSet.add("1");
		novaSet.add("2");
		novaSet.add("3");
		novaSet.add("4");
		novaSet.add("5");
		
		for (String elementosDistintos : novaSet) {
			
			System.out.println("Elemento distintos " + elementosDistintos);
			
			for (String elementosRepetidos : listaRepetidos) {
				
				if (elementosDistintos.equals(elementosRepetidos)) {
					
					listaDistintosFinal.add(elementosDistintos);
					break;
					
				}
				
				
			}
			
		}
		
		System.out.println(" ");
		
		for (String elementosDistintosfinal : listaDistintosFinal) {
			
			System.out.println("Elemento distintos final " + elementosDistintosfinal);
			
		}
	
		
		
	
	}
}

