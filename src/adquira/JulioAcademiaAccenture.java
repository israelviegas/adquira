package adquira;

import java.util.Scanner;


/*
 * Capturar e exibir mensagem dos números que são entre 0 e 10
 */

public class JulioAcademiaAccenture {

	public static void main(String[] args) {

		Scanner input = new Scanner(System.in);
		int numero;
		String resposta = "S";
		
		
		while (resposta.equals("S")) {
			System.out.println("Digite um número:");
			numero = input.nextInt();
			
			if (numero >= 0 && numero <= 10) {
				System.out.println("Número " + numero + " " + "está entre 0 e 10");
			} else {
				System.out.println("Número " + numero + " " + "não está entre 0 e 10");
			}
			
			System.out.println("Deseja continuar? (Digite S para Sim ou qualquer outra coisa caso deseje sair)" );
			resposta = input.next().toString();
		}

		if (!resposta.equals("S")) {
			System.out.println("Obrigado por usar meu programa! :) ");
		}

	}

}
