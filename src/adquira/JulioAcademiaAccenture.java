package adquira;

import java.util.Scanner;


/*
 * Capturar e exibir mensagem dos n�meros que s�o entre 0 e 10
 */

public class JulioAcademiaAccenture {

	public static void main(String[] args) {

		Scanner input = new Scanner(System.in);
		int numero;
		String resposta = "S";
		
		
		while (resposta.equals("S")) {
			System.out.println("Digite um n�mero:");
			numero = input.nextInt();
			
			if (numero >= 0 && numero <= 10) {
				System.out.println("N�mero " + numero + " " + "est� entre 0 e 10");
			} else {
				System.out.println("N�mero " + numero + " " + "n�o est� entre 0 e 10");
			}
			
			System.out.println("Deseja continuar? (Digite S para Sim ou qualquer outra coisa caso deseje sair)" );
			resposta = input.next().toString();
		}

		if (!resposta.equals("S")) {
			System.out.println("Obrigado por usar meu programa! :) ");
		}

	}

}
