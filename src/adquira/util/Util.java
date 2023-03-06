package adquira.util;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Properties;

import adquira.Pedido;

public class Util {
	
    public static String getValor(String chave) throws IOException{
    	Properties props = getProp();
        return (String)props.getProperty(chave);
    }
    
    // Arquivo de properties sendo usado fora do projeto
    public static Properties getProp() throws IOException {
        Properties props = new Properties();
        
         String arquivoProperties = "D:/JOBS/AutomacaoAdquira/configuracoes/propriedadesAdquira.properties"; 
        // String arquivoProperties = "C:/Automacao Adquira/configuracoes/gerar relatorio pedidos.properties"; 
        // Quando a integra��o do Sharepoibt estiver pronta, usar o arquivo propriedadesAdquira com Sharepoint.properties
        // String arquivoProperties = "C:/Viegas/desenvolvimento/Selenium/Adquira/configuracoes/propriedadesAdquira.properties";
         
        FileInputStream file = new FileInputStream(arquivoProperties);
        props.load(file);
        return props;
    }
    
	public static void converteValorNullParaEspacoEmBranco(Pedido pedido) {
		
		if(pedido.getContractNumber().getContrato()==null || pedido.getContractNumber().getContrato().isEmpty()	){          pedido.getContractNumber().setContrato(" ");   				}
		if(pedido.getContractNumber().getFrente()==null	|| pedido.getContractNumber().getFrente().isEmpty() ){          pedido.getContractNumber().setFrente(" ");   					}
		if(pedido.getContractNumber().getNumero()==null	|| pedido.getContractNumber().getNumero().isEmpty() ){         	pedido.getContractNumber().setNumero(" ");    		}
		if(pedido.getContractNumber().getWbs()==null || pedido.getContractNumber().getWbs().isEmpty() ){			        pedido.getContractNumber().setWbs(" ");  }
		if(pedido.getNumero()==null || pedido.getNumero().isEmpty()	){          pedido.setNumero(" ");   					}
		if(pedido.getData()==null || pedido.getData().isEmpty() 	){          pedido.setData(" ");   			}
		if(pedido.getCnpjCliente()==null || pedido.getCnpjCliente().isEmpty() ){          pedido.setCnpjCliente(" ");   			}
		if(pedido.getComprador()==null	|| pedido.getComprador().isEmpty() 	){          pedido.setComprador(" ");	      			}
		if(pedido.getMensagemErroRegraPreenchimento()==null	|| pedido.getMensagemErroRegraPreenchimento().isEmpty()	){          pedido.setMensagemErroRegraPreenchimento(" ");	    	}
		if(pedido.getObservacaoSharepoint()==null || pedido.getObservacaoSharepoint().isEmpty()	){          pedido.setObservacaoSharepoint(" ");	      			}

	}

}