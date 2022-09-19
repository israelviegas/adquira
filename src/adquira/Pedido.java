package adquira;
import java.util.ArrayList;
import java.util.List;

public class Pedido {
	
	private ContractNumber contractNumber;
	
	private String data;
	
	private String numero;
	
	private String valor;
	
	private String campoNome;
	
	private String CampoValor;
	
	private String CnpjCentro;
	
	private String CnpjCliente;
	
	private String comprador;
	
	private String estado;
	
	private boolean salvoNoSharepoint;
	
	private boolean isFaturado;
	
	private String numeroContractNumber;
	
	private String cap;
	
	private Integer item;
	
	private String prazoPagamento;
	
	String	dataExtracao;

	List<Integer> listaItem = new ArrayList<Integer>();

	List<String> listaCap = new ArrayList<String>();
	
	private String mensagemErroRegraPreenchimento = "";
	
	private String observacaoSharepoint;
	
	boolean contractNumberConforme;
	
	boolean encontrouPdfAnexo;
	
	private String mensagemDeErroNoPedido = "";
	
	public ContractNumber getContractNumber() {
		return contractNumber;
	}

	public void setContractNumber(ContractNumber contractNumber) {
		this.contractNumber = contractNumber;
	}

	public String getData() {
		return data;
	}

	public void setData(String data) {
		this.data = data;
	}

	public String getNumero() {
		return numero;
	}

	public void setNumero(String numero) {
		this.numero = numero;
	}

	public String  getValor() {
		return valor;
	}

	public void setValor(String valor) {
		this.valor = valor;
	}

	public String getCampoNome() {
		return campoNome;
	}

	public void setCampoNome(String campoNome) {
		this.campoNome = campoNome;
	}

	public String getCampoValor() {
		return CampoValor;
	}

	public void setCampoValor(String campoValor) {
		CampoValor = campoValor;
	}

	public String getCnpjCentro() {
		return CnpjCentro;
	}

	public void setCnpjCentro(String cnpjCentro) {
		CnpjCentro = cnpjCentro;
	}

	public String getCnpjCliente() {
		return CnpjCliente;
	}

	public void setCnpjCliente(String cnpjCliente) {
		CnpjCliente = cnpjCliente;
	}

	public String getComprador() {
		return comprador;
	}
	
	public void setComprador(String comprador) {
		this.comprador = comprador;
	}

	public String getEstado() {
		return estado;
	}

	public void setEstado(String estado) {
		this.estado = estado;
	}

	public boolean isSalvoNoSharepoint() {
		return salvoNoSharepoint;
	}

	public void setSalvoNoSharepoint(boolean salvoNoSharepoint) {
		this.salvoNoSharepoint = salvoNoSharepoint;
	}
	

	public boolean isFaturado() {
		return isFaturado;
	}

	public void setFaturado(boolean isFaturado) {
		this.isFaturado = isFaturado;
	}
	
	public String getNumeroContractNumber() {
		return numeroContractNumber;
	}

	public void setNumeroContractNumber(String numeroContractNumber) {
		this.numeroContractNumber = numeroContractNumber;
	}

	public String getCap() {
		return cap;
	}

	public void setCap(String cap) {
		this.cap = cap;
	}

	public Integer getItem() {
		return item;
	}

	public void setItem(Integer item) {
		this.item = item;
	}

	public String getPrazoPagamento() {
		return prazoPagamento;
	}

	public void setPrazoPagamento(String prazoPagamento) {
		this.prazoPagamento = prazoPagamento;
	}

	public String getDataExtracao() {
		return dataExtracao;
	}

	public void setDataExtracao(String dataExtracao) {
		this.dataExtracao = dataExtracao;
	}

	public List<Integer> getListaItem() {
		return listaItem;
	}
	
	public void setListaItem(List<Integer> listaItem) {
		this.listaItem = listaItem;
	}

	public List<String> getListaCap() {
		return listaCap;
	}

	public void setListaCap(List<String> listaCap) {
		this.listaCap = listaCap;
	}

	public String getMensagemErroRegraPreenchimento() {
		return mensagemErroRegraPreenchimento;
	}

	public void setMensagemErroRegraPreenchimento(String mensagemErroRegraPreenchimento) {
		this.mensagemErroRegraPreenchimento = mensagemErroRegraPreenchimento;
	}

	public String getObservacaoSharepoint() {
		return observacaoSharepoint;
	}

	public void setObservacaoSharepoint(String observacaoSharepoint) {
		this.observacaoSharepoint = observacaoSharepoint;
	}

	public boolean isContractNumberConforme() {
		return contractNumberConforme;
	}

	public void setContractNumberConforme(boolean encontrouContractNumber) {
		this.contractNumberConforme = encontrouContractNumber;
	}

	public boolean isEncontrouPdfAnexo() {
		return encontrouPdfAnexo;
	}

	public void setEncontrouPdfAnexo(boolean encontrouPdfAnexo) {
		this.encontrouPdfAnexo = encontrouPdfAnexo;
	}

	public String getMensagemDeErroNoPedido() {
		return mensagemDeErroNoPedido;
	}

	public void setMensagemDeErroNoPedido(String mensagensDeErroNoPedido) {
		this.mensagemDeErroNoPedido = mensagensDeErroNoPedido;
	}

}
