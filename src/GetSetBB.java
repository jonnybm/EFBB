import java.util.ArrayList;

public class GetSetBB {
	
//pegar a leitura do PDF e guardar momentanamente
		
	public String Autor = "";	
	public String Reu = "";
	public String Processo = "";
	public String Vara = "";
	public String Comarca = "";
	public String Estado = "";
	public String NaturezaAcao = "";
	public String DataDeposito = "";
	public String ValorOriginal = "";
	public String ValorAtualizado = "";
	public String ContaJuridica = "";
	public String Parcela = "";
	
	

	public String getAutor() {
		return Autor;
	}
	public void setAutor(String autor) {
		Autor = autor;
	}
	
	public String getReu() {
		return Reu;
	}
	public void setReu(String reu) {
		Reu = reu;
	}
	
	public String getProcesso() {
		return Processo;
	}
	public void setProcesso(String processo) {
		Processo = processo;
	}
	
	public String getVara() {
		return Vara;
	}
	public void setVara(String vara) {
		Vara = vara;
	}
	
	public String getComarca() {
		return Comarca;
	}
	public void setComarca(String comarca) {
		Comarca = comarca;
	}
	
	public String getEstado() {
		return Estado;
	}
	public void setEstado(String estado) {
		Estado = estado;
	}
	
	public String getNaturezaAcao() {
		return NaturezaAcao;
	}
	public void setNaturezaAcao(String naturezaAcao) {
		NaturezaAcao = naturezaAcao;
	}
	
	public String getDataDeposito() {
		return DataDeposito;
	}
	public void setDataDeposito(String dataDeposito) {
		DataDeposito = dataDeposito;
	}
	
	public String getValorOriginal() {
		return ValorOriginal;
	}
	public void setValorOriginal(String valorOriginal) {
		ValorOriginal = valorOriginal;
	}
	
	public String getValorAtualizado() {
		return ValorAtualizado;
	}
	public void setValorAtualizado(String valorAtualizado) {
		ValorAtualizado = valorAtualizado;
	}
	
	public String getContaJuridica() {
		return ContaJuridica;
	}
	public void setContaJuridica(String contaJuridica) {
		ContaJuridica = contaJuridica;
	}
	
	public String getParcela() {
		return Parcela;
	}
	public void setParcela(String parcela) {
		Parcela = parcela;
	}

	
	
	
//////////////////////////////////////////////////////////////////////////
	
	
	// get Set para incluir existente unico //CONTA EXISTE  É UNICA
	
	public String Autor_Unico = "";	
	public String Reu_Unico = "";
	public String Processo_Unico = "";
	public String Vara_Unico = "";
	public String Comarca_Unico = "";
	public String Estado_Unico = "";
	public String NaturezaAcao_Unico = "";
	public String DataDeposito_Unico = "";
	public String ValorOriginal_Unico = "";
	public String ValorAtualizado_Unico = "";
	public String ContaJuridica_Unico = "";
	public String Parcela_Unico = "";
	
	
	public String getAutor_Unico() {
		return Autor_Unico;
	}
	public void setAutor_Unico(String autor_Unico) {
		Autor_Unico = autor_Unico;
	}
	public String getReu_Unico() {
		return Reu_Unico;
	}
	public void setReu_Unico(String reu_Unico) {
		Reu_Unico = reu_Unico;
	}
	public String getProcesso_Unico() {
		return Processo_Unico;
	}
	public void setProcesso_Unico(String processo_Unico) {
		Processo_Unico = processo_Unico;
	}
	public String getVara_Unico() {
		return Vara_Unico;
	}
	public void setVara_Unico(String vara_Unico) {
		Vara_Unico = vara_Unico;
	}
	public String getComarca_Unico() {
		return Comarca_Unico;
	}
	public void setComarca_Unico(String comarca_Unico) {
		Comarca_Unico = comarca_Unico;
	}
	public String getEstado_Unico() {
		return Estado_Unico;
	}
	public void setEstado_Unico(String estado_Unico) {
		Estado_Unico = estado_Unico;
	}
	public String getNaturezaAcao_Unico() {
		return NaturezaAcao_Unico;
	}
	public void setNaturezaAcao_Unico(String naturezaAcao_Unico) {
		NaturezaAcao_Unico = naturezaAcao_Unico;
	}
	public String getDataDeposito_Unico() {
		return DataDeposito_Unico;
	}
	public void setDataDeposito_Unico(String dataDeposito_Unico) {
		DataDeposito_Unico = dataDeposito_Unico;
	}
	public String getValorOriginal_Unico() {
		return ValorOriginal_Unico;
	}
	public void setValorOriginal_Unico(String valorOriginal_Unico) {
		ValorOriginal_Unico = valorOriginal_Unico;
	}
	public String getValorAtualizado_Unico() {
		return ValorAtualizado_Unico;
	}
	public void setValorAtualizado_Unico(String valorAtualizado_Unico) {
		ValorAtualizado_Unico = valorAtualizado_Unico;
	}
	public String getContaJuridica_Unico() {
		return ContaJuridica_Unico;
	}
	public void setContaJuridica_Unico(String contaJuridica_Unico) {
		ContaJuridica_Unico = contaJuridica_Unico;
	}
	public String getParcela_Unico() {
		return Parcela_Unico;
	}
	public void setParcela_Unico(String parcela_Unico) {
		Parcela_Unico = parcela_Unico;
	}



//////////////////////////////////////////////////////////////////////
	
	//CONTA EXISTE E NAO É UNICA
	public ArrayList<String> arrayCJExisteNaoUnica = new ArrayList<String> ();

	public ArrayList<String> getArrayCJExisteNaoUnica() {
		return arrayCJExisteNaoUnica;
	}
	public void setArrayCJExisteNaoUnica(ArrayList<String> arrayCJExisteNaoUnica) {
		this.arrayCJExisteNaoUnica = arrayCJExisteNaoUnica;
	}
	
	
	
//////////////////////////////////////////////////////////////////////
	


	public ArrayList<String> arrayCJExisteNaoUnicaIncluir = new ArrayList<String> ();



	public ArrayList<String> getArrayCJExisteNaoUnicaIncluir() {
		return arrayCJExisteNaoUnicaIncluir;
	}
	public void setArrayCJExisteNaoUnicaIncluir(ArrayList<String> arrayCJExisteNaoUnicaIncluir) {
		this.arrayCJExisteNaoUnicaIncluir = arrayCJExisteNaoUnicaIncluir;
	}




	//CONTA NOVA COM VALOR
	public ArrayList<String> arrayCJNovaComValor = new ArrayList<String> ();
	
	public ArrayList<String> getArrayCJNovaComValor() {
		return arrayCJNovaComValor;
	}
	public void setArrayCJNovaComValor(ArrayList<String> arrayCJNovaComValor) {
		this.arrayCJNovaComValor = arrayCJNovaComValor;
	}
	

	
	//CONTA NOVA SEM VALOR	
	public ArrayList<String> arrayCJNovaSemValor = new ArrayList<String> ();

	public ArrayList<String> getArrayCJNovaSemValor() {
		return arrayCJNovaSemValor;
	}
	public void setArrayCJNovaSemValor(ArrayList<String> arrayCJNovaSemValor) {
		this.arrayCJNovaSemValor = arrayCJNovaSemValor;
	}
	
	
	
	boolean cjExiste = false;
	boolean cjExisteUnica = true;
	boolean cjNova = true;



	public boolean isCjExiste() {
		return cjExiste;
	}
	public void setCjExiste(boolean cjExiste) {
		this.cjExiste = cjExiste;
	}
	public boolean isCjExisteUnica() {
		return cjExisteUnica;
	}
	public void setCjExisteUnica(boolean cjExisteUnica) {
		this.cjExisteUnica = cjExisteUnica;
	}
	public boolean isCjNova() {
		return cjNova;
	}
	public void setCjNova(boolean cjNova) {
		this.cjNova = cjNova;
	}
	
	
	int posicaoExiste = 0;

	public int getPosicaoExiste() {
		return posicaoExiste;
	}
	public void setPosicaoExiste(int posicaoExiste) {
		this.posicaoExiste = posicaoExiste;
	} 


	
	int contadorPosicao = 11;

	public int getContadorPosicao() {
		return contadorPosicao;
	}
	public void setContadorPosicao(int contadorPosicao) {
		this.contadorPosicao = contadorPosicao;
	}
	
}
