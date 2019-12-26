
public class CeldaGuardada {
	int fila;
	int col;
	String valor;
	String valorFinal;	
	public CeldaGuardada(int fila, int col, String valor, String valorFinal) {
		this.fila=fila;
		this.col=col;
		this.valor=valor;
		this.valorFinal=valorFinal;
	}
	public int getFila() {
		return fila;
	}
	public int getCol() {
		return col;
	}
	public String getValor() {
		return valor;
	}
	public String getValorFinal() {
		return valorFinal;
	}
}
