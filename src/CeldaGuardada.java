/**
 * Clase encargada de guardar las celdas cuando se modifica un valor
 * @author rauls
 *
 */
public class CeldaGuardada {
	int fila;
	int col;
	String valor;
	String valorFinal;	
	
	/**
	 * Constructor de la clase celda
	 * @param fila modificada
	 * @param col modificada
	 * @param valor valor que usara la cola undo
	 * @param valorFinal valor que usara la cola redo
	 */
	public CeldaGuardada(int fila, int col, String valor, String valorFinal) {
		this.fila=fila;
		this.col=col;
		this.valor=valor;
		this.valorFinal=valorFinal;
	}
	
	/**
	 * Devuelve la fila
	 * @return fila
	 */
	public int getFila() {
		return fila;
	}
	
	/**
	 * Devuelve la columna
	 * @return
	 */
	public int getCol() {
		return col;
	}
	
	/**
	 * Devuelve el valor primero
	 * @return
	 */
	public String getValor() {
		return valor;
	}
	
	/**
	 * Devuelve el valor tras la modificacion
	 * @return
	 */
	public String getValorFinal() {
		return valorFinal;
	}
}
