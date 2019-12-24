

/**
 * Clase usada para crear celdas.
 * 
 * @author Raúl Sevillano González
 *
 */

public class Celda {
	
	private int fila;
	private String col;
	private String valor;
	
	/**
	 * Constructor de la clase celda
	 * @param fila en la que se encuentra la celda.
	 * @param col en la que se encuentra la celda.
	 */
	public Celda(int fila, String col) {
		this.fila=fila;
		this.col=col;
		this.valor="";
	}
	
	/**
	 * Setea el valor que tedrá la celda.
	 * @param valor
	 */
	public void setValor(String valor) {
		this.valor=valor;
	}
	
	/**
	 * Devuelve el valor que tiene la celda.
	 * @return
	 */
	public String getValor() {
		return this.valor;
	}
}
