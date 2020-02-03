

/**
 * Clase que almacena las celdas de cada hoja excel.
 * 
 * @author rauls
 *
 */

public class Hoja {
	
	private int filas;
	private int cols;
	private Celda celdas[][];
	
	
	/**
	 * Constructor de la clase hoja
	 * @param fila Numero de filas que tendrá la hoja
	 * @param col Número de columnas que tendrá la hoja
	 */
	public Hoja(int fila, int col) {
		this.filas=fila;
		this.cols=col;
		celdas=new Celda[fila][col];
		creaCelda();
	}
	
	/**
	 * Crea las celdas que tendrá nuestra hoja excel.
	 */
	public void creaCelda() {
		String actual="";
		int contAux=26;
		int contAux2=702;
		for(int i=0; i<filas; i++) {
			for (int j=0; j<cols; j++) {
				if(j<26) {
					char colum=(char) (j+65);
					actual += colum;
					Celda celda=new Celda(i+1, actual);
					celdas[i][j]=celda;
				} else if (j>=26 && j<702) {
					char colum=(char) ((j%26)+65);
					char columAux=(char)((j/26)+64);
					actual+=columAux;
					actual+=colum;
					Celda celda=new Celda((i+1), actual);
					celdas[i][j]=celda;
				} else {
					if (j > 702 && (i) % 26 == 0) {
						contAux++;
					}
					if (j > 702 && (i - 26) % 676 == 0) {
						contAux2++;
					}
					char colum=(char) ((j%26)+65);
					char columAux=(char)(contAux+65);
					char columAux2=(char)(contAux2+65);
					actual+=columAux2;
					actual+=columAux;
					actual+=colum;
					Celda celda=new Celda((i+1), actual);
					celdas[i][j]=celda;
				}
				actual="";
			}
		}
	}
	
	/**
	 * Busca por una celda.
	 * @param fila de la que se busca la celda
	 * @param col de la que se busca la celda.
	 * @return
	 */
	public Celda buscaCeldaCorrecta(int fila, int col) {
		return celdas[fila][col];
	}
}
