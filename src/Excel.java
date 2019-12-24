

import java.util.*;

/**
 * Clase que almacena todos las hojas de nuestro programa Excel
 * 
 * @author Raúl Sevillano González
 *
 */

public class Excel {

	private int filas;
	private int cols;
	private Hoja hoja;

	/**
	 * Comprueba que el numero de datos que se han introducido coincide con el
	 * numero de columnas de la hoja
	 * 
	 * @param leeDatos
	 * @param hoja
	 * @throws ExcelException
	 */
	public Excel(int filas, int cols) {
		this.filas=filas;
		this.cols=cols;
		hoja=new Hoja(filas, cols);
	}

	/**
	 * Metodo que comprueba si es un dato o una formula el dato insertado
	 * 
	 * @param valorCeldas
	 */
	public void esNumerico(String valorCeldas[][]) {
		for (int i = 0; i < filas; i++) {
			for (int j = 0; j < cols; j++) {
				if (valorCeldas[i][j].charAt(0) != '=') {
					hoja.buscaCeldaCorrecta(i, j).setValor(valorCeldas[i][j]);
				}
			}
		}
		esUnaFormula(valorCeldas);
	}

	/**
	 * Metodo que analiza la formula y aniade el valor a la celda correspondiente
	 * 
	 * @param valorCelda valor que tendra la celda.
	 */
	public void esUnaFormula(String valorCeldas[][]) {
		int valor = 0;
		String filaActual = "";
		String columnaActual = "";
		String valorS = "";
		for (int j = 0; j < filas; j++) {
			for (int k = 0; k < cols; k++) {
				if (valorCeldas[j][k].charAt(0) == '=') {
					for (int i = 0; i < valorCeldas[j][k].trim().length(); i++) {
						switch (tipoCaracter(valorCeldas[j][k].trim().charAt(i))) {
						case 0:
							filaActual += (valorCeldas[j][k].charAt(i));
							break;
						case 1:
							columnaActual += valorCeldas[j][k].charAt(i);
							break;
						case 2:
							break;
						case 3:
							valor = esSuma(valor,filaActual, columnaActual);
							filaActual = "";
							columnaActual = "";
							break;
						case 4:
							break;
						default:
							System.out.println("Formato no valido");
							System.exit(0);
						}
					}
					if (filaActual != "" && columnaActual != "") {
						valor = esSuma(valor, filaActual, columnaActual);
						filaActual = "";
						columnaActual = "";
					}
					valorS = String.valueOf(valor);
					hoja.buscaCeldaCorrecta(j, k).setValor(valorS+" ");
					valor = 0;
				}
			}
		}
	}
	
	public String[][] dameHoja(){
		String[][] datos=new String[filas][cols];
		for(int i=0;i<filas;i++) {
			for(int j=0;j<cols;j++) {
				datos[i][j]=hoja.buscaCeldaCorrecta(i, j).getValor();
			}
		}
		return datos;		
	}

	/**
	 * Busca el tipo de caracter para luego vincular con la columna
	 * 
	 * @param caracter
	 * @return Un número en función del tipo de caracter.
	 */
	public static int tipoCaracter(char caracter) {
		if (Character.isDigit(caracter)) {
			return 0;
		} else if (Character.isAlphabetic(caracter)) {
			return 1;
		} else if (caracter == ' ') {
			return 2;
		} else if (caracter == '+') {
			return 3;
		} else if (caracter == '=') {
			return 4;
		} else {
			return 5;
		}
	}

	/**
	 * Realiza la suma en caso de ser una formula.
	 * 
	 * @param valor
	 * @param fila
	 * @param columna
	 * @param hoja
	 * @return
	 */
	public int esSuma(int valor, String fila, String columna) {
		String valorSumar;
		int colAux = 0;
		if (columna.length() == 1) {
			colAux = Character.getNumericValue(columna.charAt(0)) - 10;
			valorSumar = hoja.buscaCeldaCorrecta(Integer.valueOf(fila) - 1, colAux).getValor();
			valor += Integer.parseInt(valorSumar.trim());
		} else if (columna.length() == 2) {
			colAux += Character.getNumericValue(columna.charAt(0)) - 10 + 26;
			colAux += Character.getNumericValue(columna.charAt(1)) - 10;
			valorSumar = hoja.buscaCeldaCorrecta(Integer.valueOf(fila) - 1, colAux).getValor();
			valor += Integer.parseInt(valorSumar.trim());
		} else if (columna.length() == 3) {
			colAux += Character.getNumericValue(columna.charAt(1)) - 10 + 702;
			colAux += Character.getNumericValue(columna.charAt(1)) - 10 + 26;
			colAux += Character.getNumericValue(columna.charAt(2)) - 10;
			valorSumar = hoja.buscaCeldaCorrecta(Integer.valueOf(fila) - 1, colAux).getValor();
			valor += Integer.parseInt(valorSumar.trim());
		}
		return valor;
	}
}
