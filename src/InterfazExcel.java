import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Component;
import java.awt.Dimension;
import java.awt.Font;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.util.LinkedList;
import java.util.Queue;
import java.util.Stack;
import java.util.StringTokenizer;

import javax.swing.BorderFactory;
import javax.swing.DefaultCellEditor;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JMenu;
import javax.swing.JMenuBar;
import javax.swing.JMenuItem;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.JTextField;
import javax.swing.SwingConstants;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableCellEditor;

@SuppressWarnings("serial")
public class InterfazExcel extends JFrame {
	
	//Atributo que almacen la tabla
	JTable tabla;
	
	//Tamaño de la tabla
	int fila;
	int col;
	
	//Celda seleccionada
	int filaSeleccionada = 1;
	int colSeleccionada = 1;

	//Pegatina de la celda seleccionada
	JLabel celda;
	
	//Modelo que usaremos para el diseño de nuestra tabla
	DefaultTableModel tab;
	
	//Colas para la función de atras y adelante
	Queue<CeldaGuardada> colaUndo;
	Queue<CeldaGuardada> colaRedo;
	
	//Dimension 
	Dimension screenSize;
	
	//Frame principal de ejecución
	JFrame ventanaPrincipal;
	
	//Variable que comprueba si el documento esta sin guardar
	boolean guardado;
	
	//Valor escrito en la celda
	String valor;
	
	JMenuItem hacer;
	JMenuItem deshacer;
	
	//Ultima celda sobre la que se ha trabajado en la hoja de calculo
	CeldaGuardada ultimaCelda;
	
	/**
	 * Constructor de la clase InterfazExcel 
	 * Creamos toda la interfaz de nuestro Excel
	 */
	public InterfazExcel() {
		screenSize = Toolkit.getDefaultToolkit().getScreenSize();
		
		//Pedimos las filas que tendrá la hoja
		pedirFilas();
		
		//Inicializamos ciertos parametros de nuestra hoja
		colaUndo = new LinkedList<CeldaGuardada>();
		colaRedo = new LinkedList<CeldaGuardada>();
		JMenuBar menuBarra = new JMenuBar();
		JMenu menu = new JMenu("Archivo");
		JMenuItem cargar = new JMenuItem("Cargar");
		JMenuItem archivar = new JMenuItem("Archivar");
		JMenuItem nueva = new JMenuItem("Nuevo");
		JMenu editar = new JMenu("Editar");
		hacer = new JMenuItem("Adelante");
		deshacer = new JMenuItem("Atras");
		
		//Añadimos el action listener al menu
		archivar.addActionListener(new ArchivarArchivo());	
		cargar.addActionListener(new AbrirArchivo());
		nueva.addActionListener(new NuevaHoja());
		
		//Añadimos al menu los botones
		menu.add(archivar);
		menu.add(cargar);
		menu.add(nueva);
		
		
		//Inicializamos el modelo donde la fila 0 y columna 0 serán no editables.
		tab = new DefaultTableModel(fila + 1, col + 1) {
			@Override
			public boolean isCellEditable(int row, int column) {
				if (row == 0 || column == 0) {
					return false;
				} else {
					return true;
				}
			}
		};
		
		//Inicializamos el valor de guardado
		guardado = false;
		
		//Inicializamos panel de los botones
		JPanel botones = new JPanel();
		JButton calcular = new JButton("Calcular");
		JButton limpiar = new JButton("Limpiar");
		celda = new JLabel("A1");
		
		//Añadimos los action listener a los botones
		calcular.addActionListener(new CalculaFormulas());
		limpiar.addActionListener(new LimpiaCeldas());
		
		//Añadimos los botones al frame
		botones.add(calcular);
		botones.add(limpiar);
		botones.add(celda);
		
		//Seteamos los items a que no se puedan utilizar ya que no hay ninguna accion realizada. Y añadimos los listener
		deshacer.setEnabled(false);
		deshacer.addActionListener(new Undo());
		hacer.setEnabled(false);
		hacer.addActionListener(new Redo());
		
		//Añadimos los botones
		editar.add(hacer);
		editar.add(deshacer);
		
		//Inicializamos el frame panel y tabla.
		ventanaPrincipal = new JFrame();
		tabla = new JTable(tab); //Modelo que vamos a usar en la tabla.
		JPanel panel = new JPanel();
		
		//Seteamos atributos del frame.
		ventanaPrincipal.setDefaultCloseOperation(JFrame.DO_NOTHING_ON_CLOSE);
		ventanaPrincipal.setBounds(0, 0, (int) screenSize.getWidth(), (int) screenSize.getHeight());
		ventanaPrincipal.setVisible(true);
		
		//Seteamos altura de la fila
		tabla.setRowHeight(25);
		
		//Que hara la tabla cuando se haga un cambio en ella
		DefaultTableCellRenderer render = new DefaultTableCellRenderer() {
			public Component getTableCellRendererComponent(JTable jTable1, Object value, boolean selected,
					boolean focused, int row, int column) {
				setHorizontalAlignment(SwingConstants.CENTER);
				super.getTableCellRendererComponent(jTable1, value, selected, focused, row, column);
				
				//Setea el fondo a gris si es 0,0 (Esquina superior)
				if (row == 0 && column == 0) {
					this.setBackground(Color.gray);
				//Setea el fondo gris si es un identificador
				} else if (row == 0 || column == 0) {
					this.setBackground(Color.GREEN);
				//Setea a blanco si es una celda normal.
				} else {
					this.setBackground(Color.WHITE);
				}
				//En caso de que sea la celda que esta seleccionada se setea el background a amarillo
				if (row == filaSeleccionada && column == colSeleccionada && filaSeleccionada != 0
						&& colSeleccionada != 0) {
				} else if (column == 0 && row == filaSeleccionada && filaSeleccionada > 0 && colSeleccionada > 0) {
					this.setBackground(Color.YELLOW);
				} else if (row == 0 && column == colSeleccionada && filaSeleccionada > 0 && colSeleccionada > 0) {
					this.setBackground(Color.yellow);
				}
				//Se setea el borde mas ancho en caso de que sea la seleccionada
				if (row == filaSeleccionada && column == colSeleccionada) {
					setBorder(BorderFactory.createLineBorder(Color.BLACK, 3));
				} else {
					setBorder(BorderFactory.createLineBorder(Color.BLACK, 1));
				}
				return this;
			}
		};
		
		valor = "";
		
		//Añadimos al barra los dos menus
		menuBarra.add(menu);
		menuBarra.add(editar);
		
		//Seteamos la barra de menus
		ventanaPrincipal.setJMenuBar(menuBarra);
		
		//Seteamos el render de la tabla
		tabla.setDefaultRenderer(Object.class, render);
		//Seteamos la fuente de la letra
		tabla.setFont(new Font("Arial", Font.BOLD, 15));
		
		//Añadimos los botones al frame al sur
		ventanaPrincipal.add(botones, BorderLayout.SOUTH);
		
		//Seteamos el ancho de las columnas
		for (int i = 0; i < col + 1; i++) {
			tabla.getColumnModel().getColumn(i).setPreferredWidth(200);
			tabla.getColumnModel().getColumn(i).setMaxWidth(200);
			tabla.getColumnModel().getColumn(i).setMinWidth(200);
		}
		
		//Añadimos la tabla a un panel
		panel.add(tabla);
		
		//Añadimos barras de movimiento a la tabla.
		JScrollPane barras = new JScrollPane(panel, JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED,
				JScrollPane.HORIZONTAL_SCROLLBAR_AS_NEEDED);
		
		//Aumentamos la velocidad de movimiento de las barras
		barras.getVerticalScrollBar().setUnitIncrement(16);
		barras.getHorizontalScrollBar().setUnitIncrement(16);
		
		//Añadimos la tabla al frame
		ventanaPrincipal.add(barras);
		
		//Seteamos el valor de la fila 0 segun el numero
		for (int i = 0; i < fila; i++) {
			tabla.setValueAt(i + 1, i + 1, 0);
		}
		
		//Seteamos valor celdas a 0
		for (int i = 0; i < fila + 1; i++) {
			for (int j = 0; j < col + 1; j++) {
				if (i > 0 && j > 0) {
					tabla.setValueAt(0, i, j);
				}
			}
		}
		
		//Seteamos las columnas a A-B-etc
		int contAux = 26;
		int contAux2 = 702;
		for (int i = 0; i < col; i++) {
			String actual = "";
			tabla.getColumnModel().getColumn(i).setCellEditor(getCellEditor());
			if (i <= 25) {
				actual = "";
				char column = (char) (i + 65);
				tabla.setValueAt(column, 0, i + 1);
			} else if (i > 25 && i < 702) {
				char colum = (char) ((i % 26) + 65);
				char columAux = (char) ((i / 26) + 64);
				actual += columAux;
				actual += colum;
				tabla.setValueAt(actual, 0, i + 1);
			} else {
				if (i > 702 && (i) % 26 == 0) {
					contAux++;
				}
				if (i > 702 && (i - 26) % 676 == 0) {
					contAux2++;
				}
				char colum = (char) ((i % 26) + 65);
				char columAux = (char) (contAux % 26 + 65);
				char columAux2 = (char) (contAux2 % 702 + 65);
				actual += columAux2;
				actual += columAux;
				actual += colum;
				tabla.setValueAt(actual, 0, i + 1);
			}
		}
		
		//Añadimos el listener a la tabla
		tabla.addMouseListener(new EscuchaTabla());
		ventanaPrincipal.setTitle("Hoja de Calculo");
		ventanaPrincipal.setVisible(true);
		ventanaPrincipal.addWindowListener(new CerrandoPrograma());
	}

	//Editor de las celdas
	private TableCellEditor getCellEditor() {
		JTextField f = new JTextField();
		f.setBorder(BorderFactory.createLineBorder(Color.BLACK, 3));
		return new DefaultCellEditor(f);
	}

	/**
	 * Metodo que se encarga de pedir el numero de filas que tendra la tabla
	 */
	@SuppressWarnings("unlikely-arg-type")
	public void pedirFilas() {
		String filaS;
		//Panel para pedir filas
		filaS = JOptionPane.showInputDialog("Introduzca las filas: ");
		//setear numero de filas en caso de que sea erroneo lo vuelve a pedir, sino pide las columnas
		try {
			if (filaS.equals(JOptionPane.CLOSED_OPTION)) {
				// System.exit(0);
			}
			fila = Integer.parseInt(filaS);
			if (fila < 0) {
				JOptionPane.showMessageDialog(null, "Los valores deben ser positivos", "Error entrada",
						JOptionPane.ERROR_MESSAGE);
				pedirFilas();
			} else {
				pedirCols();
			}
		} catch (NumberFormatException e) {
			JOptionPane.showMessageDialog(null, "Los valores deben ser numericos", "Error entrada",
					JOptionPane.ERROR_MESSAGE);
			pedirFilas();
		} catch (NullPointerException e) {
			System.exit(0);
		}
	}

	public void pedirCols() {
		String colS;
		//Panel  para pedir columnas
		colS = JOptionPane.showInputDialog("Introduzca las columnas: ");
		//setear numero de columnas en caso de que sea erroneo lo vuelve a pedir, sino vuelve y crea la tabla
		try {
			if (colS.equals(JOptionPane.CLOSED_OPTION)) {
				// System.exit(0);
			}
			col = Integer.parseInt(colS);
			if (col == JOptionPane.CLOSED_OPTION) {
				System.exit(0);
			}
			if (col < 0) {
				JOptionPane.showMessageDialog(null, "Los valores deben ser positivos", "Error entrada",
						JOptionPane.ERROR_MESSAGE);
				pedirCols();
			}
		} catch (NumberFormatException e) {
			JOptionPane.showMessageDialog(null, "Los valores deben ser nï¿½mericos", "Error entrada",
					JOptionPane.ERROR_MESSAGE);
			pedirCols();
		} catch (NullPointerException e) {
			System.exit(0);
		}
	}

	//Invierte la cola de undo para que el ultimo añadido sea el primero
	public void invierteColaUndo() {
		Stack<CeldaGuardada> stack = new Stack<>();
		while (!colaUndo.isEmpty()) {
			stack.add(colaUndo.peek());
			colaUndo.remove();
		}
		while (!stack.isEmpty()) {
			colaUndo.add(stack.peek());
			stack.pop();
		}
	}
	
	//Invierte la cola de redo para que el ultimo añadido sea el primero
	public void invierteColaRedo() {
		Stack<CeldaGuardada> stack = new Stack<>();
		while (!colaRedo.isEmpty()) {
			stack.add(colaRedo.peek());
			colaRedo.remove();
		}
		while (!stack.isEmpty()) {
			colaRedo.add(stack.peek());
			stack.pop();
		}
	}

	/**
	 * Clase que desencadena siempre que se haga click en una celda de la tabla
	 * @author rauls
	 *
	 */
	class EscuchaTabla extends MouseAdapter {
		@Override
		public void mousePressed(MouseEvent e) {
			//Guardamos la celda que se ha seleccionado
			int filaActual = tabla.getSelectedRow();
			int colActual = tabla.getSelectedColumn();
			String valor2 = "";
			//Guardamos en una variable local el valor de la celda antes de hacer click
			if (filaSeleccionada > 0 && colSeleccionada > 0) {
				valor2 = String.valueOf(tab.getValueAt(filaSeleccionada, colSeleccionada));
			}
			//En caso de ser diferente guardamos el resultado en la cola Undo
			if (!valor.equals(valor2) && valor.trim().length() != 0 && valor2.trim().length() != 0) {
				CeldaGuardada celda = new CeldaGuardada(filaSeleccionada, colSeleccionada, valor, valor2);
				colaUndo.add(celda);
				deshacer.setEnabled(true);
				invierteColaUndo();
				colaRedo.clear();
				hacer.setEnabled(false);
				guardado = false;
			}
			valor = "";
			//En caso de que el valor sea "" seteamos el valor a 0
			if (valor2.trim().length() == 0 && filaSeleccionada > 0 && colSeleccionada > 0) {
				tabla.setValueAt(0, filaSeleccionada, colSeleccionada);
			}
			//Guardamos la ultima celda
			if (filaActual != 0 && colActual != 0) {
				filaSeleccionada = filaActual;
				colSeleccionada = colActual;
				tabla.repaint();
				String fila = String.valueOf(tab.getValueAt(tabla.getSelectedRow(), 0));
				String col = String.valueOf(tab.getValueAt(0, tabla.getSelectedColumn()));
				//Seteamos el valor del JLabel
				celda.setText(col + fila);
				valor = String.valueOf(tab.getValueAt(tabla.getSelectedRow(), tabla.getSelectedColumn()));
				if (valor.trim().charAt(0) == '0') {
					tabla.setValueAt("", tabla.getSelectedRow(), tabla.getSelectedColumn());
				}
			}
		}
	}
	
	
	/**
	 * Clase escuchadora del boton Limpiar
	 * Setea todas las filas al valor 0
	 * @author rauls
	 *
	 */
	class LimpiaCeldas implements ActionListener {
		@Override
		public void actionPerformed(ActionEvent arg0) {
			for (int i = 0; i < fila + 1; i++) {
				for (int j = 0; j < col + 1; j++) {
					if (i > 0 && j > 0) {
						tabla.setValueAt(0, i, j);
					}
				}
			}
			guardado=false;
		}
	}
	
	/**
	 * Clase encargada de la realización de las formulas
	 * @author rauls
	 *
	 */
	class CalculaFormulas implements ActionListener {
		@Override
		public void actionPerformed(ActionEvent e) {
			//En caso de que siga editando lo paramos
			if (tabla.isEditing()) {
				tabla.getCellEditor().stopCellEditing();
			}
			//En caso de que haya alguna celda con valor "" lo seteamos a 0
			for (int i = 0; i < fila; i++) {
				for (int j = 0; j < col; j++) {
					if (String.valueOf(tabla.getValueAt(i + 1, j + 1)).trim().length() == 0) {
						tabla.setValueAt(0, i + 1, j + 1);
					}
				}
			}
			//Creamos una hoja excel
			Excel hoja = new Excel(fila, col);
			//Creamos un array de string donde almacenaremos los valores
			String[][] valores = new String[fila][col];
			for (int i = 0; i < fila; i++) {
				for (int j = 0; j < col; j++) {
					valores[i][j] = String.valueOf(tabla.getValueAt(i + 1, j + 1)).trim();
				}
			}
			try {
				//Enviamos los valores a la hoja excel donde calcularemos
				if(hoja.esNumerico(valores)==-1) {
					JOptionPane.showMessageDialog(null, "Formato invalido", "Error entrada",
							JOptionPane.ERROR_MESSAGE);
					return;
				}
				//Extraemos los valores finales tras realizar las formulas
				valores = hoja.dameHoja();
				//Seteamos los valores
				for (int i = 0; i < fila; i++) {
					for (int j = 0; j < col; j++) {
						tabla.setValueAt(valores[i][j], i + 1, j + 1);
					}
				}
			//En caso de no ser una entrada correcta
			} catch (NumberFormatException e1) {
				JOptionPane.showMessageDialog(null, "Los valores deben ser numericos", "Error entrada",
						JOptionPane.ERROR_MESSAGE);
			} catch (ArrayIndexOutOfBoundsException e1) {
				JOptionPane.showMessageDialog(null, "Celda no existente", "Error entrada",
						JOptionPane.ERROR_MESSAGE);
			}
			
		}
	}
	
	/**
	 * Clase encargada de crear una nueva hoja
	 * @author rauls
	 *
	 */
	class NuevaHoja implements ActionListener {
		@Override
		public void actionPerformed(ActionEvent e) {
			pedirFilas();
			//Creamos el modelo
			tab = new DefaultTableModel(fila + 1, col + 1) {
				@Override
				public boolean isCellEditable(int row, int column) {
					if (row == 0 || column == 0) {
						return false;
					} else {
						return true;
					}
				}
			};
			//Seteamos el modelo
			tabla.setModel(tab);
			tabla.repaint();
			tabla.setVisible(true);
			//Ponemos el tamaño a las celdas
			tabla.setRowHeight(25);
			for (int i = 0; i < col + 1; i++) {
				tabla.getColumnModel().getColumn(i).setPreferredWidth(200);
				tabla.getColumnModel().getColumn(i).setMaxWidth(200);
				tabla.getColumnModel().getColumn(i).setMinWidth(200);
			}
			
			//Seteamos el valor inicial de cada celda
			for (int i = 0; i < fila; i++) {
				tabla.setValueAt(i + 1, i + 1, 0);
			}
			for (int i = 0; i < fila + 1; i++) {
				for (int j = 0; j < col + 1; j++) {
					if (i > 0 && j > 0) {
						tabla.setValueAt(0, i, j);
					}
				}
			}
			int contAux = 26;
			int contAux2 = 702;
			for (int i = 0; i < col; i++) {
				String actual = "";
				tabla.getColumnModel().getColumn(i).setCellEditor(getCellEditor());
				if (i <= 25) {
					actual = "";
					char column = (char) (i + 65);
					tabla.setValueAt(column, 0, i + 1);
				} else if (i > 25 && i < 702) {
					char colum = (char) ((i % 26) + 65);
					char columAux = (char) ((i / 26) + 64);
					actual += columAux;
					actual += colum;
					tabla.setValueAt(actual, 0, i + 1);
				} else {
					if (i > 702 && (i) % 26 == 0) {
						contAux++;
					}
					if (i > 702 && (i - 26) % 676 == 0) {
						contAux2++;
					}
					char colum = (char) ((i % 26) + 65);
					char columAux = (char) (contAux % 26 + 65);
					char columAux2 = (char) (contAux2 % 702 + 65);
					actual += columAux2;
					actual += columAux;
					actual += colum;
					tabla.setValueAt(actual, 0, i + 1);
				}
			}
			guardado=false;
			filaSeleccionada = 1;
			colSeleccionada = 1;
			celda.setText("A1");
		}
	}
	
	/**
	 * Clase encargada de abrir una hoja excel que tengamos almacenada en un archivo .txt
	 * @author rauls
	 *
	 */
	class AbrirArchivo implements ActionListener {
		@Override
		public void actionPerformed(ActionEvent e) {
			String leido = "";
			String filasCols = "";
			String filas = "";
			String cols = "";
			String valores[][];
			int colAux, filaAux;
			int i, contador = 0, contadorAux = 0;
			JFileChooser abrir = new JFileChooser();
			//Abrimos una ventana para seleccionar el archivo
			int option = abrir.showOpenDialog(null);
			//En caso de que se haya seleccionado un archivo
			if (option == JFileChooser.APPROVE_OPTION) {
				File archivo = abrir.getSelectedFile();
				try {
					//Leemos el archivo y guardamos todos los caracteres en la variable leido
					FileReader leer = new FileReader(archivo);
					while ((i = leer.read()) != -1) {
						leido += (char) i;
					}
					//Almacenamos las columnas y filas en un nuevo String
					for (int j = 0; j < leido.length(); j++) {
						if (leido.charAt(j) == '\n') {
							leido = leido.substring(j + 1, leido.length());
							break;
						}
						filasCols += leido.charAt(j);
					}
					leer.close();
					
					//Extraemos el valor de las filas y columnas en un string que luego transformaremos en int
					for (int j = 0; j < filasCols.length(); j++) {
						if (contador == 0) {
							cols += filasCols.charAt(j);
							if (filasCols.charAt(j) == ' ')
								contador++;
						} else {
							filas += filasCols.charAt(j);
						}
					}
					//Transformamos en int
					colAux = Integer.parseInt(cols.trim());
					filaAux = Integer.parseInt(filas.trim());
					if (colAux > 0 && filaAux > 0) {
						col = Integer.parseInt(cols.trim());
						fila = Integer.parseInt(filas.trim());
					//En caso de que el valor de las columnas sea negativo
					} else {
						JOptionPane.showMessageDialog(null, "Las filas y columnas deben ser positivas", "Error entrada",
								JOptionPane.ERROR_MESSAGE);
						return;
					}
					//Almacenamos en un string el valor que tendrá cada celda de la hoja cargada
					valores = new String[fila][col];
					for (int j = 0; j < fila; j++) {
						for (int k = 0; k < col; k++) {
							valores[j][k] = "";
						}
					}
					contador = 0;
					StringTokenizer cuenta = new StringTokenizer(leido);
					if ((int) cuenta.countTokens() != (fila * col)) {
						JOptionPane.showMessageDialog(null, "Numero valores invalido", "Error entrada",
								JOptionPane.ERROR_MESSAGE);
						return;
					}
					//Introducimos los valores en la hojad e calculo
					try {
						for (int j = 0; j < leido.trim().length(); j++) {
							if (leido.trim().charAt(j) == ' ') {
								contadorAux++;
							} else if (leido.trim().charAt(j) == '\n') {
								contadorAux = 0;
								contador++;
							} else {
								valores[contador][contadorAux] += leido.trim().charAt(j);
							}
						}
					} catch (ArrayIndexOutOfBoundsException e1) {
						JOptionPane.showMessageDialog(null, "Numero valores invalido", "Error entrada",
								JOptionPane.ERROR_MESSAGE);
						return;
					}
					//Creamos el modelo de la tabla
					tab = new DefaultTableModel(fila + 1, col + 1) {
						@Override
						public boolean isCellEditable(int row, int column) {
							if (row == 0 || column == 0) {
								return false;
							} else {
								return true;
							}
						}
					};
					tabla.setModel(tab);
					tabla.repaint();
					tabla.setVisible(true);
					
					//Seteamos el tamaña de las celdas
					tabla.setRowHeight(25);
					for (int k = 0; k < col + 1; k++) {
						tabla.getColumnModel().getColumn(k).setPreferredWidth(200);
						tabla.getColumnModel().getColumn(k).setMaxWidth(200);
						tabla.getColumnModel().getColumn(k).setMinWidth(200);
					}
					
					//Seteamos el valor que tendra cada celda
					for (int k = 0; k < fila; k++) {
						tabla.setValueAt(k + 1, k + 1, 0);
					}
					for (int k = 0; k < fila + 1; k++) {
						for (int j = 0; j < col + 1; j++) {
							if (k > 0 && j > 0) {
								tabla.setValueAt(valores[k - 1][j - 1].trim(), k, j);
							}
						}
					}
					int contAux = 26;
					int contAux2 = 702;
					for (int j = 0; j < col; j++) {
						String actual = "";
						tabla.getColumnModel().getColumn(j).setCellEditor(getCellEditor());
						if (j <= 25) {
							actual = "";
							char column = (char) (j + 65);
							tabla.setValueAt(column, 0, j + 1);
						} else if (j > 25 && j < 702) {
							char colum = (char) ((j % 26) + 65);
							char columAux = (char) ((j / 26) + 64);
							actual += columAux;
							actual += colum;
							tabla.setValueAt(actual, 0, j + 1);
						} else {
							if (j > 702 && (j) % 26 == 0) {
								contAux++;
							}
							if (j > 702 && (j - 26) % 676 == 0) {
								contAux2++;
							}
							char colum = (char) ((j % 26) + 65);
							char columAux = (char) (contAux % 26 + 65);
							char columAux2 = (char) (contAux2 % 702 + 65);
							actual += columAux2;
							actual += columAux;
							actual += colum;
							tabla.setValueAt(actual, 0, j + 1);
						}
					}
				} catch (FileNotFoundException e1) {
					System.out.println("No se ha encontrado el archivo");
				} catch (IOException e1) {
					e1.printStackTrace();
				} catch (NumberFormatException e1) {
					JOptionPane.showMessageDialog(null, "Los valores deben ser numericos", "Error entrada",
							JOptionPane.ERROR_MESSAGE);
				}
			}
			celda.setText("A1");
			filaSeleccionada = 1;
			colSeleccionada = 1;
		}
	}
	
	/**
	 * Clase encargada de archivar una hoja excel
	 * @author rauls
	 *
	 */
	class ArchivarArchivo implements ActionListener {
		@Override
		public void actionPerformed(ActionEvent arg0) {
			String valores = "";
			//Si hay alguna columna que valga "" la seteamos a 0
			for (int i = 0; i < fila; i++) {
				for (int j = 0; j < col; j++) {
					if (String.valueOf(tabla.getValueAt(i + 1, j + 1)).trim().length() == 0) {
						tabla.setValueAt(0, i + 1, j + 1);
					}
				}
			}
			//Almacenamos las filas y columnas de la hoja
			valores += col + " " + fila + "\n";
			
			//Almacenamos el valor de cada celda
			for (int i = 0; i < fila; i++) {
				for (int j = 0; j < col; j++) {
					valores += String.valueOf(tabla.getValueAt(i + 1, j + 1)).trim();
					valores += " ";
				}
				valores += "\n";
			}
			
			JFileChooser guardar = new JFileChooser();
			int option = guardar.showSaveDialog(null);
			if (option == JFileChooser.APPROVE_OPTION) {
				try {
					//Escribimos en el archivo
					FileWriter escribir = new FileWriter(guardar.getSelectedFile() + ".txt");
					escribir.write(valores);
					escribir.close();
					guardado = true;
				} catch (IOException e1) {
					e1.printStackTrace();
				}
			}
		}
	}
	
	/**
	 * Clase encargada de cerrar el programa
	 * @author rauls
	 *
	 */
	class CerrandoPrograma extends WindowAdapter {
		@Override
		public void windowClosing(WindowEvent e) {
			//En caso de que todo este guardado
			if (guardado) {
				int result = JOptionPane.showConfirmDialog(null, "Quieres cerrar el programa?", "Cerrar",
						JOptionPane.YES_NO_OPTION, JOptionPane.QUESTION_MESSAGE);
				if (result == JOptionPane.YES_OPTION) {
					System.exit(0);
				}
			//En caso de que haya algo sin guardar
			} else {
				int result = JOptionPane.showConfirmDialog(null, "Tienes datos sin guardar quieres guardar?", "Cerrar",
						JOptionPane.YES_NO_CANCEL_OPTION, JOptionPane.QUESTION_MESSAGE);
				if (result == JOptionPane.YES_OPTION) {
					new ArchivarArchivo().actionPerformed(null);
					System.exit(0);
				} else if (result == JOptionPane.NO_OPTION) {
					System.exit(0);
				}
			}
		}
	}
	
	/**
	 * Clase encargada de hacer lo referente a la función undo
	 * @author rauls
	 *
	 */
	class Undo implements ActionListener {
		@Override
		public void actionPerformed(ActionEvent e) {
			//En caso de que haya algo en la cola
			if (colaUndo.size() > 0) {
				//En caso de que se este ediando dejamos de editar
				if (tabla.isEditing()) {
					tabla.getCellEditor().stopCellEditing();
				}
				//En caso de que la celda el valor sea "" lo seteamos a 0
				if (filaSeleccionada > 0 && colSeleccionada > 0) {
					if (String.valueOf(tabla.getValueAt(filaSeleccionada, colSeleccionada)).trim().length() == 0) {
						tabla.setValueAt(0, filaSeleccionada, colSeleccionada);
					}
				}
				
				filaSeleccionada = -1;
				colSeleccionada = -1;
				
				//Cogemos el ultimo valor modificado
				ultimaCelda = colaUndo.poll();
				int fila = ultimaCelda.getFila();
				int col = ultimaCelda.getCol();
				
				//Añadimos el valor hacia atras a la cola de hacia delante
				colaRedo.add(ultimaCelda);
				
				//Seteamos el valor
				String valor = ultimaCelda.getValor();
				tabla.setValueAt(valor, fila, col);
				tabla.repaint();
				//Invertimos las colas
				invierteColaUndo();
				invierteColaRedo();
				hacer.setEnabled(true);
			}
			if (colaUndo.size() == 0) {
				deshacer.setEnabled(false);
			}
		}
	}
	
	/**
	 * Clase encargada de hacer lo referente a redo
	 * @author rauls
	 *
	 */
	class Redo implements ActionListener {
		@Override
		public void actionPerformed(ActionEvent e) {
			if (colaRedo.size() != 0) {
				//En caso de que se este ediando dejamos de editar
				if (tabla.isEditing()) {
					tabla.getCellEditor().stopCellEditing();
				}
				//En caso de que la celda el valor sea "" lo seteamos a 0
				if (filaSeleccionada > 0 && colSeleccionada > 0) {
					if (String.valueOf(tabla.getValueAt(filaSeleccionada, colSeleccionada)).trim().length() == 0) {
						tabla.setValueAt(0, filaSeleccionada, colSeleccionada);
					}
				}
				//Cogemos el ultimo valor de la cola redo
				ultimaCelda = colaRedo.poll();
				int fila = ultimaCelda.getFila();
				int col = ultimaCelda.getCol();
				
				//Seteamos el valor
				String valor = ultimaCelda.getValorFinal();
				tabla.setValueAt(valor, fila, col);
				invierteColaRedo();
				
				//Lo añadimos a la cola undo
				colaUndo.add(ultimaCelda);
				invierteColaUndo();
				deshacer.setEnabled(true);
			}
			if (colaRedo.size() == 0) {
				hacer.setEnabled(false);
			}
		}
	}
}
