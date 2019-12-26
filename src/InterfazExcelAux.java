import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Component;
import java.awt.Dimension;
import java.awt.Font;
import java.awt.GridLayout;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.KeyEvent;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.awt.event.MouseListener;
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
import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JMenu;
import javax.swing.JMenuBar;
import javax.swing.JMenuItem;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollBar;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import javax.swing.KeyStroke;
import javax.swing.SwingConstants;
import javax.swing.border.Border;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableModel;

public class InterfazExcelAux extends JFrame {

	JTable tabla;
	int fila;
	int col;
	int filaSeleccionada = 1;
	int colSeleccionada = 1;
	JLabel celda;
	DefaultTableModel tab;
	Queue<CeldaGuardada> colaUndo;
	Queue<CeldaGuardada> colaRedo;
	Dimension screenSize;
	JFrame ventanaPrincipal;
	boolean guardado;
	String valor;
	JMenuItem hacer;
	JMenuItem deshacer;
	CeldaGuardada ultimaCelda;

	public InterfazExcelAux() {
		screenSize = Toolkit.getDefaultToolkit().getScreenSize();
		pedirFilas();
		colaUndo = new LinkedList<CeldaGuardada>();
		colaRedo = new LinkedList<CeldaGuardada>();
		JMenuBar menuBarra = new JMenuBar();
		JMenu menu = new JMenu("Archivo");
		JMenuItem archivar = new JMenuItem("Archivar");
		archivar.addActionListener(new ArchivarArchivo());

		JMenuItem cargar = new JMenuItem("Cargar");
		cargar.addActionListener(new AbrirArchivo());

		JMenuItem nueva = new JMenuItem("Nuevo");
		nueva.addActionListener(new NuevaHoja());
		menu.add(archivar);
		menu.add(cargar);
		menu.add(nueva);
		JMenu editar = new JMenu("Editar");
		hacer = new JMenuItem("Adelante");
		deshacer = new JMenuItem("Atras");
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
		guardado = false;
		JPanel botones = new JPanel();
		JButton calcular = new JButton("Calcular");
		JButton limpiar = new JButton("Limpiar");
		calcular.addActionListener(new CalculaFormulas());
		limpiar.addActionListener(new LimpiaCeldas());
		celda = new JLabel("A1");
		botones.add(calcular);
		botones.add(limpiar);
		botones.add(celda);
		deshacer.setEnabled(false);
		deshacer.addActionListener(new Undo());
		KeyStroke ctrlZ = KeyStroke.getKeyStroke("control Z");
		deshacer.setAccelerator(ctrlZ);
		hacer.setEnabled(false);
		hacer.addActionListener(new Redo());
		KeyStroke mayusctrlZ = KeyStroke.getKeyStroke(KeyEvent.VK_Z, KeyEvent.SHIFT_MASK | KeyEvent.CTRL_MASK);
		hacer.setAccelerator(mayusctrlZ);
		editar.add(hacer);
		editar.add(deshacer);
		ventanaPrincipal = new JFrame();
		tabla = new JTable(tab);
		JPanel panel = new JPanel();
		ventanaPrincipal.setDefaultCloseOperation(JFrame.DO_NOTHING_ON_CLOSE);
		// ventanaPrincipal.setBounds(0, 0, (int) screenSize.getWidth(), (int)
		// screenSize.getHeight());
		ventanaPrincipal.setBounds(0, 0, 600, 600);
		ventanaPrincipal.setVisible(true);
		tabla.setRowHeight(25);
		DefaultTableCellRenderer render = new DefaultTableCellRenderer() {
			public Component getTableCellRendererComponent(JTable jTable1, Object value, boolean selected,
					boolean focused, int row, int column) {
				setHorizontalAlignment(SwingConstants.CENTER);
				super.getTableCellRendererComponent(jTable1, value, selected, focused, row, column);
				if (row == 0 && column == 0) {
					this.setBackground(Color.gray);
				} else if (row == 0 || column == 0) {
					this.setBackground(Color.GREEN);
				} else {
					this.setBackground(Color.WHITE);
				}
				if (row == filaSeleccionada && column == colSeleccionada && filaSeleccionada != 0
						&& colSeleccionada != 0) {
				} else if (column == 0 && row == filaSeleccionada && filaSeleccionada != 0 && colSeleccionada != 0) {
					this.setBackground(Color.YELLOW);
				} else if (row == 0 && column == colSeleccionada && filaSeleccionada != 0 && colSeleccionada != 0) {
					this.setBackground(Color.yellow);
				}
				if (row == filaSeleccionada && column == colSeleccionada) {
					setBorder(BorderFactory.createLineBorder(Color.BLACK, 3));
				} else {
					setBorder(BorderFactory.createLineBorder(Color.BLACK, 1));
				}
				return this;
			}
		};
		valor = "";
		menuBarra.add(menu);
		menuBarra.add(editar);
		ventanaPrincipal.setJMenuBar(menuBarra);
		tabla.setDefaultRenderer(Object.class, render);
		tabla.setFont(new Font("Arial", Font.BOLD, 15));
		tabla.setCellSelectionEnabled(true);
		ventanaPrincipal.add(botones, BorderLayout.SOUTH);
		for (int i = 0; i < col + 1; i++) {
			tabla.getColumnModel().getColumn(i).setPreferredWidth(200);
			tabla.getColumnModel().getColumn(i).setMaxWidth(200);
			tabla.getColumnModel().getColumn(i).setMinWidth(200);
		}
		panel.add(tabla);
		JScrollPane barras = new JScrollPane(panel, JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED,
				JScrollPane.HORIZONTAL_SCROLLBAR_AS_NEEDED);
		barras.getVerticalScrollBar().setUnitIncrement(16);
		barras.getHorizontalScrollBar().setUnitIncrement(16);
		ventanaPrincipal.add(barras);
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
		tabla.addMouseListener(new EscuchaTabla());
		ventanaPrincipal.setTitle("Hoja de Calculo");
		ventanaPrincipal.setVisible(true);
		ventanaPrincipal.addWindowListener(new CerrandoPrograma());
	}

	@SuppressWarnings("unlikely-arg-type")
	public void pedirFilas() {
		String filaS;
		filaS = JOptionPane.showInputDialog("Introduzca las filas: ");
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
		colS = JOptionPane.showInputDialog("Introduzca las columnas: ");
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
			JOptionPane.showMessageDialog(null, "Los valores deben ser n�mericos", "Error entrada",
					JOptionPane.ERROR_MESSAGE);
			pedirCols();
		} catch (NullPointerException e) {
			System.exit(0);
		}
	}

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

	class EscuchaTabla extends MouseAdapter {
		@Override
		public void mousePressed(MouseEvent e) {
			int filaActual = tabla.getSelectedRow();
			int colActual = tabla.getSelectedColumn();
			String valor2 = String.valueOf(tab.getValueAt(filaSeleccionada, colSeleccionada));
			if (!valor.equals(valor2) && valor.trim().length() != 0 && valor2.trim().length() != 0) {
				System.out.println("Almacenando: " + filaSeleccionada + "" + colSeleccionada + " --> " + valor);
				CeldaGuardada celda = new CeldaGuardada(filaSeleccionada, colSeleccionada, valor, valor2);
				colaUndo.add(celda);
				deshacer.setEnabled(true);
				invierteColaUndo();
				guardado = false;
			}
			valor = "";
			if (valor2.trim().length() == 0) {
				tabla.setValueAt(0, filaSeleccionada, colSeleccionada);
			}
			if (filaActual != 0 && colActual != 0) {
				filaSeleccionada = filaActual;
				colSeleccionada = colActual;
				tabla.repaint();
				String fila = String.valueOf(tab.getValueAt(tabla.getSelectedRow(), 0));
				String col = String.valueOf(tab.getValueAt(0, tabla.getSelectedColumn()));
				celda.setText(col + fila);
				valor = String.valueOf(tab.getValueAt(tabla.getSelectedRow(), tabla.getSelectedColumn()));
				if (valor.trim().charAt(0) == '0') {
					tabla.setValueAt("", tabla.getSelectedRow(), tabla.getSelectedColumn());
				}
			}
		}
	}

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
		}
	}

	class CalculaFormulas implements ActionListener {
		@Override
		public void actionPerformed(ActionEvent e) {
			for (int i = 0; i < fila; i++) {
				for (int j = 0; j < col; j++) {
					if (String.valueOf(tabla.getValueAt(i + 1, j + 1)).trim().length() == 0) {
						tabla.setValueAt(0, i + 1, j + 1);
					}
				}
			}
			Excel hoja = new Excel(fila, col);
			String[][] valores = new String[fila][col];
			for (int i = 0; i < fila; i++) {
				for (int j = 0; j < col; j++) {
					valores[i][j] = String.valueOf(tabla.getValueAt(i + 1, j + 1)).trim();
				}
			}
			try {
				hoja.esNumerico(valores);
				valores = hoja.dameHoja();
				for (int i = 0; i < fila; i++) {
					for (int j = 0; j < col; j++) {
						tabla.setValueAt(valores[i][j], i + 1, j + 1);
					}
				}
			} catch (NumberFormatException e1) {
				JOptionPane.showMessageDialog(null, "Los valores deben ser numericos", "Error entrada",
						JOptionPane.ERROR_MESSAGE);
			}
		}
	}

	class NuevaHoja implements ActionListener {
		@Override
		public void actionPerformed(ActionEvent e) {
			pedirFilas();
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
			tabla.setRowHeight(25);
			for (int i = 0; i < col + 1; i++) {
				tabla.getColumnModel().getColumn(i).setPreferredWidth(200);
				tabla.getColumnModel().getColumn(i).setMaxWidth(200);
				tabla.getColumnModel().getColumn(i).setMinWidth(200);
			}
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
			filaSeleccionada = 1;
			colSeleccionada = 1;
		}
	}

	class AbrirArchivo implements ActionListener {
		@SuppressWarnings({ "serial", "serial" })
		@Override
		public void actionPerformed(ActionEvent e) {
			String leido = "";
			String filasCols = "";
			String filas = "";
			String cols = "";
			String valores[][];
			int i, contador = 0, contadorAux = 0;
			JFileChooser abrir = new JFileChooser();
			int option = abrir.showOpenDialog(null);
			if (option == JFileChooser.APPROVE_OPTION) {
				File archivo = abrir.getSelectedFile();
				try {
					FileReader leer = new FileReader(archivo);
					while ((i = leer.read()) != -1) {
						leido += (char) i;
					}
					for (int j = 0; j < leido.length(); j++) {
						if (leido.charAt(j) == '\n') {
							leido = leido.substring(j + 1, leido.length());
							break;
						}
						filasCols += leido.charAt(j);
					}
					leer.close();
					filasCols = filasCols.substring(0, filasCols.length() - 1);
					for (int j = 0; j < filasCols.length(); j++) {
						if (contador == 0) {
							cols += filasCols.charAt(j);
							if (filasCols.charAt(j) == ' ')
								contador++;
						} else {
							filas += filasCols.charAt(j);
						}
					}
					cols = cols.substring(0, 1);
					col = Integer.parseInt(cols);
					fila = Integer.parseInt(filas);
					valores = new String[fila][col];
					for (int j = 0; j < fila; j++) {
						for (int k = 0; k < col; k++) {
							valores[j][k] = "";
						}
					}
					contador = 0;
					for (int j = 0; j < leido.length(); j++) {
						valores[contador][contadorAux] += leido.charAt(j);
						StringTokenizer cuenta = new StringTokenizer(leido);
						if (cuenta.countTokens() != fila * col) {
							JOptionPane.showMessageDialog(null, "Numero valores invalido", "Error entrada",
									JOptionPane.ERROR_MESSAGE);
							return;
						}
						if (leido.charAt(j) == ' ') {
							contadorAux++;
						} else if (leido.charAt(j) == '\n') {
							contadorAux = 0;
							contador++;
						}
					}
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
					tabla.setRowHeight(25);
					for (int k = 0; k < col + 1; k++) {
						tabla.getColumnModel().getColumn(k).setPreferredWidth(200);
						tabla.getColumnModel().getColumn(k).setMaxWidth(200);
						tabla.getColumnModel().getColumn(k).setMinWidth(200);
					}
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
					JOptionPane.showMessageDialog(null, "Los valores deben ser n�mericos", "Error entrada",
							JOptionPane.ERROR_MESSAGE);
				}
			}
			filaSeleccionada = 1;
			colSeleccionada = 1;
		}
	}

	class ArchivarArchivo implements ActionListener {
		@Override
		public void actionPerformed(ActionEvent arg0) {
			String valores = "";
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

	class CerrandoPrograma extends WindowAdapter {
		@Override
		public void windowClosing(WindowEvent e) {
			if (guardado) {
				int result = JOptionPane.showConfirmDialog(null, "Quieres cerrar el programa?", "Cerrar",
						JOptionPane.YES_NO_OPTION, JOptionPane.QUESTION_MESSAGE);
				if (result == JOptionPane.YES_OPTION) {
					System.exit(0);
				}
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

	class Undo implements ActionListener {
		@Override
		public void actionPerformed(ActionEvent e) {
			if (colaUndo.size() != 0) {
				ultimaCelda = colaUndo.poll();
				colaRedo.add(ultimaCelda);
				int fila = ultimaCelda.getFila();
				int col = ultimaCelda.getCol();
				String valor = ultimaCelda.getValor();
				System.out.println("Deshacer: " + fila + "" + col + " --> " + valor);
				tabla.setValueAt(valor, fila, col);
				invierteColaUndo();
				invierteColaRedo();
				hacer.setEnabled(true);
			}
			if (colaUndo.size() == 0) {
				deshacer.setEnabled(false);
			}
		}
	}

	class Redo implements ActionListener {
		@Override
		public void actionPerformed(ActionEvent e) {
			if (colaRedo.size() != 0) {
				ultimaCelda = colaRedo.poll();
				int fila = ultimaCelda.getFila();
				int col = ultimaCelda.getCol();
				String valor = ultimaCelda.getValorFinal();
				System.out.println("Hacer: " + fila + "" + col + " --> " + valor);
				tabla.setValueAt(valor, fila, col);
				invierteColaRedo();
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
