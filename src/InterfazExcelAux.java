import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Component;
import java.awt.Dimension;
import java.awt.Font;
import java.awt.GridLayout;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
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
	Queue<String> cola;
	Dimension screenSize;
	JFrame ventanaPrincipal;

	public InterfazExcelAux() {
		screenSize = Toolkit.getDefaultToolkit().getScreenSize();
		pedirFilas();
		cola = new LinkedList<String>();
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
		JMenuItem hacer = new JMenuItem("Adelante");
		JMenuItem deshacer = new JMenuItem("Atras");
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
		JPanel botones = new JPanel();
		JButton calcular = new JButton("Calcular");
		JButton limpiar = new JButton("Limpiar");
		calcular.addActionListener(new CalculaFormulas());
		limpiar.addActionListener(new LimpiaCeldas());
		celda = new JLabel("A1");
		botones.add(calcular);
		botones.add(limpiar);
		botones.add(celda);
		editar.add(hacer);
		editar.add(deshacer);
		ventanaPrincipal = new JFrame();
		tabla = new JTable(tab);
		JPanel panel = new JPanel();
		ventanaPrincipal.setDefaultCloseOperation(JFrame.DO_NOTHING_ON_CLOSE);
		ventanaPrincipal.setBounds(0, 0, (int) screenSize.getWidth(), (int) screenSize.getHeight());
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

	public void pedirFilas() {
		String filaS;
		filaS = JOptionPane.showInputDialog("Introduzca las filas: ");
		try {
			fila = Integer.parseInt(filaS);
			if (fila == JOptionPane.CLOSED_OPTION) {
				System.exit(0);
			}
			if (fila < 0) {
				JOptionPane.showMessageDialog(null, "Los valores deben ser positivos", "Error entrada",
						JOptionPane.ERROR_MESSAGE);
				pedirFilas();
			} else {
				pedirCols();
			}
		} catch (NumberFormatException e) {
			JOptionPane.showMessageDialog(null, "Los valores deben ser númericos", "Error entrada",
					JOptionPane.ERROR_MESSAGE);
			pedirFilas();
		}
	}

	public void pedirCols() {
		String colS;
		colS = JOptionPane.showInputDialog("Introduzca las columnas: ");
		try {
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
			JOptionPane.showMessageDialog(null, "Los valores deben ser númericos", "Error entrada",
					JOptionPane.ERROR_MESSAGE);
			pedirCols();
		}
	}

	public void invierteCola() {
		Stack<String> stack = new Stack<>();
		while (!cola.isEmpty()) {
			stack.add(cola.peek());
			cola.remove();
		}
		while (!stack.isEmpty()) {
			cola.add(stack.peek());
			stack.pop();
		}
	}

	class EscuchaTabla extends MouseAdapter {
		@Override
		public void mousePressed(MouseEvent e) {
			invierteCola();
			int filaActual = tabla.getSelectedRow();
			int colActual = tabla.getSelectedColumn();
			String valor2 = String.valueOf(tab.getValueAt(filaSeleccionada, colSeleccionada));
			if (valor2.length() == 0) {
				tabla.setValueAt(0, filaSeleccionada, colSeleccionada);
			}
			if (filaActual != 0 && colActual != 0) {
				filaSeleccionada = filaActual;
				colSeleccionada = colActual;
				tabla.repaint();
				String fila = String.valueOf(tab.getValueAt(tabla.getSelectedRow(), 0));
				String col = String.valueOf(tab.getValueAt(0, tabla.getSelectedColumn()));
				celda.setText(col + fila);
				String valor = String.valueOf(tab.getValueAt(tabla.getSelectedRow(), tabla.getSelectedColumn()));
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
			Excel hoja = new Excel(fila, col);
			String[][] valores = new String[fila][col];
			for (int i = 0; i < fila; i++) {
				for (int j = 0; j < col; j++) {
					valores[i][j] = String.valueOf(tabla.getValueAt(i + 1, j + 1)).trim();
				}
			}
			hoja.esNumerico(valores);
			valores = hoja.dameHoja();
			for (int i = 0; i < fila; i++) {
				for (int j = 0; j < col; j++) {
					tabla.setValueAt(valores[i][j], i + 1, j + 1);
				}
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
					JOptionPane.showMessageDialog(null, "Los valores deben ser númericos", "Error entrada",
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
				} catch (IOException e1) {
					e1.printStackTrace();
				}
			}
		}
	}

	class CerrandoPrograma extends WindowAdapter {
		public void windowClosing(WindowEvent e) {
			int result = JOptionPane.showConfirmDialog(null,
					"Quieres cerrar el programa?", "Cerrar", JOptionPane.YES_NO_OPTION, JOptionPane.QUESTION_MESSAGE);
			if (result == JOptionPane.YES_OPTION) {
				System.exit(0);
			}
		}
	}
}
