import java.awt.Color;
import java.awt.GridLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;

import javax.swing.BorderFactory;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JMenu;
import javax.swing.JMenuBar;
import javax.swing.JMenuItem;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JTextField;
import javax.swing.SwingConstants;
import javax.swing.border.Border;
import javax.swing.event.UndoableEditEvent;
import javax.swing.event.UndoableEditListener;
import javax.swing.undo.UndoManager;
//imagen.jpg.b64
//defaultTabModel
//abastractTableModel
public class InterfazExcel extends JPanel {
	static JPanel panelPrincipal;
	static JTextField texto;
	static JTextField[][] celdas;
	static JTextField celdaMarca;
	static JFrame ventanaPrincipal;
	int fila;
	int col;
	static UndoManager opciones;

	public InterfazExcel() {
		ventanaPrincipal = new JFrame();
		String filaS = JOptionPane.showInputDialog("Introduzca el numero de filas: ");
		String colS = JOptionPane.showInputDialog("Introduzca el numero de columnas: ");
		fila = Integer.parseInt(filaS);
		col = Integer.parseInt(colS);
		celdas = new JTextField[col][fila];
		panelPrincipal = new JPanel(new GridLayout(fila, col));
		opciones=new UndoManager();
		colocaCeldas(fila, col);
		JMenuBar menuBarra = new JMenuBar();

		JMenu menu = new JMenu("Archivo");
		JMenuItem archivar = new JMenuItem("Archivar");
		archivar.addActionListener(new ArchivarArchivo());

		JMenuItem cargar = new JMenuItem("Cargar");
		cargar.addActionListener(new AbrirArchivo());

		JMenuItem nueva = new JMenuItem("Nuevo");
		nueva.addActionListener(new ActionListener() {

			@Override
			public void actionPerformed(ActionEvent arg0) {
				for (int i = 0; i < col; i++) {
					for (int j = 0; j < fila; j++) {
						panelPrincipal.remove(celdas[i][j]);
					}
				}
				String filaS=JOptionPane.showInputDialog("Introduzca el numero de filas:");
				String colS=JOptionPane.showInputDialog("Introduzca el numero de filas:");
				fila=Integer.parseInt(filaS);
				col=Integer.parseInt(colS);
				panelPrincipal.setLayout(new GridLayout(fila, col));
				celdas=new JTextField[col][fila];
				colocaCeldas(fila, col);
				ventanaPrincipal.setBounds(0,0,401,401);
				revalidate();
				repaint();
			}
		});
		menu.add(archivar);
		menu.add(cargar);
		menu.add(nueva);
		JMenu editar = new JMenu("Editar");
		JMenuItem hacer = new JMenuItem("Adelante");
		hacer.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				opciones.redo();		
			}
			
		});
		JMenuItem deshacer = new JMenuItem("Atras");
		deshacer.addActionListener(new ActionListener() {

			@Override
			public void actionPerformed(ActionEvent e) {
				opciones.undo();
			}
			
		});
		editar.add(hacer);
		editar.add(deshacer);

		menuBarra.add(menu);
		menuBarra.add(editar);
		ventanaPrincipal.setJMenuBar(menuBarra);
		ventanaPrincipal.add(panelPrincipal);

		panelPrincipal.setBorder((BorderFactory.createEmptyBorder(10, 10, 10, 10)));

		ventanaPrincipal.setBounds(0, 0, 400, 400);
		ventanaPrincipal.setVisible(true);
		ventanaPrincipal.setTitle("Hoja de calculo");
		ventanaPrincipal.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
	}

	public static void colocaCeldas(int fila, int col) {
		for (int i = 0; i < col; i++) {
			for (int j = 0; j < fila; j++) {
				celdas[i][j] = new JTextField();
				celdas[i][j].setHorizontalAlignment(SwingConstants.CENTER);
				if (i == 0 && j == 0) {
					Border border = BorderFactory.createLineBorder(Color.BLACK, 3);
					celdas[i][j].setBorder(border);
					celdaMarca = celdas[i][j];
				} else {
					celdas[i][j].setBorder(BorderFactory.createLineBorder(Color.black, 1));
				}
				panelPrincipal.add(celdas[i][j]);
				celdas[i][j].addMouseListener(new CasillaPulsada());
				celdas[i][j].getDocument().addUndoableEditListener(new UndoableEditListener() {

					@Override
					public void undoableEditHappened(UndoableEditEvent arg0) {
						opciones.addEdit(arg0.getEdit());
						
					}
					
				});
			}
		}

	}
}

class ArchivarArchivo implements ActionListener {

	JFileChooser guardar;

	public void actionPerformed(ActionEvent e) {
		guardar=new JFileChooser();
		int option=guardar.showSaveDialog(null);
		if(option==JFileChooser.APPROVE_OPTION) {
			try {
				FileWriter escribir=new FileWriter(guardar.getSelectedFile()+".txt");
				escribir.write("HOLA QUE TAL");
				escribir.close();
			} catch (IOException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}
		}
	}
}

class CasillaPulsada extends MouseAdapter {

	public void mousePressed(MouseEvent e) {
		if (e.getButton() == MouseEvent.BUTTON1) {
			InterfazExcel.celdaMarca.setBorder(BorderFactory.createLineBorder(Color.black, 1));
			JTextField bordeClicked = (JTextField) e.getComponent();
			bordeClicked.setBorder(BorderFactory.createLineBorder(Color.black, 3));
			InterfazExcel.celdaMarca = bordeClicked;
		}
	}
}

class AbrirArchivo implements ActionListener{

	@Override
	public void actionPerformed(ActionEvent e) {
		String leido="";
		int i;
		JFileChooser abrir=new JFileChooser();
		int option=abrir.showOpenDialog(null);
		if(option==JFileChooser.APPROVE_OPTION) {
			File archivo=abrir.getSelectedFile();
			try {
				FileReader leer=new FileReader(archivo);
				try {
					while((i=leer.read())!=-1) {
						leido+=(char)i;
					}
					System.out.println(leido);
					analizaLeido(leido);
				} catch (IOException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
			} catch (FileNotFoundException e1) {
				System.out.println("No se ha encontrado el archivo");
			}
		}
	}
	
	public void analizaLeido(String texto) {

	}
}
