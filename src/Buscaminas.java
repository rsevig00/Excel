import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.GridLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JMenu;
import javax.swing.JMenuBar;
import javax.swing.JMenuItem;
import javax.swing.JOptionPane;
import javax.swing.JPanel;

public class Buscaminas extends JFrame{
    int lado=10;
    JButton[][] posiciones;
    int [][] vecinos;
    int numMinas;
    boolean gameover=false;
    JMenuBar barra;
    JMenu juego, tamanio;
    JMenuItem opcNuevo, opc1, opc2, opc3, opcSalir;
    int casillasClicadas=0;
    JPanel pano;
    
    private void clicaCasilla(int i, int j)
    {
    	if (posiciones[i][j].getBackground()==Color.GREEN) casillasClicadas++;
		posiciones[i][j].setBackground(Color.gray);
		int numVecinos=vecinos[i][j];
		if (numVecinos>0) posiciones[i][j].setText(new Integer(numVecinos).toString());
    }
    
    private void calculaMinasVecinos()
    {
    	numMinas=(int) (1.5*lado);
    	  for (int num =0; num<numMinas; num++)
	        {
	            int x=(int) (Math.random()*lado);
	            int y=(int) (Math.random()*lado);
	            
	            while (vecinos[x][y]==-1)
	            {
	                x=(int) (Math.random()*lado);
	                y=(int) (Math.random()*lado);
	            }
	            
	            vecinos[x][y]=-1;
	            
	            //Actualizamos las celdas de alrededor
	            if ((x+1<lado) && (vecinos[x+1][y]!=-1)) vecinos[x+1][y]++;
	            if ((x+1<lado) && (y+1<lado) && (vecinos[x+1][y+1]!=-1)) vecinos[x+1][y+1]++;
	            
	            if ((y+1<lado) && (vecinos[x][y+1]!=-1)) vecinos[x][y+1]++;
	            if ((x-1>=0) && (y+1<lado) && (vecinos[x-1][y+1]!=-1)) vecinos[x-1][y+1]++;
	            
	            if ((x-1>=0) && (vecinos[x-1][y]!=-1)) vecinos[x-1][y]++;
	            if ((x-1>=0) && (y-1>=0) && (vecinos[x-1][y-1]!=-1)) vecinos[x-1][y-1]++;
	            
	            if ((y-1>=0) && (vecinos[x][y-1]!=-1)) vecinos[x][y-1]++;
	            if ((x+1<lado) && (y-1>=0) && (vecinos[x+1][y-1]!=-1)) vecinos[x+1][y-1]++;
	        }
    }
    
    
    private void recolocaBotones()
    {
    	for (int i=0; i<lado; i++)
        {
                for (int j=0; j<lado; j++)
                {
                	posiciones[i][j].setBackground(Color.GREEN);
                	posiciones[i][j].setText("");
                    pano.add(posiciones[i][j]);
                    vecinos[i][j]=0;
                }
        }
    }
    
    public Buscaminas ()
    {
    	pano=new JPanel();
    	barra=new JMenuBar();

    	///////////
    	//TO DO
    	//CREAR LA BARRA DE MENU, MENUS y OPCIONES DE MENU
    	GridLayout grid=new GridLayout(lado,lado);
    	add(pano);
    	pano.setLayout(grid);

    	vecinos=new int[lado][lado];
    	setJMenuBar(barra);
    	juego=new JMenu("Juego");
    	tamanio=new JMenu("Tamanio");
    	barra.add(juego);
    	barra.add(tamanio);
    	opcNuevo=new JMenuItem("Nuevo");
    	opcSalir=new JMenuItem("Salir");
    	juego.add(opcNuevo);
    	juego.add(opcSalir);
    	opc1=new JMenuItem("5x5");
    	opc2=new JMenuItem("10x10");
    	opc3=new JMenuItem("20x20");
    	tamanio.add(opc1);
    	tamanio.add(opc2);
    	tamanio.add(opc3);
    	posiciones=new JButton[20][20];

        //Opcion de menu de salir
    	opcSalir.addActionListener(new ActionListener(){
			@Override
			public void actionPerformed(ActionEvent arg0) {
				System.exit(0);
			}
    	});
    	
    	//5x5
    	opc1.addActionListener(new ActionListener(){
			@Override
			public void actionPerformed(ActionEvent arg0) {
				
			    //Quitamos los botones del panel
		        for (int i=0; i<lado; i++)
		        {
		                for (int j=0; j<lado; j++)
		                {
		                    pano.remove(posiciones[i][j]);
		                    
		                }
		        }
				
		        //Cambiamos de tamaño del tablero
		        lado=5;
		        
		        //Ya se puede hacer clic en los botones:
				gameover=false;
				
				////////////////
				//TO DO
				//Creamos un Layout del tamaño adecuado
				pano.setLayout(new GridLayout(lado,lado));
				vecinos=new int[lado][lado];
				
				
				
				//Añadimos los botones
				recolocaBotones();
				
				//Recalculamos las minas y los vecinos
				calculaMinasVecinos();
				
		        setBounds(0,0,300,300);
		        revalidate();
				repaint();
			}
    	});

    	//10x10
    	opc2.addActionListener(new ActionListener(){
			@Override
			public void actionPerformed(ActionEvent arg0) {
				 	for (int i=0; i<lado; i++)
			        {
			                for (int j=0; j<lado; j++)
			                {
			                    pano.remove(posiciones[i][j]);
			                }
			        }
					
				 	lado=10;
					gameover=false;
					
					///////////////
					//TO DO
					//Creamos un Layout del tamaño adecuado
					//
					//
					pano.setLayout(new GridLayout(lado,lado));
					vecinos=new int[lado][lado];
					
					recolocaBotones();
					
					calculaMinasVecinos();
					
					setBounds(0,0,500,500);
			        revalidate();
				repaint();
			}
    	});

    	//20x20
    	opc3.addActionListener(new ActionListener(){
			@Override
			public void actionPerformed(ActionEvent arg0) {
				 for (int i=0; i<lado; i++)
			        {
			                for (int j=0; j<lado; j++)
			                {
			                    pano.remove(posiciones[i][j]);
			                }
			        }
				 
					lado=20;
					gameover=false;
					
					////////////////
					//TO DO
					//Creamos un Layout del tamaño adecuado
					pano.setLayout(new GridLayout(lado,lado));
					vecinos=new int[lado][lado];
					
					recolocaBotones();
				
					calculaMinasVecinos();
					
			        setBounds(0,0,850,850);
			        revalidate();
					
				repaint();
			}
    	});

    	//Nueva partida
    	opcNuevo.addActionListener(new ActionListener(){
			@Override
			public void actionPerformed(ActionEvent arg0) {
		       for (int i=0; i<lado; i++)
		           for (int j=0; j<lado; j++)
		           {
		               vecinos[i][j]=0;
		               posiciones[i][j].setBackground(Color.GREEN);
		               posiciones[i][j].setText("");
		           }
		       
		       gameover=false;
		       
		       calculaMinasVecinos();
			}
    	});

        
        for (int i=0; i<lado; i++)
            for (int j=0; j<lado; j++)
            {
                vecinos[i][j]=0;
            }
        
        calculaMinasVecinos();
        
        EscuchadorBotones EB=new EscuchadorBotones();
        
        for (int i=0; i<20; i++)
        {
                for (int j=0; j<20; j++)
                {
                    posiciones[i][j]=new JButton();
                    posiciones[i][j].setBackground(Color.GREEN);
                    posiciones[i][j].setActionCommand(new Integer((i*100)+j).toString());
                    posiciones[i][j].addActionListener(EB);
                }
        }
        
        //Colgamos los botones
        for (int i=0; i<lado; i++)
        {
                for (int j=0; j<lado; j++)
                {
                    pano.add(posiciones[i][j]);
                }
        }
            
        getContentPane().add(pano);
    }

    private boolean gameWon()
    {
    	if (((lado*lado) - casillasClicadas)==numMinas)
		{
			return true;
		}
		return false;
    }
    
    class EscuchadorBotones implements ActionListener
    {
        @Override
        public void actionPerformed(ActionEvent arg0) {
        	if (!gameover){
            JButton b=(JButton) arg0.getSource();
            String s=arg0.getActionCommand();
            int i=new Integer(s).intValue()/100;
            int j=new Integer(s).intValue()%100;
            int numVecinos=vecinos[i][j];
            casillasClicadas++;
            
            if (b.getBackground()==Color.GREEN)
            {
                b.setBackground(Color.gray);
                
                if (numVecinos>0)
                {
                    b.setText(new Integer(numVecinos).toString());
                }
                else if (numVecinos==0)
                {
                	//Si se ha pulsado en 'mar abierto', se recorren las 8 casillas de alrededor
                	//poniendo las casillas en gris.
                	
                	//Casilla 1
                	if (i+1<lado) clicaCasilla(i+1,j);
                	
                	//Casilla 2
                	if ((i+1<lado) && (j+1<lado)) clicaCasilla(i+1,j+1);
                    
                	//Casilla 3
                	if ((j+1<lado)) clicaCasilla(i,j+1);    
                	
                	//Casilla 4
                	if ((i-1>0) && (j+1<lado)) clicaCasilla(i-1,j+1);
                	
                	if ((i-1>0)) clicaCasilla(i-1,j);
                	
                	if ((i-1>0) && (j-1>0)) clicaCasilla(i-1,j-1);
                	
                	if ((j-1>0)) clicaCasilla(i,j-1);
                	
                	if ((i+1<lado) && (j-1>0)) clicaCasilla(i+1,j-1);
                }
                else 
                	if (numVecinos==-1)
                	{
                		posiciones[i][j].setBackground(new Color(255,20,20));
                		JOptionPane.showMessageDialog(null, "¡¡¡¡¡BOOOOM!!!!!");
                		for (int ii=0; ii<lado; ii++)
                		{
                               for (int jj=0; jj<lado; jj++)
                               {
                               //    posiciones[ii][jj].setBackground(new Color(255,200,200));
                             //      posiciones[ii][jj].setText("");
                               }
                		}
                		gameover=true;
                      //  opcNuevo.doClick();	
                	}
            	}
            	else
            		//Hasta aqui era si el boton pulsado era verde, pero si es gris...
            		if (b.getBackground()==Color.GRAY)
            		{
            			numVecinos=vecinos[i][j];
            			if (numVecinos==0)
            			{
            				//Casilla 1
                        	if (i+1<lado) clicaCasilla(i+1,j);
                        	
                        	//Casilla 2
                        	if ((i+1<lado) && (j+1<lado)) clicaCasilla(i+1,j+1);
                            
                        	//Casilla 3
                        	if ((j+1<lado)) clicaCasilla(i,j+1);    
                        	
                        	//Casilla 4
                        	if ((i-1>0) && (j+1<lado)) clicaCasilla(i-1,j+1);
                        	
                        	if ((i-1>0)) clicaCasilla(i-1,j);
                        	
                        	if ((i-1>0) && (j-1>0)) clicaCasilla(i-1,j-1);
                        	
                        	if ((j-1>0)) clicaCasilla(i,j-1);
                        	
                        	if ((i+1<lado) && (j-1>0)) clicaCasilla(i+1,j-1);
            			}
            		}
        	}
        }
    }
    
    public static void main(String args[])
    {
        Buscaminas busca=new Buscaminas();
        busca.setVisible(true);
        busca.setBounds(0,0,500,500);
    }
}