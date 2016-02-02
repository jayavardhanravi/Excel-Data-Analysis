/**
 * 
 */
package maid.mycontent.orig;

import java.awt.Button;
import java.awt.Color;
import java.awt.Dimension;
import java.awt.FlowLayout;
import java.awt.Font;
import java.awt.Frame;
import java.awt.GridLayout;
import java.awt.Label;
import java.awt.Panel;
import java.awt.Rectangle;
import java.awt.TextField;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;

import javax.swing.JProgressBar;

/**
 * @author
 *
 */
public class ExcelModify {
	
	private Frame mainFrame;
	   private Label headerLabel;
	   private Label statusLabel;
	   private Panel controlPanel;
	   private Panel controlPanel2;
	   private Label footerLabel;
	   private Font plainFont = new Font("Serif", Font.PLAIN, 16);
	   private static JProgressBar barDo;


	   public ExcelModify() {
	      prepareGUI();
	   }

	   public static void main(String[] args){
		   System.out.println("Initialized the user interface.");
		   ExcelModify  awtControlDemo = new ExcelModify();
	      awtControlDemo.showTextFieldDemo();
	        
	   }

	   private void prepareGUI(){
	      mainFrame = new Frame("MAID - Excel Data Modification Tool");
	      mainFrame.setSize(700,600);
	      Rectangle bounds = new Rectangle(700, 600);
		mainFrame.setMaximizedBounds(bounds);
		Dimension minimumSize = new Dimension(700, 600);
		mainFrame.setMinimumSize(minimumSize );
	      mainFrame.setLayout(new GridLayout(6, 1));
	      mainFrame.addWindowListener(new WindowAdapter() {
	         public void windowClosing(WindowEvent windowEvent){
	            System.exit(0);
	         }        
	      });    
	      headerLabel = new Label();
	      headerLabel.setAlignment(Label.CENTER);
	      headerLabel.setFont(plainFont);
	      footerLabel = new Label();
	      footerLabel.setAlignment(Label.CENTER);
	      statusLabel = new Label();        
	      statusLabel.setAlignment(Label.CENTER);
	      String intro = "Format the excel files correctly before you generate the report. Place all the files in the same directory as ExcelMod.jar file.";
	      statusLabel.setText(intro);
	      barDo = new JProgressBar(0, 100);
	      barDo.setBackground(Color.WHITE);
	      barDo.setBorderPainted(false);
	      controlPanel = new Panel();
	      controlPanel.setLayout(new GridLayout(3, 3,0,8));
	      controlPanel2 = new Panel();
	      controlPanel2.setLayout(new GridLayout(3,1));
	      controlPanel2.add(barDo);
	      mainFrame.add(headerLabel);
	      mainFrame.add(controlPanel);
	      mainFrame.add(statusLabel);
	      mainFrame.add(controlPanel2);
	      mainFrame.add(new Label());
	      mainFrame.add(footerLabel);
	      mainFrame.setVisible(true); 
	      OutputConsole console = new OutputConsole();
	   }

	   private void showTextFieldDemo(){
	      headerLabel.setText(" ******* Make sure the file is not in use / properly formatted*******"); 
	      headerLabel.setForeground(Color.RED);
	      Label  namelabel= new Label("Excel Main File Path: ", Label.RIGHT);
	      final TextField userText = new TextField(60);
	      Label  namelabel1= new Label("Excel Data File Path: ", Label.RIGHT);
	      final TextField userText1 = new TextField(60);
	      System.out.println("Generated the user interface - [Success]");
	      Button modifyButton = new Button("Generate");	      
	      modifyButton.addActionListener(new ActionListener() {
	         public void actionPerformed(ActionEvent e) {   	

		        new Thread(new thread1()).start();
	            String data = "Analysing the file :      **** WAIT..... ****";
	            statusLabel.setText(data);
	            statusLabel.setForeground(Color.BLUE);
	            ExcelModStart obj = new ExcelModStart();
	            data = obj.ExcelMethodStart(userText.getText(),userText1.getText());
	            statusLabel.setText(data);
	         }
	      }); 
	      
	      Button clearDb = new Button("Clean Database");
	      clearDb.addActionListener(new ActionListener() {
		         public void actionPerformed(ActionEvent e) {   	

			        new Thread(new thread1()).start();
		            String data = "Cleaning the Database :       **** WAIT..... ****";
		            statusLabel.setText(data);
		            statusLabel.setForeground(Color.BLUE);
		            ExcelModStart obj = new ExcelModStart();
		            data = obj.cleanDataBase();
		            statusLabel.setText(data);
		         }
		      }); 
	      
	      footerLabel.setText("............... Custom made Application...................");
	      
	      controlPanel.add(namelabel);
	      controlPanel.add(userText);
	      controlPanel.add(new Label());
	      controlPanel.add(namelabel1);
	      controlPanel.add(userText1);
	      controlPanel.add(clearDb);
	      controlPanel.add(new Label());
	      controlPanel.add(modifyButton);
	      controlPanel.add(new Label());
	      mainFrame.setVisible(true);  
	      
	   }
	 
		public static class thread1 implements Runnable{
			public void run(){
				for (int i=0; i<=100; i++){ 
					barDo.setValue(i); 
					barDo.setStringPainted(true);
					barDo.setString(String.valueOf(i)+"%");
					barDo.repaint();
					try{Thread.sleep(200);} 
					catch (InterruptedException err){}
				}
				barDo.setString("Process Completed !!!!");
			}
		}
	}

