import java.awt.BorderLayout;
import java.awt.GridLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

import javax.swing.BorderFactory;
import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JTextField;

public class GUI implements ActionListener {
	private static JPanel panel;
	private static JFrame frame;
	private static JLabel label;
	private static JLabel label2;
	private static JLabel label3;
	private static JButton button;
	private static JLabel running;
	private static JTextField locationEx;
	private static JTextField locationDoc;
	private static JTextField locationPic;
	private static JLabel without;
	{
		 
	
	}
	public static void main(String[] args) {
		//default setup
		 panel = new JPanel();
		 frame = new JFrame(); 
		frame.setSize(825,300);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.add(panel);
		panel.setLayout(null);
		
		//first textfield
		 label = new JLabel("Location of Excel file: ");
		label.setBounds(10,20,160,25);
		panel.add(label);
		locationEx = new JTextField();
		locationEx.setBounds(200,20,500,25);
		panel.add(locationEx);
		
		//label for first 
		without = new JLabel("(without .xlsx)");
		without.setBounds(710,20,100,25);
		panel.add(without);
		
		//second textfield
		label2 = new JLabel("Location of Output Folder:");
		label2.setBounds(10,50,160,25);
		panel.add(label2);
		locationDoc = new JTextField();
		locationDoc.setBounds(200,50,500,25);
		panel.add(locationDoc);
		
		//third textfield
		label3 = new JLabel("Location of Picture folder:");
		label3.setBounds(10,80,160,25);
		panel.add(label3);
		locationPic = new JTextField();
		locationPic.setBounds(200,80,500,25);
		panel.add(locationPic);
		
		//button
		button = new JButton("Enter");
		button.setBounds(600,130,100 , 25);
		panel.add(button);
		
		running = new JLabel("");
		running.setBounds(10, 180, 600, 25);
		panel.add(running);
		button.addActionListener(new GUI());
		
		frame.setVisible(true);
		GUI test = new GUI();
		
	}
	@Override
	// C:\Users\Msi\Downloads\Employee Excel to Word\HR Shwe Palin CV Data
	// C:\Users\Msi\Downloads\Employee Excel to Word
	// C:\Users\Msi\Downloads\Employee Excel to Word\Picture
	public void actionPerformed(ActionEvent arg0) {
		String locEx = locationEx.getText();
		String locDoc = locationDoc.getText();
		String locPic = locationPic.getText();
		
		try {
			running.setText("Running... This may take a minute");
			Main.mainMethod(locEx, locDoc, locPic);
			running.setText("Your document has been created!");
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			running.setText("Oops... something went wrong");
		}
	
	
	}

	
}
