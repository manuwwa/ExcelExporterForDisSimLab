package excelExporter;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.KeyEvent;
import java.awt.event.WindowEvent;
import java.io.IOException;

import javax.swing.AbstractButton;
import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.WindowConstants;

public class ExportChoisePannel  extends JPanel implements ActionListener{

	protected JButton ToCSVButton, ToXLSButton, DoNotExportButoon;

	public ExportChoisePannel() {
		ToCSVButton = new JButton("Export to CSV");
		ToCSVButton.setVerticalTextPosition(AbstractButton.CENTER);
		ToCSVButton.setHorizontalTextPosition(AbstractButton.LEADING); //aka LEFT, for left-to-right locales
		ToCSVButton.setMnemonic(KeyEvent.VK_D);
		ToCSVButton.setActionCommand("CSV");
		ToXLSButton = new JButton("Export to XLS");
		ToXLSButton.setVerticalTextPosition(AbstractButton.BOTTOM);
		ToXLSButton.setHorizontalTextPosition(AbstractButton.CENTER);
		ToXLSButton.setMnemonic(KeyEvent.VK_M);
		ToXLSButton.setActionCommand("XLS");
		DoNotExportButoon = new JButton("Do not export");
        //Use the default text position of CENTER, TRAILING (RIGHT).
		DoNotExportButoon.setMnemonic(KeyEvent.VK_E);
		DoNotExportButoon.setActionCommand("cancel");
		
		
		ToCSVButton.addActionListener(this);
		ToXLSButton.addActionListener(this);
		DoNotExportButoon.addActionListener(this);
		
		ToCSVButton.setToolTipText("Click this button to Export monitored var to csv");
		ToXLSButton.setToolTipText("Click this button to Export monitored var to xls");
		DoNotExportButoon.setToolTipText("Click this button if you don't need to export monitored var");
		
        add(ToCSVButton);
        add(ToXLSButton);
        add(DoNotExportButoon);
		
		
	}
    public static void createAndShowGUI() {
        //Create and set up the window.
        JFrame frame = new JFrame("Czy chcesz wyeksportowac monitorowane wartosci");
        frame.setDefaultCloseOperation(WindowConstants.HIDE_ON_CLOSE);
        //Create and set up the content pane.
        ExportChoisePannel newContentPane = new ExportChoisePannel();
        newContentPane.setOpaque(true); //content panes must be opaque
        frame.setContentPane(newContentPane);
 
        //Display the window.
        frame.pack();
        frame.setVisible(true);
    }
	@Override
	public void actionPerformed(ActionEvent e) {
		if ("CSV".equals(e.getActionCommand()))
		{
			try {
				Exporter.exportToCSV();
			} catch (IOException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}
		}
		if ("XLS".equals(e.getActionCommand()))
		{
			try {
				Exporter.exportToXLS();
			} catch (IOException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}
			
		}
		if ("Export to XLS".equals(e.getActionCommand()))
		{
			
		}
		
	}
	

}
