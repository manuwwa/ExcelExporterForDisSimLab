package excelExporter;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.util.LinkedList;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import dissimlab.monitors.Change;
import dissimlab.monitors.ChangesList;

import org.jxls.common.Context;
import org.jxls.transform.Transformer;
import org.jxls.util.JxlsHelper;
import org.jxls.util.TransformerFactory;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;



import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.util.Map;
import java.util.HashMap;
import java.io.FileOutputStream;
/**
 * Main class for package excelExpoer.
 * <p>
 * Class allows to export monitored values by using static method exportToCSV. 
 */
public class Exporter {
	/**
	 * Creates exporter instance and runs file chooser.
	 * <p>
	 * Class allows to create exporter instance. 
	 */
	private Exporter()
	{
		chozeDirectorySwing();
		if(directory.substring(directory.length() - 1)!="\\")
		{
			directory=directory+"\\";
		}
		
		

		if(test)System.out.println(this.directory);
	}
	/**
	 * Internal method used to make user friendly chooser using java swing
	 * <p>
	 * This method always to run write to CSV file
	 */
	private void chozeDirectorySwing()
	{
		JFileChooser chooser = new JFileChooser();
	    chooser.setCurrentDirectory(new java.io.File("."));
	    chooser.setDialogTitle("Choose directory to save monitored values");
	    chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
	    chooser.setAcceptAllFileFilterUsed(false);

	    if (chooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
	      //System.out.println("getCurrentDirectory(): " + chooser.getCurrentDirectory());
	      //System.out.println("getSelectedFile() : " + chooser.getSelectedFile());
	    	this.directory=chooser.getSelectedFile().toString();
	    } else {
	      System.out.println("No Selection ");
	    }
	}
	private String directory;
	private static boolean test=true;
	/**
	 * Internal method used do write to CSV file
	 * <p>
	 * This method always to run write to CSV file
	 * @throws IOException if program cna't create directory for export or if program have problem with csv file
	 */
	private void exportCSV() throws IOException
	{
		LinkedList<MonitoredVarWithExport> list= MonitoredVarWithExport.ListOfMonitored;
		BufferedWriter writer = null;
		boolean isMoreThreads=createDirForThread();
		for (MonitoredVarWithExport ex : list) {
			try {
				writer = new BufferedWriter(new FileWriter(directory+ex.threadDirectory(isMoreThreads)+ex.getName()+".csv"));
				
			writer.write("Time;value\n");
				ChangesList clist= ex.getChanges();
				int N=clist.size();
				if(test)System.out.println(N);
				for(int i=0;i<N;i++)
				{
					Change	c= clist.get(i);
					writer.write(c.getTime()+";"+c.getValue()+"\n");
				}
			} catch (IOException e) {
				throw e;
			}
			finally 
			{
	            try {
	                
	                writer.close();
	            } catch (Exception e) {
	            }
	        }
		}
		
	}


	private static Map<String, CellStyle> createStyles(Workbook wb){
        Map<String, CellStyle> styles = new HashMap<String, CellStyle>();
        CellStyle style;
        Font titleFont = wb.createFont();
        titleFont.setFontHeightInPoints((short)18);
        titleFont.setBold(true);
        style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFont(titleFont);
        styles.put("title", style);

        Font monthFont = wb.createFont();
        monthFont.setFontHeightInPoints((short)11);
        monthFont.setColor(IndexedColors.WHITE.getIndex());
        style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFont(monthFont);
        style.setWrapText(true);
        
        styles.put("header", style);

        style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        styles.put("cell", style);
        
        
        style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        styles.put("cell2", style);
        
        
        
        style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setDataFormat(wb.createDataFormat().getFormat("0.00"));
        styles.put("formula", style);

        style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFillForegroundColor(IndexedColors.GREY_40_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setDataFormat(wb.createDataFormat().getFormat("0.00"));
        styles.put("formula_2", style);

        return styles;
    }
	
	/**
	 * Internal method used do write to CSV file
	 * <p>
	 * This method always to run write to CSV file
	 * @throws Exception 
	 */
	
	private void exportXLSX() throws Exception
	{
		
		WorkbookForThreadList wb=new WorkbookForThreadList();
		Map<String, CellStyle> styles=createStyles(wb.getFirst().workbook);
		for (WorkbookForThread w:wb)
		{
			 styles = createStyles(w.workbook);
		}
		
		LinkedList<MonitoredVarWithExport> list= MonitoredVarWithExport.ListOfMonitored;
		
		for (MonitoredVarWithExport ex : list) 
		{
			int rownum=0;
			Sheet sheet = wb.findWorkbookByThread(ex.getAppThread().getId()).workbook.createSheet(ex.getName());
			
			
			
			// Ustawienia strony druku
	        PrintSetup printSetup = sheet.getPrintSetup();
	        printSetup.setLandscape(true);
	        
	        sheet.setHorizontallyCenter(true);
	        //Title
	        Row titleRow = sheet.createRow(rownum++);
	        titleRow.setHeightInPoints(45);
	        Cell titleCell = titleRow.createCell(0);
	        titleCell.setCellValue("Row data");
	        titleCell.setCellStyle(styles.get("title"));
	        sheet.addMergedRegion(CellRangeAddress.valueOf("$A$1:$B$1"));
	        //hedders
	        Row headerRow = sheet.createRow(rownum++);
	        headerRow.setHeightInPoints(25);
	        Cell headerCell;
	        
	            headerCell = headerRow.createCell(0);
	            headerCell.setCellStyle(styles.get("header"));
	            headerCell.setCellValue("Time");
	            
	            headerCell = headerRow.createCell(1);
	            headerCell.setCellStyle(styles.get("header"));
	            headerCell.setCellValue("Value");
	            
	            
	            

	        
	        ChangesList changes=ex.getChanges();
	        for(int i=0; i<changes.size();i++)
	        {
	        	Row row = sheet.createRow(rownum++);
	        		
	        	
	        		Cell cell = row.createCell(0);
	        		if(i%2==1) cell.setCellStyle(styles.get("cell"));
	        		else cell.setCellStyle(styles.get("cell2"));
	        		
	        		cell.setCellValue(changes.get(i).getTime());
	        		cell = row.createCell(1);
	        		 if(i%2==1) cell.setCellStyle(styles.get("cell"));
	        		 else cell.setCellStyle(styles.get("cell2"));
	        		cell.setCellValue(changes.get(i).getValue());
	        		
	        	
	        	
	        }
	        //add statistic
	        int created=CreateStatistic(sheet,styles);
	      //cell sizes set
	        //For statistic
	        for(int indexx=0;indexx<created;indexx++)
	        {
	        	sheet.setColumnWidth(3+indexx,4500);
	        }
	        
	        
	        
		}
		for (WorkbookForThread w:wb)
		{
			String file = directory+"thread"+w.thread+".xlsx";
			FileOutputStream out = new FileOutputStream(file);
			w.workbook.write(out);
			out.close();
			w.workbook.close();
		}

	}
	private int CreateStatistic(Sheet sheet,Map<String, CellStyle> styles)
	{
		int place=3;
		int Created=0;
		Row headerRowt;
		Row Row;
		int LastRow=sheet.getLastRowNum();
		LinkedList<Formulas> formulas =new LinkedList<Formulas>();
		formulas.add(new Formulas("AVERAGE(B3:B"+(LastRow+1)+")", "Arithmetic mean"));
		formulas.add(new Formulas("HARMEAN(B3:B"+(LastRow+1)+")", "Harmonic mean"));
		formulas.add(new Formulas("max(B3:B"+(LastRow+1)+")-min(B3:B"+(LastRow+1)+")", "Value Range"));
		formulas.add(new Formulas("VAR(B3:B"+(LastRow+1)+")", "Variance"));
		formulas.add(new Formulas("STDEVP(B3:B"+(LastRow+1)+")", "Standard deviation"));
		formulas.add(new Formulas("min(B3:B"+(LastRow+1)+")", "Minimal value"));
		formulas.add(new Formulas("max(B3:B"+(LastRow+1)+")", "Maximal value"));
		formulas.add(new Formulas("COUNT(B3:B"+(LastRow+1)+")", "Number of samples"));
		//formulas.add(new Formulas("CONFIDENCE(B3:B"+(LastRow+1)+")", "Estimated value"));
		
		
		
		headerRowt= sheet.getRow(1);
		Row= sheet.getRow(2);
		for(Formulas formula:formulas)
		{
			
	        Cell cell = headerRowt.createCell(place);
	        
	        cell.setCellStyle(styles.get("header"));
	        cell.setCellValue(formula.headerName);
	        
	        
	        cell = Row.createCell(place);
	        cell.setCellFormula(formula.formula);
	        cell.setCellStyle(styles.get("cell"));
	        
	        place++;
	        Created++;
		}
		
        return Created;
        
	}
	 private class Formulas
	 {
		 Formulas(String formula, String headerName)
		 {
			 this.formula=formula;
			 this.headerName=headerName;
		 }
		 String formula;
		 String headerName;
	 }
	/**
	 * Exports monitored values to CSV file using java swing (simple user interface)
	 * <p>
	 * This method always to run export procedure
	 * @throws IOException if program cna't create directory for export or if program have problem with csv file
	 */
	public static void exportToCSV() throws IOException
	{
		Exporter ex=new Exporter();
		ex.exportCSV();
	}
	public static void exportToXLS() throws IOException
	{
		Exporter ex=new Exporter();
		try
		{
			ex.exportXLSX();
		}
		catch(IOException e)
		{
			throw e;
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	private LinkedList<Long> getThreads()
	{
LinkedList<Long> treads= new LinkedList<Long>();
		
		for (MonitoredVarWithExport ex : MonitoredVarWithExport.ListOfMonitored)
		{
			if(!treads.contains(ex.getAppThread().getId())) 
				{
				treads.addFirst(ex.getAppThread().getId());
					
				}
		}
		return treads;
	}

	/**
	 * Creates If needed directory for simulations threads.
	 * <p>
	 * This method always to prepare export task for multi-thread simulations
	 *
	 * @return if simulation is multi-thread 
	 * @throws IOException if program cna't create directory for export
	 */
	private boolean createDirForThread() throws IOException
	{

		LinkedList<Long> treads= new LinkedList<Long>();
		
		for (MonitoredVarWithExport ex : MonitoredVarWithExport.ListOfMonitored)
		{
			if(!treads.contains(ex.getAppThread().getId())) 
				{
				treads.addFirst(ex.getAppThread().getId());
					
				}
		}
		if(treads.size()>1)
		{
			for(Long x : treads)
			{
				boolean success = (new File(directory+"//"+x)).mkdirs();
				if(!success)
				{
					throw new IOException("Simulation encountered a problem when exporting. The program can not create a directory : "+directory+"//"+x+" needed to export. Please check exporter class");

				}
			}
		}
		return treads.size()>1;

	}
	
}
