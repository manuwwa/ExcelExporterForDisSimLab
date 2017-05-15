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
    private static final String[] titles = {
            "Person",	"ID", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun",
            "Total\nHrs", "Overtime\nHrs", "Regular\nHrs"
    };

    private static Object[][] sample_data = {
            {"Yegor Kozlov", "YK", 5.0, 8.0, 10.0, 5.0, 5.0, 7.0, 6.0},
            {"Gisella Bronzetti", "GB", 4.0, 3.0, 1.0, 3.5, null, null, 4.0},
    };
	public static void test() throws IOException
	{
		Workbook wb;
		wb = new XSSFWorkbook();
		Map<String, CellStyle> styles = createStyles(wb);
		Sheet sheet = wb.createSheet("Timesheet");
        PrintSetup printSetup = sheet.getPrintSetup();
        printSetup.setLandscape(true);
        sheet.setFitToPage(true);
        sheet.setHorizontallyCenter(true);

        //title row
        Row titleRow = sheet.createRow(0);
        titleRow.setHeightInPoints(45);
        Cell titleCell = titleRow.createCell(0);
        titleCell.setCellValue("Weekly Timesheet");
        titleCell.setCellStyle(styles.get("title"));
        sheet.addMergedRegion(CellRangeAddress.valueOf("$A$1:$L$1"));

        //header row
        Row headerRow = sheet.createRow(1);
        headerRow.setHeightInPoints(40);
        Cell headerCell;
        for (int i = 0; i < titles.length; i++) {
            headerCell = headerRow.createCell(i);
            headerCell.setCellValue(titles[i]);
            headerCell.setCellStyle(styles.get("header"));
        }

        int rownum = 2;
        for (int i = 0; i < 10; i++) {
            Row row = sheet.createRow(rownum++);
            for (int j = 0; j < titles.length; j++) {
                Cell cell = row.createCell(j);
                if(j == 9){
                    //the 10th cell contains sum over week days, e.g. SUM(C3:I3)
                    String ref = "C" +rownum+ ":I" + rownum;
                    cell.setCellFormula("SUM("+ref+")");
                    cell.setCellStyle(styles.get("formula"));
                } else if (j == 11){
                    cell.setCellFormula("J" +rownum+ "-K" + rownum);
                    cell.setCellStyle(styles.get("formula"));
                } else {
                    cell.setCellStyle(styles.get("cell"));
                }
            }
        }

        //row with totals below
        Row sumRow = sheet.createRow(rownum++);
        sumRow.setHeightInPoints(35);
        Cell cell;
        cell = sumRow.createCell(0);
        cell.setCellStyle(styles.get("formula"));
        cell = sumRow.createCell(1);
        cell.setCellValue("Total Hrs:");
        cell.setCellStyle(styles.get("formula"));

        for (int j = 2; j < 12; j++) {
            cell = sumRow.createCell(j);
            String ref = (char)('A' + j) + "3:" + (char)('A' + j) + "12";
            cell.setCellFormula("SUM(" + ref + ")");
            if(j >= 9) cell.setCellStyle(styles.get("formula_2"));
            else cell.setCellStyle(styles.get("formula"));
        }
        rownum++;
        sumRow = sheet.createRow(rownum++);
        sumRow.setHeightInPoints(25);
        cell = sumRow.createCell(0);
        cell.setCellValue("Total Regular Hours");
        cell.setCellStyle(styles.get("formula"));
        cell = sumRow.createCell(1);
        cell.setCellFormula("L13");
        cell.setCellStyle(styles.get("formula_2"));
        sumRow = sheet.createRow(rownum++);
        sumRow.setHeightInPoints(25);
        cell = sumRow.createCell(0);
        cell.setCellValue("Total Overtime Hours");
        cell.setCellStyle(styles.get("formula"));
        cell = sumRow.createCell(1);
        cell.setCellFormula("K13");
        cell.setCellStyle(styles.get("formula_2"));

        //set sample data
        for (int i = 0; i < sample_data.length; i++) {
            Row row = sheet.getRow(2 + i);
            for (int j = 0; j < sample_data[i].length; j++) {
                if(sample_data[i][j] == null) continue;

                if(sample_data[i][j] instanceof String) {
                    row.getCell(j).setCellValue((String)sample_data[i][j]);
                } else {
                    row.getCell(j).setCellValue((Double)sample_data[i][j]);
                }
            }
        }

        //finally set column widths, the width is measured in units of 1/256th of a character width
        sheet.setColumnWidth(0, 30*256); //30 characters wide
        for (int i = 2; i < 9; i++) {
            sheet.setColumnWidth(i, 6*256);  //6 characters wide
        }
        sheet.setColumnWidth(10, 10*256); //10 characters wide

        // Write the output to a file
        String file = "timesheet.xls";
        if(wb instanceof XSSFWorkbook) file += "x";
        FileOutputStream out = new FileOutputStream(file);
        wb.write(out);
        out.close();
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
	 * @throws IOException if program cna't create directory for export or if program have problem with csv file
	 */
	
	private void exportXLSX() throws IOException
	{
		
		Workbook wb;
		wb = new XSSFWorkbook();
		Map<String, CellStyle> styles = createStyles(wb);
		LinkedList<MonitoredVarWithExport> list= MonitoredVarWithExport.ListOfMonitored;
		for (MonitoredVarWithExport ex : list) 
		{

			Sheet sheet = wb.createSheet(ex.getName());
			// Ustawienia strony druku
	        PrintSetup printSetup = sheet.getPrintSetup();
	        printSetup.setLandscape(true);
	        
	        sheet.setHorizontallyCenter(true);

	        //hedders
	        Row headerRow = sheet.createRow(0);
	        headerRow.setHeightInPoints(25);
	        Cell headerCell;
	        
	            headerCell = headerRow.createCell(0);
	            headerCell.setCellStyle(styles.get("header"));
	            headerCell.setCellValue("Time");
	            
	            headerCell = headerRow.createCell(1);
	            headerCell.setCellStyle(styles.get("header"));
	            headerCell.setCellValue("Value");

	        int rownum=1;
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
			
		}
		String file = "elo.xlsx";
		FileOutputStream out = new FileOutputStream(file);
		wb.write(out);
		out.close();
		wb.close();
	}
	private void exportXLSXold() throws IOException
	{
		LinkedList<MonitoredVarWithExport> list= MonitoredVarWithExport.ListOfMonitored;
		BufferedWriter writer = null;
		boolean isMoreThreads=createDirForThread();

		for (MonitoredVarWithExport ex : list) 
		{
			
//			try 
//			{

				
				List<Change> changes =CastChangesListToList(ex.getChanges());
				for(Change i : changes)
				{
					System.out.println("("+i.getTime()+":"+i.getValue()+")");
				}
				//throw new IOException();
				InputStream is = Exporter.class.getResourceAsStream("Q:\\home\\OneDrive\\Magyster\\msk\\DisSimLab2017\\resources\\excelExporter\\export_template.xls");
				OutputStream os = new FileOutputStream("C:\\temp\\wyjscie.xls");
//				OutputStream os = new FileOutputStream(directory+"wyjscie.xls");
				Context context = new Context();
				context.putVar("changes", changes);
				
				JxlsHelper jxlsHelper = JxlsHelper.getInstance();
				jxlsHelper.setUseFastFormulaProcessor(false);
				System.out.println("byc");
				jxlsHelper.processTemplate(is, os, context);
			//	JxlsHelper.getInstance().processTemplateAtCell(is, os, context, "Result!A1");
//			}
//			catch (IOException e) {
//			
//			throw e;
//		}
//		finally 
//		{
//            try 
//            	{
//
//                
//            	}
//            catch (Exception e) 
//            	{
//            	}
//		}
		}
		
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
		}
	}
	private static ExcelDataModel createDataModel(ChangesList Clist)
	{
		ExcelDataModel ret= new ExcelDataModel("Time;Value");
		for(int i=0; i<Clist.size();i++)
		{
			String time= ""+
			Clist.get(i).getTime();
			String value=""+
			Clist.get(i).getValue();
			String[] s= {time,value};
			try {
				ret.AddRow(s);
			} catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
		
		return ret;
	}
	//#old
	private static List<Change> CastChangesListToList(ChangesList Clist)
	{
		List<Change> list = new LinkedList<Change>();
		for(int i=0; i<Clist.size();i++)
		{
			list.add(Clist.get(i));
		}
		return list;
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
