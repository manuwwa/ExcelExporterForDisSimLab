package excelExporter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
class WorkbookForThread {
	public Workbook workbook;
	public Long thread;
	public WorkbookForThread(Long thread)
	{
		this.workbook= new XSSFWorkbook();;
		this.thread=thread;
	}
}
