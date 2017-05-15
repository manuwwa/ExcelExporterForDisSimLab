package excelExporter;

import java.util.LinkedList;

class ExcelDataModel {
	String[] Titles;
LinkedList<String[]> Data=new LinkedList<String[]>();
	public ExcelDataModel(String[] Titles)
	{
		this.Titles=Titles;
	}
	public ExcelDataModel(String Titles)
	{
		this.Titles=Titles.split(";");
		
	}
	public ExcelDataModel(String Titles,String separator)
	{
		this.Titles=Titles.split(separator);
		
	}
	public void AddRow(String Row) throws Exception
	{
		AddRow(Row.split(";"));
	}
	public void AddRow(String Row,String separator) throws Exception
	{
		AddRow(Row.split(separator));
	}
	public void AddRow(String[] Row) throws Exception
	{
		if(Row.length!=Titles.length) throw new Exception("Row lenght is not the same as Titles lenght. Row lenght must be:"+Titles.length);
		Data.addLast(Row);
	}
	public LinkedList<String[]> getData() {
		return Data;
	}
//public String[][] getData() {
//	return Data.toArray(new String[][] {});
//}
public String[] getTitles() {
	return Titles;
}
}
