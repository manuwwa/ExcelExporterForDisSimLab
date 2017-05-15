package excelExporter;
import java.util.LinkedList;
import java.util.List;
import dissimlab.monitors.*;
public class ChangesListWithexporterToList extends ChangesList {
	
	public List<Change> toList()
	{
		List<Change> list = new LinkedList<Change>();
		for(int i=0; i<super.size();i++)
		{
			list.add(super.get(i));
		}
		return list;
	}
	@Override
	public String toString() {
		String ret="";
		for(int i=0; i<super.size();i++)
		{
			ret=ret+"("+super.get(i).getTime()+","+super.get(i).getValue()+"),";
		}
		if (super.size()>0)
		ret=ret.substring(0, ret.length()-1);
		return ret;
	}
}
