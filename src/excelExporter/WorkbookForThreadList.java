package excelExporter;

import java.util.LinkedList;
public class WorkbookForThreadList extends LinkedList<WorkbookForThread> {
	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	public WorkbookForThread findWorkbookByThread(Long thread) throws Exception
	{
		for(WorkbookForThread sh:this)
		{
			if(sh.thread==thread) return sh;
		}
		throw new Exception("thread not found");
	}
	public WorkbookForThreadList() {
		LinkedList<Long> treads=getThreads();
		for(Long i:treads)
		{
			this.add(new WorkbookForThread(i));
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
}
