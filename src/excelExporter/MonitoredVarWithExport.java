package excelExporter;


import java.util.LinkedList;

import dissimlab.monitors.*;
/**
 * Class stores monitored variables.
 * <p>
 * Class allows to export monitored variables to CSV files. Class is based on MonitoredVar. 
 */
public class MonitoredVarWithExport extends MonitoredVar  {
	private String name="unset";
	private Thread appThread;
	public static LinkedList<MonitoredVarWithExport>ListOfMonitored=new LinkedList<MonitoredVarWithExport>();
	/**
	 * Creates MonitoredVar 
	 * <p>
	 * This method always to create MonitoredVar and set variable value
	 *
	 * @param  Monitor name used as exported file name
	 * @param  
	 */
	public MonitoredVarWithExport(String name,double value)
	{
		super(value);
		appThread=Thread.currentThread();
		this.setName(name);
		MonitoredVarWithExport.ListOfMonitored.add(this);
	}
	/**
	 * Creates MonitoredVar 
	 * <p>
	 * This method always to create MonitoredVar
	 *
	 * @param  Monitor name used as exported file name
	 */
	public MonitoredVarWithExport(String name)
	{
		super();
		appThread=Thread.currentThread();
		this.setName(name);
		MonitoredVarWithExport.ListOfMonitored.add(this);
	}
	/**
	 * Returns name of the monitor
	 * <p>
	 * This method always to get monitor name 
	 *
	 * @return Name of the monitor
	 */
	public String getName() {
		return name;
	}
	/**
	 * Sets name for the monitor 
	 * <p>
	 * This method always to set monitor name 
	 *
	 * @param  Monitor name used as exported file name
	 */
	private void setName(String name) {
		this.name = name;
	}
	public Thread getAppThread() {
		return appThread;
	}
	public String threadDirectory(boolean isMoreThread)
	{
		if(isMoreThread)
		{
			return this.getAppThread()+"\\";
		}
		else
		{
			return "";
		}
	}
}
