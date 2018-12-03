/*
  Programmer: Patrick Nercessian

  Class Purpose:
     This class creates reports of orders based on specified aspects of those orders.
 */

package system;

import java.io.*;
import org.apache.poi.xssf.usermodel.*;

import java.nio.file.Files;
import java.nio.file.StandardCopyOption;

import java.util.Date;
import java.text.SimpleDateFormat;

import java.lang.StackTraceElement;

public class OpenReport {
    private File reportFile;
    private XSSFWorkbook wb;
    private XSSFSheet report;
    private XSSFRow row;
    
    private String[][] list;
    private String reportFileName;

    private double[] totals = new double[6];

    public OpenReport(String[][] list, String ehfSalesRep, String reportName) {
	MasterLog.appendEntry("Creating new " + reportName + " for " + ehfSalesRep + "...");
	this.list = list;

	Date date = new Date();
	SimpleDateFormat sdf = new SimpleDateFormat("MM-dd-yyyy");

	try {
	    reportFile = new File("ADMIN/OTHER/maven/src/main/resources/file/" + reportName + "s/" + ehfSalesRep + "/" + sdf.format(date) + " " + reportName + ".xlsm");
	    if (!reportFile.exists())
		reportFile.mkdirs();

	    //these lines are needed because it makes a directory otherwise for some weird reason
	    File template = new File("ADMIN/OTHER/maven/src/main/resources/file/Empty Excel.xlsm");
	    Files.copy(template.toPath(), reportFile.toPath(), StandardCopyOption.REPLACE_EXISTING);

	    
	    wb = new XSSFWorkbook(new FileInputStream(reportFile));
	    wb.setSheetName(0, reportName);
	    report = wb.getSheet(reportName);
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	}//try-catch
    }//OpenReport(String[][], String)

    private double populateOrderType(int orderType) {
	report.disableLocking();
	String[] orderTypeList = list[orderType];
	String currentOrder;
	String sc, client, factory, model, date;
	int startIndex, endIndex;
	double totalValue = 0, value;

	String orderTypeString = "ERROR";
	switch (orderType) {
	case 0: orderTypeString = "PENDING";break;
	case 1: orderTypeString = "CONFIRMED";break;
	case 2: orderTypeString = "BOOKING";break;
	case 3: orderTypeString = "SHIPPED";break;
	case 4: orderTypeString = "CANCELED";break;
	case 5: orderTypeString = "REINSTATED";break;
	}
						
	row.createCell(0).setCellValue(orderTypeString);
	
	for (int i = 0; i < orderTypeList.length && orderTypeList[i] != null; i++) {
	    currentOrder = orderTypeList[i];
	    row = report.createRow(row.getRowNum() + 1); //incrementing row
	    startIndex = currentOrder.indexOf("SC#") + 5;
	    endIndex = currentOrder.indexOf(" - Client:");
	    sc = currentOrder.substring(startIndex, endIndex);

	    startIndex = currentOrder.indexOf("Client:") + 8;
	    endIndex = currentOrder.indexOf(" - Factory:");
	    client = currentOrder.substring(startIndex, endIndex);

	    startIndex = currentOrder.indexOf("Factory:") + 9;
	    endIndex = currentOrder.indexOf(" - Model:");
	    factory = currentOrder.substring(startIndex, endIndex);

	    startIndex = currentOrder.indexOf("Model:") + 7;
	    endIndex = currentOrder.indexOf(" - Updated:");
	    model = currentOrder.substring(startIndex, endIndex);

	    date = currentOrder.substring(0, 10);

	    row.createCell(0).setCellValue(sc);
	    row.createCell(1).setCellValue(client);
	    row.createCell(2).setCellValue(factory);
	    row.createCell(3).setCellValue(model);
	    row.createCell(4).setCellValue(date);

	    //entering value
	    XSSFCellStyle dollarStyle = wb.createCellStyle();
	    XSSFDataFormat df = wb.createDataFormat();
	    dollarStyle.setDataFormat(df.getFormat("$#,##0.00"));
	    try {
		ExistingDocuments ed = new ExistingDocuments(Log.findFile(sc));
		if (orderType == 3) {
		    value = ed.getValue("CI-PL");
		    totalValue += value;
		    
		    row.createCell(5).setCellValue(value);
		} else {
		    value = ed.getValue("PI");
		    totalValue += value;		    
		    
		    row.createCell(5).setCellValue(value);
		}//if-else
		row.getCell(5).setCellStyle(dollarStyle);		
	    } catch (Exception ex) {
		MasterLog.appendEntry("ERROR on SC#" + sc);
		MasterLog.appendError(ex);
	    }//try-catch
	}//for
	row = report.createRow(row.getRowNum() + 1); //incrementing row
	row = report.createRow(row.getRowNum() + 1); //incrementing row
	report.enableLocking();	
	return Math.round(totalValue * 100) / 100;
    }

    public void createPopulate() {
	try {
	report.disableLocking();	
	String[] pendingList = list[0];
	String[] confirmedList = list[1];
	String[] bookingList = list[2];
	String[] shippedList = list[3];
	String[] canceledList = list[4];
	String[] reinstatedList = list[5];

	
	row = report.createRow(0);

	row.createCell(0).setCellValue("SC#");
	row.createCell(1).setCellValue("CLIENT");
	row.createCell(2).setCellValue("FACTORY");
	row.createCell(3).setCellValue("MODEL");
	row.createCell(4).setCellValue("ISSUE DATE");
	row.createCell(5).setCellValue("VALUE");

	row = report.createRow(1);
	row.createCell(0).setCellValue("PENDING ORDERS:");
	totals[0] = populateOrderType(0);

	row.createCell(0).setCellValue("CONFIRMED ORDERS:");
	totals[1] = populateOrderType(1);

	row.createCell(0).setCellValue("BOOKING ORDERS:");
	totals[2] = populateOrderType(2);	

	row.createCell(0).setCellValue("SHIPPED ORDERS:");
	totals[3] = populateOrderType(3);

	row.createCell(0).setCellValue("CANCELED ORDERS:");
	totals[4] = populateOrderType(4);

	row.createCell(0).setCellValue("REINSTATED ORDERS:");
	totals[5] = populateOrderType(5);

	

	for (int i = 0; i < 10; i++)
	    report.autoSizeColumn(i);

	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	}//try-catch	
	try {
	    wb.write(new FileOutputStream(reportFile));
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	}
	MasterLog.append("Open Report Created");
	report.enableLocking();	
    }//createPopulate()

    public void modelLackingExcel() {
	try {
	    row = report.createRow(0);

	    row.createCell(0).setCellValue("SC#");
	    row.createCell(1).setCellValue("CLIENT");
	    row.createCell(2).setCellValue("FACTORY");
	    row.createCell(3).setCellValue("MODEL");
	    row.createCell(4).setCellValue("Status - EHF#");

	    for (int r = 0; r < list.length; r++) {
		row = report.createRow(r+1);
		for (int c = 0; c < list[r].length; c++) {
		    row.createCell(c).setCellValue(list[r][c]);
		}//for
	    }//for

	    for (int i = 0; i < 5; i++)
		report.autoSizeColumn(i);

	    wb.write(new FileOutputStream(reportFile));	    
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	}//try-catch		    

	MasterLog.append("Reminder Created");
	report.enableLocking();		
    }//modelLackingExcel()

    public double[] getTotals() {
	double[] copy = new double[totals.length];

	for (int i = 0; i < totals.length; i++)
	    copy[i] = totals[i];
	return copy;
    }//getTotals()

    protected File getReportFile() {
	return reportFile;
    }//getReportFile()
}
