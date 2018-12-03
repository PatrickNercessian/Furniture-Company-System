/*
  Programmer: Patrick Nercessian

  Class Purpose:
     This class contains code pertaining to updating the Order Log, the log of broad
     information for every order. This is intended to be used through the Admin System
 */

package system;

import java.io.*;
import org.apache.poi.xssf.usermodel.*;

import java.lang.StackTraceElement;

import java.util.Date;
import java.util.Calendar;
import java.util.ArrayList;
import java.util.Arrays;
import java.text.SimpleDateFormat;

public class OrderLog {

    private static File file = new File("ADMIN/OTHER/maven/src/main/resources/file/Order Log.xlsm");
    private static File logFile = new File("ADMIN/OTHER/maven/src/main/resources/file/Log - Sales Contract.txt");
    private static XSSFWorkbook workbook = null;
    private static XSSFSheet sheet;
    private static ArrayList<String> scList;

    public static void setScList() throws IOException, FileNotFoundException{
	ArrayList<String> list = new ArrayList<String>(5000);
	/*
	workbook = new XSSFWorkbook(new BufferedInputStream(new FileInputStream(file)));
	sheet = workbook.getSheet("LOG");
	for (int i = 1; i < sheet.getLastRowNum(); i++) {
	    try {
		list.add(sheet.getRow(i).getCell(0).getStringCellValue());
	    } catch (NullPointerException npe) {
	    } catch (IllegalStateException ise) {
		list.add("" + (int) sheet.getRow(i).getCell(0).getNumericCellValue());
	    }
	    
	}//for
	*/

	try {
	    String currentOrder;
	    BufferedReader br = new BufferedReader(new FileReader(logFile));
	    while ((currentOrder = br.readLine()) != null) {
		if (currentOrder.equals("") || currentOrder.contains("reserved")) //skip if line is empty
		    continue;
		String sc = currentOrder.substring(currentOrder.indexOf("SC#:") + 5, currentOrder.indexOf(" - Client:"));
		//		MasterLog.appendEntry(sc);		
		list.add(sc);
	    }//while
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	}//try-catch	
	String[] arr = new String[list.size()];
	arr = Arrays.stream(list.toArray(arr)).distinct().toArray(String[]::new); //removing duplicates
	scList = new ArrayList<String>(Arrays.asList(arr));
    }

    public static int findOrderRow(String scStr) throws IOException, FileNotFoundException {
	//	workbook = new XSSFWorkbook(new BufferedInputStream(new FileInputStream(file)));
	//	sheet = workbook.getSheet("LOG");
	int start = 1;
	int mid;
	int end = sheet.getLastRowNum();
	while (sheet.getRow(end).getCell(0) == null)
	    end--;

	int sc = Integer.parseInt(scStr);
	int midSC;	
	int desiredIndex = 1;

	int compareto;
	
	while (start <= end) {
	    mid = (start + end) / 2;
	    if (sheet.getRow(mid) == null)
		MasterLog.appendEntry(mid + " is null");
	    else if (sheet.getRow(mid).getCell(0) == null)
		MasterLog.appendEntry(mid  + " cell is null");

	    midSC = (int) sheet.getRow(mid).getCell(0).getNumericCellValue();

	    if (sc > midSC) {
		start = mid + 1;
	    } else if (sc < midSC) {
		end = mid - 1;
	    } else {
		desiredIndex = mid;
		break;
	    }

	    if (start > end) {
		    if (end == 0)
			desiredIndex = 1;
		    else if (sc < sheet.getRow(end).getCell(0).getNumericCellValue())
			desiredIndex = end;
		    else
			desiredIndex = end + 1;
	    }//if
	}//while
	return desiredIndex;
    }

    public static void populateAllOrders() throws IOException, FileNotFoundException{
	workbook = new XSSFWorkbook(new BufferedInputStream(new FileInputStream(file)));	
	sheet = workbook.getSheet("LOG");
	sheet.disableLocking();
	//	String[] list = listOfSC();
	//	for (int i = 0; i < list.size(); i++) {
	setScList();
	while (scList.size() != 0) {
	    MasterLog.appendEntry(scList.get(0));
	    OrderLog.populateOrder(scList.get(0));
	}
	sheet.enableLocking();
	writeFile();
    }

    public static void populateOrder(String scStr) throws IOException, FileNotFoundException {
	/*
	MasterLog.appendEntry("finding row...");	
	int index = findOrderRow(scStr);
	MasterLog.appendEntry("found row.");		
	String indexedSC = "";
	try {
	    indexedSC = "" + (int) sheet.getRow(index).getCell(0).getNumericCellValue();
	} catch (NullPointerException npe) {
	}
	
	MasterLog.appendEntry("shifting rows...");
	if (!indexedSC.equals(scStr))
	    sheet.shiftRows(index, sheet.getLastRowNum() + 10, 1);
	MasterLog.appendEntry("shifted rows.");	
	XSSFRow row = sheet.createRow(index);
	*/

	int index = sheet.getLastRowNum() + 1;
	for (int i = 0; i < sheet.getLastRowNum(); i++) {
	    try {
		if (sheet.getRow(i).getCell(0).getNumericCellValue() == Integer.parseInt(scStr)) {
		    index = i;
		    break;
		}//if
	    } catch (IllegalStateException ise) {} //if string skip check
	}//for

	XSSFRow row = sheet.createRow(index);

	String orderEntry = findOrderEntry(scStr);
	String client = findClient(orderEntry);
	try {
	    XSSFCellStyle pctStyle = workbook.createCellStyle();
	    pctStyle.setDataFormat(workbook.createDataFormat().getFormat("0%"));

	    row.createCell(0).setCellValue(Integer.parseInt(scStr));
	    XSSFWorkbook wb = getWB(scStr);
	    if (wb != null) {
		row.createCell(1).setCellValue(findPiDate(wb));
		row.createCell(2).setCellValue(findPiAmount(wb));
		row.createCell(3).setCellValue(findEhfNum(orderEntry));
		row.createCell(4).setCellValue(findPoDate(wb));
		row.createCell(5).setCellValue(findPoAmount(wb));
		row.createCell(6).setCellValue(findSalesRep(client));
		row.createCell(7).setCellValue(client);
		row.createCell(8).setCellValue(findCountry(client));
		row.createCell(9).setCellValue(findFactory(orderEntry));
		row.createCell(10).setCellValue(findModel(orderEntry));
		row.createCell(11).setCellValue(findNumContainers(wb));
		row.createCell(12).setCellValue(findShipDate(wb));
		row.createCell(13).setCellValue(findCiAmount(wb));
		row.createCell(14).setCellFormula("IF(C" + (index+1) + "-F" + (index+1) + "<>0,C" + (index+1) + "-F" + (index+1) + ",\"\")");
		row.createCell(15).setCellFormula("IF(O" + (index+1) + "/C" + (index+1) + "<>0,O" + (index+1) + "/C" + (index+1) + ",\"\")");
		row.getCell(15).setCellStyle(pctStyle);	    
		row.createCell(16).setCellFormula("IF(N" + (index+1) + "<>0,IF(N" + (index+1) + "-F" + (index+1) + "<>0,N" + (index+1) + "-F" + (index+1) + ",\"\"),\"\")");
		row.createCell(17).setCellFormula("IF(Q" + (index+1) + "<>\"\",Q" + (index+1) + "/N" + (index+1) + ",\"\")");
		row.getCell(17).setCellStyle(pctStyle);
		row.createCell(18).setCellValue(Log.getOrderStatus(scStr));
	    }//if
	} catch (IOException ioe) { //do nothing if ioexception
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	}//try-catch
	
	scList.remove(scStr);

	//	writeFile();
    }//onCreation(String)

    

    private static void writeFile() {
	try {
	    //	    workbook = new XSSFWorkbook(new BufferedInputStream(new FileInputStream(file)));
	    workbook.write(new BufferedOutputStream(new FileOutputStream(file)));
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	}//try-catch
    }//writeFile()


    public static String findOrderEntry(String scStr) {
	BufferedReader br;
	try {
	    br = new BufferedReader(new FileReader(logFile));
	    String currentOrder;
	    while ((currentOrder = br.readLine()) != null) {
		if (currentOrder.equals("") || currentOrder.contains("reserved"))
		    continue;
		if (currentOrder.substring(currentOrder.indexOf("SC#:") + 5, currentOrder.indexOf(" - Client:")).equals(scStr))
		    return currentOrder;
	    }//while
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	}//try-catch
	return "OrderNotFoundError";
    }

    public static String findClient(String orderEntry) {
	String client = "ClientNotFoundError";
	if (!orderEntry.equals("OrderNotFoundError"))
	    client = orderEntry.substring(orderEntry.indexOf("Client:") + 8, orderEntry.indexOf(" - Factory:"));
	return client;
    }//getClient(String)
    
    public static String findModel(String orderEntry) {
	String model = "ModelNotFoundError";
	if (!orderEntry.equals("OrderNotFoundError"))	
	    model = orderEntry.substring(orderEntry.indexOf("Model:") + 7, orderEntry.indexOf(" - Updated:"));
	return model;
    }//getModel(String)

    public static String findFactory(String orderEntry) {
	String factory = "FactoryNotFoundError";
	if (!orderEntry.equals("OrderNotFoundError"))		
	    factory = orderEntry.substring(orderEntry.indexOf("Factory:") + 9, orderEntry.indexOf(" - Model:"));
	return factory;
    }//getFactory(String)

    public static String findEhfNum(String orderEntry) {
	String ehfNum = "EHF#NotFoundError";
	if (!orderEntry.equals("OrderNotFoundError")) { 
	    if (orderEntry.contains(" - EHF#:"))
		ehfNum = orderEntry.substring(orderEntry.indexOf(" - EHF#:") + 9);
	}//if
	return ehfNum;
    }
    
    public static String findCountry(String client) {
	try {
	    XSSFSheet db = new XSSFWorkbook(new FileInputStream("ADMIN/OTHER/maven/src/main/resources/file/ClientDatabase.xlsx"))
		.getSheet("DB-Customers");

	    int column = CreateDocuments.findColumnIndex(db, client, 0);
	    return db.getRow(CreateDocuments.findRowIndex(db, "Country", 0)).getCell(column).getStringCellValue();
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	}//try-catch
	return "CountryNotFoundError";	
    }//getCountry(String)

    public static String findSalesRep(String client) {
	try {
	    XSSFSheet db = new XSSFWorkbook(new FileInputStream("ADMIN/OTHER/maven/src/main/resources/file/ClientDatabase.xlsx"))
		.getSheet("DB-Customers");

	    int column = CreateDocuments.findColumnIndex(db, client, 0);
	    return db.getRow(CreateDocuments.findRowIndex(db, "EHF Sales Rep", 0)).getCell(column).getStringCellValue();
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	}//try-catch
	return "CountryNotFoundError";	
    }//getCountry(String)

    public static XSSFWorkbook getWB(String scStr) throws FileNotFoundException, IOException {
	File f = Log.findFile(scStr);
	if (f != null)
	    return new XSSFWorkbook(new FileInputStream(f));
	else
	    return null;
    }

    public static String findPiDate(XSSFWorkbook wb) {
	String piDate = "N/A";
	XSSFCell cell = wb.getSheet("PI").getRow(2).getCell(9);
	try {
	    if (cell != null && !cell.getStringCellValue().equals(""))
		piDate = cell.getStringCellValue().replaceAll("'", "");
	} catch (IllegalStateException ise) {
	    piDate = cell.getDateCellValue().toString();
	}
	return piDate;
    }
    
    public static String findPoDate(XSSFWorkbook wb) {
	String poDate = "N/A";
	XSSFCell cell = wb.getSheet("PO").getRow(2).getCell(8);
	try {
	    if (cell != null && !cell.getStringCellValue().equals(""))
		poDate = cell.getStringCellValue().replaceAll("'", "");
	} catch (IllegalStateException ise) {
	    poDate = cell.getDateCellValue().toString();
	}	
	return poDate;
    }//getPoDate(String)

    public static String findShipDate(XSSFWorkbook wb) {
	String ciDate = "N/A";
	XSSFSheet ci = wb.getSheet("CI-PL");
	if (ci != null) {
	    XSSFCell cell = ci.getRow(2).getCell(9);
	    try {
		if (cell != null && !cell.getStringCellValue().equals(""))
		    ciDate = cell.getStringCellValue().replaceAll("'", "");
	    } catch (IllegalStateException ise) {
		Date d = cell.getDateCellValue();
		SimpleDateFormat sdf = new SimpleDateFormat("MM/dd/yyyy");
		ciDate = sdf.format(d);
	    }
	}//if
	return ciDate;
    }//getShipDate(XSSFWorkbook)

    public static Date findShipDateDate(XSSFWorkbook wb) {
	Date date = null;
	XSSFSheet ci = wb.getSheet("CI-PL");
	if (ci != null) {
	    XSSFCell cell = ci.getRow(2).getCell(9);
	    try {
		if (cell != null && cell.getNumericCellValue() != 0)
		    date = cell.getDateCellValue();
	    } catch (Exception ex) {
		MasterLog.appendError(ex);
	    }
	}//if
	return date;
    }//getShipDate(XSSFWorkbook)    

    public static double findPiAmount(XSSFWorkbook wb) {
	XSSFSheet pi = wb.getSheet("PI");
	int col = CreateDocuments.findColumnIndex(pi, "AMOUNT", CreateDocuments.findRowIndex(pi, "MODEL", 0));
	int row = CreateDocuments.findRowIndex(pi, "TOTAL:", 5);
	return pi.getRow(row).getCell(col).getNumericCellValue();
    }
    public static double findPoAmount(XSSFWorkbook wb) {
	XSSFSheet po = wb.getSheet("PO");
	int col = CreateDocuments.findColumnIndex(po, "PRICE", CreateDocuments.findRowIndex(po, "MODEL", 0)) + 1;
	int row = CreateDocuments.findRowIndex(po, "TOTAL:", 3);
	return po.getRow(row).getCell(col).getNumericCellValue();
    }
    public static double findCiAmount(XSSFWorkbook wb) {
	XSSFSheet ci = wb.getSheet("CI-PL");
	if (ci != null) {
	    int col = CreateDocuments.findColumnIndex(ci, "AMOUNT", CreateDocuments.findRowIndex(ci, "MODEL", 0));
	    int row = CreateDocuments.findRowIndex(ci, "TOTAL:", 5);
	    return ci.getRow(row).getCell(col).getNumericCellValue();
	} else {
	    return 0;
	}//if-else
    }

    public static int findNumContainers(XSSFWorkbook wb) {
	XSSFSheet pi = wb.getSheet("PI");
	int row = CreateDocuments.findRowIndex(pi, "CONTAINER:", 0);
	String container = pi.getRow(row).getCell(1).getStringCellValue();
	return Integer.parseInt(container.substring(0, container.indexOf("X")));
    }//numContainers(XSSFWorkbook)
}
