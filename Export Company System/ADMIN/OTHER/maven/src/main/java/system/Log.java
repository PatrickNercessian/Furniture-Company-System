/*
  Programmer: Patrick Nercessian

  Class Purpose:
     This class aims to modify the Sales Contract Log and to gather information from it.
*/

package system;

import java.io.*;
import org.apache.poi.xssf.usermodel.*;

import java.lang.StackTraceElement;

import java.nio.file.Files;
import java.nio.file.attribute.*;
import java.nio.file.FileSystems;

import java.util.Date;
import java.text.SimpleDateFormat;

import java.nio.file.Paths;
import java.nio.file.Files;

import java.util.Properties;
import java.util.ArrayList;

import javax.activation.*;
import javax.mail.*;
import javax.mail.internet.*;

public class Log {
    private static File textFile = new File("ADMIN/OTHER/maven/src/main/resources/file/Log - Sales Contract.txt");

    
    /**
     * Compiles list of Pending orders for given EHF Sales Rep
     * 
     * @param ehfSalesRep the EHF Sales Rep
     * @param status the order status to compile list of
     * @return the list of Pending orders
     */
    public static String[] compileList(String ehfSalesRep, String status, Date minDate, Date maxDate) {
	String[] list = new String[20];
	String currentOrder;
		
	BufferedReader br;
	try {
	    br = new BufferedReader(new FileReader(textFile));
	    
	    int index = 0;
	    while ((currentOrder = br.readLine()) != null) {
		if (currentOrder.equals("")) //skip if line is empty
		    continue;

		if (status.equals("SHIPPED")) {
		    String sc = currentOrder.substring(currentOrder.indexOf("SC#:") + 5, currentOrder.indexOf(" - Client:"));
		    try {
			XSSFSheet ci = new XSSFWorkbook(new BufferedInputStream(new FileInputStream(Log.findFile(sc)))).getSheet("CI-PL");
			Date shipDate = ci.getRow(2).getCell(9).getDateCellValue();
			if (shipDate.compareTo(minDate) == -1 || shipDate.compareTo(maxDate) == 1)
			    continue;
		    } catch (NullPointerException npe) {
		    } catch (IOException ioe) {
			if (!ioe.getMessage().contains("Unable to parse xml bean"))
			    MasterLog.appendError(ioe);
		    } catch (IllegalStateException ise) { //temporary, remove after fix
			MasterLog.appendError(ise);
			XSSFSheet ci = new XSSFWorkbook(new BufferedInputStream(new FileInputStream(Log.findFile(sc)))).getSheet("CI-PL");
			try {
			    MasterLog.appendEntry("SC: " + sc + " - " + ci.getRow(2).getCell(10).getStringCellValue());
			} catch (IllegalStateException abc) {
			    MasterLog.appendEntry("SC: " + sc + " - " + ci.getRow(2).getCell(10).getNumericCellValue());
			}
		    }//try-catch

		} else if (status.equals("CANCELED")) {
		    int day = Integer.parseInt(currentOrder.substring(3, 5));
		    int month = Integer.parseInt(currentOrder.substring(0, 2)) - 1;
		    int year = Integer.parseInt(currentOrder.substring(6, 10)) - 1900;
		    Date orderDate = new Date(year, month, day);
		
		    if (orderDate.compareTo(minDate) == -1 || orderDate.compareTo(maxDate) == 1)
			continue;
		}
		    
		
		//checking if the order is for the Sales Rep and if it is the same as status parameter
		int spacesIndex = currentOrder.indexOf("       ");
		int ehfIndex = currentOrder.indexOf(" - EHF#:");
		if (!currentOrder.contains(" - EHF#:"))
		    ehfIndex = currentOrder.length();


		String client = currentOrder.substring(currentOrder.indexOf("Client:") + 8, currentOrder.indexOf(" - Factory"));
		if (isClient(client, ehfSalesRep) && currentOrder.substring(spacesIndex, ehfIndex).trim().equals(status)) {
		    if (list[list.length - 1] != null) { //if list is full, create larger array and copy over elements
			String[] copy = new String[list.length + 10];
			for (int x = 0; x < list.length; x++)
			    copy[x] = list[x];
			list = copy;
		    }//if
		    
		    list[index++] = currentOrder; //add order to list, increment index
		}//if
	    }//while
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	}//try-catch
	return list;
    }//compileList(String)

    public static ArrayList<String> modelLackingSCs() {
	ArrayList<String> list = new ArrayList<>();
	String currentOrder;

	BufferedReader br;
	try {
	    br = new BufferedReader(new FileReader(textFile));
	    int index = 0;
	    while ((currentOrder = br.readLine()) != null) {
		if (currentOrder.equals("") || currentOrder.contains("reserved")) //skip if line is empty
		    continue;
		String isUpdated = currentOrder.substring(currentOrder.indexOf("Updated:") + 9, currentOrder.indexOf("      "));
		if (isUpdated.equals("TRUE"))
		    list.add(currentOrder.substring(currentOrder.indexOf("SC#:") + 5, currentOrder.indexOf(" - Client:")));
	    }//while
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	}//try-catch
	return list;
    }

    public static String[] modelLackingOrders(String ehfSalesRep) {
	String[] list = new String[20];
	String currentOrder;
		
	BufferedReader br;
	try {
	    br = new BufferedReader(new FileReader(textFile));
	    int index = 0;
	    while ((currentOrder = br.readLine()) != null) {
		if (currentOrder.equals("") || currentOrder.contains("reserved")) //skip if line is empty
		    continue;
		String client = currentOrder.substring(currentOrder.indexOf("Client:") + 8, currentOrder.indexOf(" - Factory:"));
		String sc = currentOrder.substring(currentOrder.indexOf("SC#:") + 5, currentOrder.indexOf(" - Client:"));
		String status = getOrderStatus(sc);
		if ((status.equals("CONFIRMED") || status.equals("BOOKING") || status.equals("SHIPPED")) && isClient(client, ehfSalesRep)) {
		    String updated = currentOrder.substring(currentOrder.indexOf("Updated:") + 9, currentOrder.indexOf("      "));
		    if (updated.equals("FALSE")) {
			if (list[list.length - 1] != null) { //if list is full, create larger array and copy over elements
			    String[] copy = new String[list.length + 10];
			    for (int x = 0; x < list.length; x++)
				copy[x] = list[x];
			    list = copy;
			}//if
			list[index++] = currentOrder.substring(currentOrder.indexOf("SC#:"), currentOrder.indexOf(" - Updated:"))
			                + currentOrder.substring(currentOrder.indexOf("        "));
		    }//if
		}//if
	    }//while
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	}//try-catch
	return list;
    }//modelLackingOrders(String)


    private static boolean isClient(String client, String ehfSalesRep) throws FileNotFoundException, IOException{
	XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream("ADMIN/OTHER/maven/src/main/resources/file/ClientDatabase.xlsx"));
	XSSFSheet sheet = wb.getSheet("DB-Customers");
	int column = CreateDocuments.findColumnIndex(sheet, client, 0);
	int row = CreateDocuments.findRowIndex(sheet, "EHF Sales Rep", 0);
	if (column < 0)
	    MasterLog.appendEntry(client);
	if (sheet.getRow(row).getCell(column).getStringCellValue().equals(ehfSalesRep))
	    return true;
	return false;
    }

    public static int countLines() throws IOException {
	InputStream is = new BufferedInputStream(new FileInputStream(textFile));
	try {
	    byte[] c = new byte[1024];
	    int count = 0;
	    int readChars = 0;
	    boolean empty = true;
	    while ((readChars = is.read(c)) != -1) {
		empty = false;
		for (int i = 0; i < readChars; ++i) {
		    if (c[i] == '\n')
			count++;
		}//for
	    }//while
	    return (count == 0 && !empty) ? 1 : count;
	} finally {
	    is.close();
	}//try-finally
    }//countLines(filename)

    /**
     * Updates the Text Log for Sales Contract Entry with the logEntry parameter
     *
     * @param logEntry the String to be appended to the log 
     */
    protected static void updateLog(File f, String logEntry) {
	
	FileWriter fw = null;
	BufferedWriter bw = null;
	try {	    
	    if (f.createNewFile())
		Files.setPosixFilePermissions(f.toPath(), PosixFilePermissions.fromString("rw-r--r--"));
	    fw = new FileWriter(f, true);
	    bw = new BufferedWriter(fw);
	    bw.append(logEntry);
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	} finally {
	    try {
		if (bw != null)
		    bw.close();
		if (fw != null)
		    fw.close();
	    } catch (IOException e) {
		MasterLog.appendError(e);
	    }//try-catch
	}//try-catch-finally
	
    }//updateLog(String)

    protected static void updateExcelLog(String logEntry) {

	try {
	    File file = new File("ADMIN/OTHER/maven/src/main/resources/file/BackupLog.xlsx");
	    XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(file));
	    XSSFSheet sheet = wb.getSheet("LOG");
	    int rowNum = 0;
	    while (sheet.getRow(rowNum) != null)
		rowNum++;
	    sheet.disableLocking();
	    sheet.createRow(rowNum).createCell(0).setCellValue(logEntry);
	    sheet.enableLocking();	    
	    
	    BufferedOutputStream bos = new BufferedOutputStream(new FileOutputStream(file));
	    wb.write(bos);
	} catch (Exception ex) {
	    MasterLog.appendEntry(ex.toString());
	    StackTraceElement[] arr = ex.getStackTrace();
	    for (int i = 0; i < arr.length; i++)
		MasterLog.append(arr[i].toString());
	}//try-catch
    }//updateExcelLog(logEntry)

    /**
     * Adds the EHF# to the Log Entry
     *
     * @param ehfNum the EHF# to add
     */
    protected static void appendEHF(String ehfNum, String salesContractNum) {
	String currentOrder, desiredOrder = "", undesiredOrder = "";
	String oldText = "", newText = "";

	FileWriter fw = null;
	BufferedWriter bw = null;
	BufferedReader br = null;

	try {
	    br = new BufferedReader(new FileReader(textFile));
	    int numPreviousOrders = 0;
	    while ((currentOrder = br.readLine()) != null) {
		//		currentOrder = br.readLine();
		oldText += currentOrder + System.lineSeparator(); //by end of loop, oldText will be the entire file
		
		if (currentOrder.equals("") || currentOrder.contains("reserved")) //skip this line if it's empty or a reserve line
		    continue;
		
		numPreviousOrders++;

		if (currentOrder.substring(currentOrder.indexOf("SC#:") + 5, currentOrder.indexOf(" - Client"))
		                    .equals(salesContractNum)) { //if it's the correct sales contract order
		    desiredOrder = currentOrder + " - EHF#: " + ehfNum;
		    undesiredOrder = currentOrder;

		    //Excel Backup
		    File file = new File("ADMIN/OTHER/maven/src/main/resources/file/BackupLog.xlsx");
		    XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(file));
		    XSSFSheet sheet = wb.getSheet("LOG");
		    sheet.disableLocking();
		    sheet.getRow(numPreviousOrders - 1).getCell(0).setCellValue(desiredOrder);
		    sheet.enableLocking();
		    
		    BufferedOutputStream bos = new BufferedOutputStream(new FileOutputStream(file));
		    wb.write(bos);
		}//if
	    }//while
	    newText = oldText.replace(undesiredOrder, desiredOrder);
	    if (!newText.equals("") && newText.contains("SC#:")) {
		fw = new FileWriter(textFile);
		bw = new BufferedWriter(fw);
		
		fw.write(newText);
	    }//if
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	} finally {
	    try {
		if (bw != null)
		    bw.close();
		if (fw != null)
		    fw.close();
		if (br != null)
		    br.close();
	    } catch (IOException e) {
		MasterLog.appendError(e);
	    }
	}//try-catch-finally
	  
    }//appendEHF(String)

    /**
     * Changes the PENDING part of the Log Entry to orderStatus. This is usually CONFIRMED or CANCELLED
     *
     * @param orderStatus the new status
     * @param salesContractNum the Log Entry to change
     */
    protected static void changeOrderStatus(String orderStatus, String salesContractNum) {
	String currentOrder, desiredOrder = "", undesiredOrder = "";
	String oldText = "";
	String newText = "";
	
	FileWriter fw = null;
	BufferedWriter bw = null;
	BufferedReader br = null;

	try {
	    br = new BufferedReader(new FileReader(textFile));
	    int numPreviousOrders = 0;
	    while ((currentOrder = br.readLine()) != null) {
		oldText += currentOrder + System.lineSeparator(); //by end of loop, oldText will be the entire file

		if (currentOrder.equals("") || currentOrder.contains("reserved")) //skip this line if it's empty or a reserve line
		    continue;
		numPreviousOrders++;

		if (currentOrder.substring(currentOrder.indexOf("SC#:") + 5, currentOrder.indexOf(" - Client"))
		                        .equals(salesContractNum)) { //if it's the correct sales contract order
		    int ehfIndex = currentOrder.indexOf("EHF#:");
		    desiredOrder = currentOrder.substring(0, currentOrder.indexOf("         ") + 12) + orderStatus;
		    if (ehfIndex != -1) //adds back EHF# if present
			desiredOrder += " - " + currentOrder.substring(ehfIndex);
		    undesiredOrder = currentOrder;
		    	    
		    //Excel Backup
		    File file = new File("ADMIN/OTHER/maven/src/main/resources/file/BackupLog.xlsx");
		    XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(file));
		    XSSFSheet sheet = wb.getSheet("LOG");
		    sheet.disableLocking();
		    wb.getSheet("LOG").getRow(numPreviousOrders - 1).getCell(0).setCellValue(desiredOrder);
		    sheet.enableLocking();
		    
		    BufferedOutputStream bos = new BufferedOutputStream(new FileOutputStream(file));
		    wb.write(bos);
		    
		}//if
	    }//while
	    newText = oldText.replace(undesiredOrder, desiredOrder);
	    if (!newText.equals("")) {
		fw = new FileWriter(textFile);
		bw = new BufferedWriter(fw);
		bw.write(newText);
	    }
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	} finally {
	    try {		
		if (bw != null)
		    bw.close();
		if (fw != null)
		    fw.close();
		if (br != null)
		    br.close();
	    } catch (IOException e) {
		MasterLog.appendError(e);
	    }
	}//try-catch-finally
    }

    protected static String getOrderStatus(String sc) {
	String currentOrder;
	BufferedReader br = null;
	try {
	    br = new BufferedReader(new FileReader(textFile));
	    while ((currentOrder = br.readLine()) != null) {
		
		if (currentOrder.equals("") || currentOrder.contains("reserved")) //skip this line if it's empty or a reserve line
		    continue;

		if (currentOrder.substring(currentOrder.indexOf("SC#:") + 5, currentOrder.indexOf(" - Client"))
		                     .equals(sc)) { //if it's the correct sales contract order
		    int ehfIndex = currentOrder.indexOf(" - EHF#:");
		    if (ehfIndex == -1)
			return currentOrder.substring(currentOrder.indexOf("       ") + 12);
		    else
			return currentOrder.substring(currentOrder.indexOf("       ") + 12, ehfIndex);
		}//if
	    }//while
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	}//try-catch
	
	return "";
    
    }//getOrderStatus()


    protected static void updateModels(String scNum) {
	String currentOrder, desiredOrder = "", undesiredOrder = "";
	String oldText = "";
	String newText = "";
	
	FileWriter fw = null;
	BufferedWriter bw = null;
	BufferedReader br = null;

	try {
	    br = new BufferedReader(new FileReader(textFile));
	    int numPreviousOrders = 0;
	    while ((currentOrder = br.readLine()) != null) {
		oldText += currentOrder + System.lineSeparator(); //by end of loop, oldText will be the entire file

		if (currentOrder.equals("") || currentOrder.contains("reserved")) //skip this line if it's empty or a reserve line
		    continue;
		numPreviousOrders++;

		if (currentOrder.substring(currentOrder.indexOf("SC#:") + 5, currentOrder.indexOf(" - Client"))
		                        .equals(scNum)) { //if it's the correct sales contract order
		    int ehfIndex = currentOrder.indexOf("EHF#:");
		    desiredOrder = currentOrder.substring(0, currentOrder.indexOf("Updated:") + 9) + "TRUE" + currentOrder.substring(currentOrder.indexOf("         "));
		    undesiredOrder = currentOrder;
		    	    
		    //Excel Backup
		    File file = new File("ADMIN/OTHER/maven/src/main/resources/file/BackupLog.xlsx");
		    XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(file));
		    XSSFSheet sheet = wb.getSheet("LOG");
		    sheet.disableLocking();
		    wb.getSheet("LOG").getRow(numPreviousOrders - 1).getCell(0).setCellValue(desiredOrder);
		    sheet.enableLocking();
		    
		    BufferedOutputStream bos = new BufferedOutputStream(new FileOutputStream(file));
		    wb.write(bos);
		    
		}//if
	    }//while
	    newText = oldText.replace(undesiredOrder, desiredOrder);
	    if (!newText.equals("")) {
		fw = new FileWriter(textFile);
		bw = new BufferedWriter(fw);
		bw.write(newText);
	    }
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	} finally {
	    try {		
		if (bw != null)
		    bw.close();
		if (fw != null)
		    fw.close();
		if (br != null)
		    br.close();
	    } catch (IOException e) {
		MasterLog.appendError(e);
	    }
	}//try-catch-finally
    }    

    protected static String orderListString(String[] list, String orderType) {
	String str = "";
	int usernameIndex;
	int endIndex;	
	int count = 0;

	for (int i = 0; i < list.length; i++) { //gets number of orders
	    if (list[i] != null)
		count++;
	}

	str += "\n\n" + count + " " + orderType + " ORDERS:\n";
	
	return str;
    }//orderListString()

    public static File findFile(String sc) throws FileNotFoundException{
	String client = "", factory = "";
	String currentOrder;

	BufferedReader br;
	try { //finding client and factory names based on SC# or PO#
	    br = new BufferedReader(new FileReader(textFile));
	    while ((currentOrder = br.readLine()) != null) {
		if (currentOrder.equals("") || currentOrder.contains("reserved")) //skip if line is empty or is a reserve line
		    continue;

		//sc#
		if (currentOrder.substring(currentOrder.indexOf("SC#:") + 5, currentOrder.indexOf(" - Client:")).equals(sc)) {
		    client = currentOrder.substring(currentOrder.indexOf("Client:") + 8, currentOrder.indexOf(" - Factory:"));
		    factory = currentOrder.substring(currentOrder.indexOf("Factory:") + 9, currentOrder.indexOf(" - Model:"));
		    break;
		}//if

		//po#
		if (currentOrder.endsWith("EHF#: " + sc)) {
		    client = currentOrder.substring(currentOrder.indexOf("Client:") + 8, currentOrder.indexOf(" - Factory:"));
		    factory = currentOrder.substring(currentOrder.indexOf("Factory:") + 9, currentOrder.indexOf(" - Model:"));
		    break;
		}//if
		
	    }//while
	} catch (IOException ioe) {
	    MasterLog.appendError(ioe);
	}//try-catch

	File folder = new File("EXPORT HOME FURNISHINGS/" + client + "/" + factory);
	File[] listOfFiles = folder.listFiles();
	String name;
	for (int i = 0; i < listOfFiles.length; i++) {
	    name = listOfFiles[i].getName();
	    if (!name.contains(".xlsm"))
		continue;	    
	    if (name.startsWith(sc)) //searching for sc number
		return listOfFiles[i];
	    if (name.substring(0, name.indexOf(".xlsm")).endsWith("EHF " + sc)) //searching for po number
		return listOfFiles[i];
	}//for
	return null;
    }//findFile()

    public static boolean isValidPO(String po, String sc) {
	int poNum = Integer.parseInt(po);
	String currentOrder;
	int maxPO = 0;

	BufferedReader br;

	try {
	    br = new BufferedReader(new FileReader(textFile));
	    while ((currentOrder = br.readLine()) != null) {
		if (currentOrder.equals("") || currentOrder.contains("reserved"))
		    continue;
		
		int ehfIndex = currentOrder.indexOf("EHF#:");
		int ehfNum;
		
		if (ehfIndex != -1)
		    ehfNum = Integer.parseInt(currentOrder.substring(ehfIndex + 6));
		else
		    ehfNum = -1;

		if (currentOrder.substring(currentOrder.indexOf("SC#:") + 5, currentOrder.indexOf(" - Client")).equals(sc) && !currentOrder.contains("CONFIRMED")) //if the order is not confirmed
		    return false;
		
		if (ehfNum == poNum) //duplicate po
		    return false;
		
		if (ehfNum > maxPO)
		    maxPO = ehfNum;

	    }//while
	    //	    if (poNum < maxPO - 30 || poNum > maxPO + 30)
	    //		return false;
	} catch (IOException ioe) {
	    MasterLog.appendError(ioe);
	    return false;
	}//try-catch

	return true;
	    
    }//isValidPO(String, String)

    public static boolean isUniqueSC(String sc) { //can remove this after Reissue Button is removed
	BufferedReader br;
	String currentOrder;
	try {
	    br = new BufferedReader(new FileReader(textFile));
	    while ((currentOrder = br.readLine()) != null) {
		if (currentOrder.equals("") || currentOrder.contains("reserved"))
		    continue;
		if (currentOrder.substring(currentOrder.indexOf("SC#:") + 5, currentOrder.indexOf(" - Client")).equals(sc))
		    return false;
	    }//while
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	    return false;
	}//try-catch
	return true;
    }//isUniqueSC(String)

    public static ArrayList<String> listOfOrders(String status) {
	ArrayList<String> list = new ArrayList<>();
	BufferedReader br;
	String currentOrder;
	try {
	    br = new BufferedReader(new FileReader(textFile));
	    while ((currentOrder = br.readLine()) != null) {
		if (currentOrder.equals("") || currentOrder.contains("reserved"))
		    continue;
		
		int ehfIndex = currentOrder.indexOf(" - EHF#:");
		String currentOrderStatus;
		if (ehfIndex == -1)
		    currentOrderStatus = currentOrder.substring(currentOrder.indexOf("       ") + 12);
		else
		    currentOrderStatus = currentOrder.substring(currentOrder.indexOf("       ") + 12, ehfIndex);

		if (status.equals(currentOrderStatus))
		    list.add(currentOrder);
	    }//while
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	}//try-catch
	return list;
    }
}
