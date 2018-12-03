/*
  Programmer: Patrick Nercessian

  Class Purpose:
     This class provides methods which check that certain parts of the program
     worked successfully. This helps to prevent problems (e.g. loss of internet)
     from doing too much damage to the system.
 */

package system;

import java.io.*;

import org.apache.poi.xssf.usermodel.*;

import java.util.Date;
import java.text.SimpleDateFormat;

import java.lang.StackTraceElement;

public class CheckCompletion {

    /**
     * Checks that an order has the Carbon Copy (ensures it was successfully confirmed)
     *
     * @param sc the Sales Contract Number for the order
     * @return whether or not the confirmation was validated
     */
    public static boolean checkConfirm(String sc) {
	try {
	    XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(Log.findFile(sc)));
	    Date date = new Date();
	    SimpleDateFormat sdf = new SimpleDateFormat("MM-dd-yyyy");
	    
	    int num = wb.getNumberOfSheets() - 5;
	    if (wb.getSheet("CO") == null) num++;
	    for (int i = 0; i < wb.getNumberOfSheets(); i++) {
		String sheetName = wb.getSheetAt(i).getSheetName();
		if (!(sheetName.equals("PI") || sheetName.equals("PO") || sheetName.equals("CO") || sheetName.equals("CALC SHEET")
		                                              || sheetName.equals("EMAIL TEMPLATE") || sheetName.contains("CC ")))
		    num--;
	    }//for
	    if (wb.getSheet("CC " + sdf.format(date) + " " + num) == null)
		return false;
	    
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	    return false;
	}//try-catch
	return true;
    }//checkConfirm(File)

    /**
     * Checks that an order has a Commercial Invoice / Packing List (ensures it was successfuly placed for Booking)
     *
     * @param sc the Sales Contract Number for the order
     * @return whether or not the booking was validated
     */    
    public static boolean checkBooking(String sc) {
	try {
	    XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(Log.findFile(sc)));
	    SimpleDateFormat sdf = new SimpleDateFormat("MM-dd-yyyy");

	    for (int i = 0; i < wb.getNumberOfSheets(); i++)
		if (wb.getSheetAt(i).getSheetName().startsWith("PI BOOKING " + sdf.format(new Date())))
		    return true;
	} catch (Exception ex) {
	    MasterLog.appendError(ex);	    
	    return false;
	}//try-catch
	return false;
    }//checkConfirm(File)    

    /**
     * Checks that an order has an EHF# and a PO Date (ensures that EHF#s were assigned successfully)
     *
     * @param sc the Sales Contract Number for the order
     * @return whether or not the assignment was validated
     */    
    public static boolean checkAssign(String sc) {
	boolean correct = true;
	try {
	    File f = Log.findFile(sc);
	    if (!f.getName().contains("EHF"))
		correct = false;
	    
	    XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(f));
	    XSSFRow row = wb.getSheet("PO").getRow(2);
	    if (row.getCell(4).getStringCellValue().equals(""))
		correct = false;
	    if (row.getCell(8).getStringCellValue().equals(""))
		correct = false;
	
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	    return false;	    
	}//try-catch
	return correct;
    }//checkAssign(File)

    /**
     * Checks that an order has the first line of CONSIGNEE (ensures it was created successfully)
     *
     * @param documentsWB the XSSFWorkbook attached to the order
     * @return whether or not the booking was created
     */        
    public static boolean checkCreated(XSSFWorkbook documentsWB) {
	try {
	    XSSFSheet pi = documentsWB.getSheet("PI");
	    XSSFCell firstCon = pi.getRow(CreateDocuments.findRowIndex(pi, "CONSIGNEE:", 0)).getCell(1);
	    if (firstCon == null || firstCon.getStringCellValue().equals(""))
		return false;
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	    return false;	    
	}//try-catch
	return true;
    }//checkCreated(XSSFWorkbook)
}
