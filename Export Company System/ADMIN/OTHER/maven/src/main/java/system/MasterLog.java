/*
  Programmer: Patrick Nercessian

  Class Purpose:
     This class updates the Master Log whenever certain aspects of the program
     are used. Furthermore, it updates the Error Log whenever an error occurs,
     and emails each error to pjnercessian@gmail.com
 */

package system;

import java.io.*;

import org.apache.poi.xssf.usermodel.*;

import java.nio.file.Files;
import java.nio.file.attribute.PosixFilePermissions;

import java.util.Date;
import java.text.SimpleDateFormat;

public class MasterLog {
    private static File masterLogFile = new File("ADMIN/OTHER/maven/src/main/resources/file/Master Log.txt");
    private static File errorLogFile = new File("ADMIN/OTHER/maven/src/main/resources/file/Error Log.txt");

    private static String errorMsg;

    public static String appendEntry(String entry) {
	String result = null;
	
	Date date = new Date();
	SimpleDateFormat sdf = new SimpleDateFormat("MM/dd/yyyy hh:mm:ss");
	
	FileWriter fw = null;
	BufferedWriter bw = null;
	try {
	    if (masterLogFile.createNewFile()) {
		Files.setPosixFilePermissions(masterLogFile.toPath(), PosixFilePermissions.fromString("rw-r--r--"));
	    }//if
	    fw = new FileWriter(masterLogFile, true);
	    bw = new BufferedWriter(fw);
	    
	    result = "\n\n" + sdf.format(date) + " - " + System.getProperty("user.name") + " - " + entry;
	    bw.append(result);
	} catch (Exception ex) {
	    ex.printStackTrace();
	} finally {
	    try {
		if (bw != null)
		    bw.close();
		if (fw != null)
		    fw.close();
	    } catch (IOException ioe) {
		ioe.printStackTrace();
	    }//try-catch
	}//try-catch-finally
	updateExcelLog(entry);
	return result;
    }//appendEntry(String)

    public static String append(String str) {
	String result = null;
	
	FileWriter fw = null;
	BufferedWriter bw = null;
	try {
	    if (masterLogFile.createNewFile()) {
		Files.setPosixFilePermissions(masterLogFile.toPath(), PosixFilePermissions.fromString("rw-r--r--"));
	    }//if
	    fw = new FileWriter(masterLogFile, true);
	    bw = new BufferedWriter(fw);
	    result = "\n" + str;
	    bw.append(result);
	} catch (Exception ex) {
	    ex.printStackTrace();
	} finally {
	    try {
		if (bw != null)
		    bw.close();
		if (fw != null)
		    fw.close();
	    } catch (IOException ioe) {
		ioe.printStackTrace();
	    }//try-catch
	}//try-catch-finally
	updateExcelLog(str);	
	return result;
    }//append(String)

    public static String appendError(Exception ex) {
	String result = null;
	Date date = new Date();
	SimpleDateFormat sdf = new SimpleDateFormat("MM/dd/yyyy hh:mm:ss");

	errorMsg = ex.toString();
	StackTraceElement[] arr = ex.getStackTrace();
	for (int i = 0; i < arr.length; i++)
	    errorMsg += "\n" + arr[i].toString();
	
	FileWriter fw = null;
	FileWriter fwMaster = null;	
	BufferedWriter bw = null;
	BufferedWriter bwMaster = null;
	try {
	    if (errorLogFile.createNewFile()) {
		Files.setPosixFilePermissions(errorLogFile.toPath(), PosixFilePermissions.fromString("rw-r--r--"));
	    }//if
	    fw = new FileWriter(errorLogFile, true);
	    fwMaster = new FileWriter(masterLogFile, true);	    
	    bw = new BufferedWriter(fw);
	    bwMaster = new BufferedWriter(fwMaster);
	    
	    result = "\n\n\n\n" + sdf.format(date) + " - " + System.getProperty("user.name") + " - " + errorMsg;
	    bw.append(result);
	    bwMaster.append(result);
	    Thread t = new Thread(() -> {
		    if (Email.checkLogin("office@ehfurnishings.com", "Cattiger321") && !ex.getMessage().contains("The process cannot access the file because it is being used by another process"))
			Email.sendEmail("office@ehfurnishings.com", "pjnercessian@gmail.com", "There has been an error in the EHF System!", errorMsg);
		});
	    t.setName("Error Email Thread " + Thread.activeCount());
	    t.setDaemon(true);
	    t.start();
	} catch (Exception exception) {
	    exception.printStackTrace();
	} finally {
	    try {
		if (bw != null)
		    bw.close();
		if (fw != null)
		    fw.close();
	    } catch (IOException ioe) {
		ioe.printStackTrace();
	    }//try-catch
	}//try-catch-finally
	return result;
    }

    protected static void updateExcelLog(String logEntry) {

	try {
	    Date date = new Date();
	    SimpleDateFormat sdf = new SimpleDateFormat("MM/dd/yyyy hh:mm:ss");
	    
	    File file = new File("ADMIN/OTHER/maven/src/main/resources/file/BackupMasterLog.xlsx");
	    XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(file));
	    XSSFSheet sheet = wb.getSheet("LOG");
	    int rowNum = 0;
	    while (sheet.getRow(rowNum) != null)
		rowNum++;
	    sheet.disableLocking();
	    sheet.createRow(rowNum).createCell(0).setCellValue(sdf.format(date) + " - " + logEntry);
	    sheet.enableLocking();	    
	    
	    BufferedOutputStream bos = new BufferedOutputStream(new FileOutputStream(file));
	    wb.write(bos);
	} catch (Exception ex) {
	    ex.printStackTrace();
	}//try-catch
    }//updateExcelLog(logEntry)
}//MasterLog
