package system;

import  java.io.*;
import  org.apache.poi.hssf.usermodel.HSSFSheet;
import  org.apache.poi.hssf.usermodel.HSSFWorkbook;
import  org.apache.poi.hssf.usermodel.HSSFRow;
import  org.apache.poi.hssf.usermodel.HSSFCell;

import java.util.Date;
import java.text.DateFormat;
import java.text.SimpleDateFormat;

public class SalesContract {

    private HSSFWorkbook workbook;
    private HSSFSheet sheet;

    private String client;
    private String[] toAddress;
    private String[] attention;
    private String factory;
    private static int id;


    public SalesContract(String client, String[] toAddress, String[] attention, String factory) {
	this.workbook = new HSSFWorkbook();
	this.sheet = workbook.createSheet("Proforma");

	this.client = client;
	this.toAddress = toAddress; //may have to make a for loop and copy over strings
	this.attention = attention; //may have to make a for loop and copy over strings
	this.factory = factory;

	fillConstants();

	this.id++;
    }//SalesContract()

    private void fillConstants() {
	short row = 0;
	HSSFRow info;

	DateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy");
	Date date = new Date();


	info = sheet.createRow(row++);
	info.createCell(0).setCellValue("E.H.F, Inc.");

	info = sheet.createRow(row++);
	info.createCell(0).setCellValue("3635 PEACHTREE INDUSTRIAL BLVD (SUITE 700)");

	info = sheet.createRow(row++);
	info.createCell(0).setCellValue("DULUTH, GA 30096 U.S.A.");

	row++; //skip a row

	info = sheet.createRow(row++);
	info.createCell(0).setCellValue("EMAIL: EHF@EHFURNISHINGS.COM");

	info = sheet.createRow(row++);
	info.createCell(0).setCellValue("TEL: +1 (678) 646-0476");

	info = sheet.createRow(row++);
	info.createCell(0).setCellValue("FAX: +1 (678) 646-0482");

	row++; //skip a row

	info = sheet.createRow(row++); //DOESNT INCREMENT
	info.createCell(0).setCellValue("TO:");
	info.createCell(1).setCellValue(this.client);
	info.createCell(3).setCellValue("DATE: " + dateFormat.format(date));
	
	for (int i = 0; i < this.toAddress.length; i++) {
		info = sheet.createRow(row++);
		info.createCell(1).setCellValue(this.toAddress[i]);
	}//for

	row++; //skip a row

	info = sheet.createRow(row); //DOESNT INCREMENT
	info.createCell(0).setCellValue("ATTENTION:");	

	for (int i = 0; i < this.attention.length; i++) {
	    if (i == 0) {
		info.createCell(1).setCellValue(this.attention[0]);
		row++;
	    } else {
		info = sheet.createRow(row++);
		info.createCell(1).setCellValue(this.toAddress[i]);
	    }//if-else
	}//for

	info = sheet.createRow(row++);
	info.createCell(0).setCellValue("SALES CONTRACT:");
	info.createCell(1).setCellValue(id);

	info = sheet.createRow(row++);
	info.createCell(0).setCellValue("FACTORY:");
	info.createCell(1).setCellValue(factory);

	row++; //skip a row

	info = sheet.createRow(row++);
	info.createCell(0).setCellValue("WE CONFIRM HAVING SOLD YOU THE FOLLOWING:");

	row++; //skip a row
	row++; //skip a row

	info = sheet.createRow(row);
	info.createCell(0).setCellValue("MODEL");
	info.createCell(1).setCellValue("DESCRIPTION");
	info.createCell(2).setCellValue("FABRIC/COLOR");
	info.createCell(3).setCellValue("QTY");
	info.createCell(4).setCellValue("PRICE");
	info.createCell(5).setCellValue("AMOUNT");

	
	System.out.println(sheet.getColumnWidth(0));
	sheet.setColumnWidth(0, 4608);
	sheet.setColumnWidth(1, 9472);
	sheet.setColumnWidth(2, 7084);
	sheet.setColumnWidth(3, 2048);
	sheet.setColumnWidth(4, 3072);
	sheet.setColumnWidth(5, 3072);


    }//fillConstants()

    public void autoPopulate(File file) {
	try {
	    FileInputStream clientDB = new FileInputStream(file);
	    HSSFWorkbook databaseReader = new HSSFWorkbook(clientDB);
	    HSSFSheet sheet = databaseReader.getSheetAt(0);
	    HSSFRow row = sheet.getRow(0);
	    HSSFCell currentCell;
	    
	    //assigns HSSFRow to the correct row for the client
	    for (int i = 1; i < sheet.getLastRowNum(); i++) {
		row = sheet.getRow(i);
		if (row.getCell(0).getStringCellValue().equalsIgnoreCase(this.client))
		    break;
		if (i == sheet.getLastRowNum() - 1) //if not found
		    System.out.println("ERROR: Client not found");
	    }//for

	    currentCell = row.getCell(1);
	} catch (Exception ex) {
	    System.out.println(ex);
	}//try-catch
    }//autoPopulate()

    public static void main(String[] args) {
	
	String[] addy = {"SOS. FABRICA DE GLUCOZA NO. 21. SECT 2", "BUCHAREST ZIP 020332 - ROMANIA"};
	String[] atten= {"MS. ROXANA GAINA", "GENERAL MANAGER"};
	
	SalesContract scTest = new SalesContract("LINEA MEX SRL", addy, atten, "AMERICAN FURNITURE");

	try {
	    FileOutputStream fileOut = new FileOutputStream("/Users/patricknercessian/Desktop/Export Home Furnishings Project/test.xls");
	    scTest.workbook.write(fileOut);
	    fileOut.close();
	} catch (Exception ex) {
	    System.out.println(ex);
	}//try-catch
    }//main(String[])

}//SalesContract