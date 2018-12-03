/*
  Programmer: Patrick Nercessian

  Class Purpose:
     This class aims to provide ways to modify and check aspects of existing Excel Documents.
 */

package system;

import java.io.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.commons.io.FileUtils;

import java.lang.StackTraceElement;

import java.nio.file.Files;

import java.util.Date;
import java.text.SimpleDateFormat;

public class ExistingDocuments {

    private File file;
    private XSSFWorkbook wb = null;

    public ExistingDocuments(File f) {
	try { 
	    this.file = f;
	    BufferedInputStream bif = new BufferedInputStream(new FileInputStream(f));
	    wb = new XSSFWorkbook(bif);
	} catch (Exception ex) {
	    if (!ex.getMessage().contains("Unable to parse xml bean"))
		MasterLog.appendError(ex);
	}//try-catch
    }//ExistingDocuments

    public boolean containsPO() {
	if (wb != null) {
	    boolean contains;
	
	    XSSFSheet po = wb.getSheet("PO");
	    XSSFCell cell = po.getRow(2).getCell(4);
	
	    if (cell == null || cell.getStringCellValue().equals("")) contains = false;
	    else contains = true;

	    cell = po.getRow(2).getCell(8);
	    if (cell == null || cell.getStringCellValue().equals("")) contains = false;
	    else contains = true;

	    return contains;
	} else {
	    return false;
	}
    }

    public String insertPO(String num) {
	if (wb != null) {	
	    XSSFSheet po = wb.getSheet("PO");
	    XSSFSheet et = wb.getSheet("EMAIL TEMPLATE");
	    XSSFCell cell;
	
	    Date date = new Date();
	    SimpleDateFormat sdf = new SimpleDateFormat("MM/dd/yyyy");

	    cell = po.getRow(2).getCell(4);
	    if (cell == null || cell.getStringCellValue().equals("")) {
		po.getRow(2).createCell(4).setCellValue("EHF " + num);
	    } else {
		return "ERROR: PO NOT ASSIGNED: PO# " + num + " already exists in either PO sheet or CI-PL sheet";
	    }//if-else

	    po.getRow(2).createCell(8).setCellValue(sdf.format(date));

	    //finishing email template
	    et.getRow(CreateDocuments.findRowIndex(et, "Subject:", 2)).createCell(3).setCellValue("EHF " + num);
	    et.getRow(CreateDocuments.findRowIndex(et, "LINE 3", 2)).createCell(3).setCellValue("Please find attached our purchase"
												+ " order EHF " + num + " and proceed"
												+ " accordingly.");
	    try {//renaming file
		BufferedOutputStream bos = new BufferedOutputStream(new FileOutputStream(file));
		wb.write(bos);
		String name = file.getName();
		name = name.substring(0, name.indexOf(".xl")) + " - EHF " + num + name.substring(name.indexOf(".xl"));
		Files.move(file.toPath(), file.toPath().resolveSibling(name));
	    } catch (Exception ex) {
		MasterLog.appendError(ex);
	    }//try-catch
	
	    return "";
	} else {
	    return "";
	}
    }//insertPO(File)

    public void addClient(boolean addingNew, String name, String ehfRep, String country, String portOfDischarge, String[] consignee, String[] notify, String[] requirements, String[] notes, boolean[] redNotes, String payment) {
	if (wb != null) {	
	    XSSFSheet sheet = wb.getSheet("DB-Customers");
	    sheet.disableLocking();
	
	    String[][] allArrays = {consignee, notify, requirements, notes};

	    //correlates with allArrays[][]
	    int[] indices = {CreateDocuments.findRowIndex(sheet, "Consignee", 0), CreateDocuments.findRowIndex(sheet, "Notify", 0),
			     CreateDocuments.findRowIndex(sheet, "Colums/Requirements", 0), CreateDocuments.findRowIndex(sheet, "Notes", 0)};
	
	    int column = 0;
	    if (addingNew) {
		while (sheet.getRow(0).getCell(column) != null)
		    column++;
	    } else {
		while (!sheet.getRow(0).getCell(column).getStringCellValue().equals(name))
		    column++;   
	    }
	
	    XSSFCell cell = sheet.getRow(0).createCell(column);
	    cell.setCellValue(name);

	    for (int i = 0; i < allArrays.length; i++) {
		cell = sheet.getRow(indices[i]).createCell(column); //move to next subsection
		for (int j = 0; j < allArrays[i].length; j++) {
		    cell.setCellValue(allArrays[i][j]);
		    if (allArrays[i] == notes && redNotes[j] == true)
			cell.setCellValue(allArrays[i][j] + "(RED)");
		    if (sheet.getRow(cell.getRowIndex() + 1) == null) sheet.createRow(cell.getRowIndex() + 1);
		    cell = sheet.getRow(cell.getRowIndex() + 1).createCell(column); //moving cell down one
		}//for
	    }//for

	    cell = sheet.getRow(CreateDocuments.findRowIndex(sheet, "Payment", 0)).createCell(column);
	    cell.setCellValue(payment);

	    sheet.getRow(CreateDocuments.findRowIndex(sheet, "EHF Sales Rep", 0)).createCell(column).setCellValue(ehfRep);

	    sheet.getRow(CreateDocuments.findRowIndex(sheet, "Country", 0)).createCell(column).setCellValue(country);

	    sheet.getRow(CreateDocuments.findRowIndex(sheet, "Port of Discharge", 0)).createCell(column).setCellValue(portOfDischarge);

	    sheet.autoSizeColumn(column);
	
	    sheet.enableLocking();
	    try {
		BufferedOutputStream bos = new BufferedOutputStream(new FileOutputStream(file));
		wb.write(bos);
	    } catch (Exception ex) {
		MasterLog.appendError(ex);
	    }//try-catch
	}//if
    }//addClient()


    public void addFactory(boolean addingNew, String name, String address1, String address2, String contact, String ship,
			   double discount, String to, String cc, String bcc) {
	if (wb != null) {	
	    XSSFSheet sheet = wb.getSheet("DB-Factories");
	    sheet.disableLocking();
	
	    int row = 1;
	    if (addingNew) {
		while (sheet.getRow(row) != null && !sheet.getRow(row).getCell(0).getStringCellValue().equals(""))
		    row++;
		sheet.createRow(row);
	    } else {
		while (!sheet.getRow(row).getCell(0).getStringCellValue().equals(name))
		    row++;
		sheet.createRow(row);
	    }//if-else
	
	    sheet.getRow(row).createCell(0).setCellValue(name);
	    sheet.getRow(row).createCell(1).setCellValue(address1);
	    sheet.getRow(row).createCell(2).setCellValue(address2);
	    sheet.getRow(row).createCell(3).setCellValue(contact);
	    sheet.getRow(row).createCell(4).setCellValue(ship);
	    sheet.getRow(row).createCell(5).setCellValue(discount);
	    sheet.getRow(row).getCell(5).setCellStyle(sheet.getRow(row-1).getCell(5).getCellStyle());
	    sheet.getRow(row).createCell(6).setCellValue(to);
	    sheet.getRow(row).createCell(7).setCellValue(cc);
	    sheet.getRow(row).createCell(8).setCellValue(bcc);

	    sheet.enableLocking();
	
	    try {
		BufferedOutputStream bos = new BufferedOutputStream(new FileOutputStream(file));
		wb.write(bos);
	    } catch (Exception ex) {
		MasterLog.appendError(ex);
	    }//try-catch
	}//if
    }//addFactory(String, String, String, String, double, String, String, String)

    

    public boolean hasContainerNum() {
	if (wb != null) {	
	    XSSFSheet sheet = wb.getSheet("CI-PL");
	    if (sheet == null) return false; //CI-PL has not yet been created
	    
	    int row = CreateDocuments.findRowIndex(sheet, "CONTAINER/TRAILER ID:", 3);
	    try {
		if (sheet != null) {
		    if (sheet.getRow(row).getCell(4) == null)
			return false;
		    else if (sheet.getRow(row).getCell(4).getStringCellValue().equals(""))
			return false;
		}//if
	    } catch (Exception ex) {
		MasterLog.appendError(ex);
		if (ex instanceof IllegalStateException) {
		    if (sheet.getRow(row).getCell(4).getNumericCellValue() == 0)
			return false;
		}
	    }
	    return true;
	} else {
	    return false;
	}
    }//hasContainerNumber()

    /**
     * Get the value of an order
     * 
     * @param sheetName should be PI (or CI-PL if it's a SHIPPED order)
     * @return the value
     */
    public double getValue(String sheetName) {
	if (wb != null) {	
	    XSSFSheet sheet = wb.getSheet(sheetName);
	
	    int col = CreateDocuments.findColumnIndex(sheet, "AMOUNT", CreateDocuments.findRowIndex(sheet, "MODEL", 0));
	    int row = CreateDocuments.findRowIndex(sheet, "TOTAL:", 5);		

	    return sheet.getRow(row).getCell(col).getNumericCellValue();
	} else {
	    return 0;
	}
    }

    public void confirmedCopyPI() {
	if (wb != null) {
	    Date date = new Date();
	    SimpleDateFormat sdf = new SimpleDateFormat("MM-dd-yyyy");

	    int count = 0;
	    for (int i = 0; i < wb.getNumberOfSheets(); i++)
		if (wb.getSheetAt(i).getSheetName().startsWith("CC ")) count++;
	    XSSFSheet newSheet = wb.cloneSheet(wb.getSheetIndex("PI"), "CC " + sdf.format(date) + " " + (count+1));
	    newSheet.protectSheet("ar786");

	    try {
		BufferedOutputStream bos = new BufferedOutputStream(new FileOutputStream(file));
		wb.write(bos);
	    } catch (Exception ex) {
		MasterLog.appendError(ex);
	    }
	}//if
    }//copyPI()

    public void bookingCopyPI() {
	if (wb != null) {	
	    Date date = new Date();
	    SimpleDateFormat sdf = new SimpleDateFormat("MM-dd-yyyy");

	    XSSFSheet newSheet = wb.cloneSheet(wb.getSheetIndex("PI"), "PI BOOKING " + sdf.format(date));
	    newSheet.protectSheet("ar786");

	    try {
		BufferedOutputStream bos = new BufferedOutputStream(new FileOutputStream(file));
		wb.write(bos);
	    } catch (Exception ex) {
		MasterLog.appendError(ex);
	    }
	}//if
    }//copyPI()

    public void createCI() {
	if (wb != null) {	
	    XSSFSheet pi = wb.getSheet("PI");
	    XSSFSheet po = wb.getSheet("PO");
	    XSSFSheet newSheet = wb.cloneSheet(wb.getSheetIndex("PI"), "CI-PL");
	    newSheet.disableLocking();
	    int invoiceIndex = CreateDocuments.findRowIndex(pi, "NOTES:", 6);

	    newSheet.getRow(0).getCell(0).setCellValue("COMMERCIAL INVOICE / PACKING LIST");
	

	    newSheet.getRow(invoiceIndex).createCell(3).setCellValue("INVOICE NUMBER:");
	    newSheet.getRow(invoiceIndex).createCell(4).setCellValue(po.getRow(2).getCell(4).getStringCellValue());
	    newSheet.getRow(invoiceIndex).getCell(3).setCellStyle(pi.getRow(2).getCell(0).getCellStyle());
	
	    newSheet.getRow(invoiceIndex+1).createCell(3).setCellValue("SHIPMENT DATE:");
	    newSheet.getRow(invoiceIndex+1).getCell(3).setCellStyle(pi.getRow(2).getCell(0).getCellStyle());
	
	    newSheet.getRow(invoiceIndex+2).createCell(3).setCellValue("CONTAINER/TRAILER ID:");
	    newSheet.getRow(invoiceIndex+2).getCell(3).setCellStyle(pi.getRow(2).getCell(0).getCellStyle());
	
	    newSheet.getRow(invoiceIndex+3).createCell(3).setCellValue("SEAL NUMBER:");
	    newSheet.getRow(invoiceIndex+3).getCell(3).setCellStyle(pi.getRow(2).getCell(0).getCellStyle());

	    newSheet.getRow(invoiceIndex+4).createCell(3).setCellValue("TOTAL GROSS WEIGHT:");
	    newSheet.getRow(invoiceIndex+4).getCell(3).setCellStyle(pi.getRow(2).getCell(0).getCellStyle());

	    newSheet.getRow(invoiceIndex+5).createCell(3).setCellValue("TOTAL CARTONS:");
	    newSheet.getRow(invoiceIndex+5).getCell(3).setCellStyle(pi.getRow(2).getCell(0).getCellStyle());

	    newSheet.getRow(2).createCell(9);

	    try {
		BufferedOutputStream bos = new BufferedOutputStream(new FileOutputStream(file));
		wb.write(bos);
	    } catch (Exception ex) {
		MasterLog.appendError(ex);
	    }
	}//if
    }//createCI()

    public void populateCO(String sc) {
	if (wb != null) {
	    XSSFSheet clientSheet;	
	    try {
		clientSheet = new XSSFWorkbook(new BufferedInputStream(new FileInputStream("ADMIN/OTHER/maven/src/main/resources/file/ClientDatabase.xlsx"))).getSheet("DB-Customers");
	    } catch (Exception ex) {
		MasterLog.appendError(ex);
		return;
	    }
	    XSSFSheet coSheet = wb.getSheet("CO");
	    XSSFSheet piSheet = wb.getSheet("PI");
	    XSSFRow clientRow;
	    XSSFCell coCell;
	    int columnIndex = CreateDocuments.findColumnIndex(clientSheet, OrderLog.findClient(OrderLog.findOrderEntry(sc)), 0);

	    //only populate if it's empty
	    if ((coSheet.getRow(10).getCell(3) == null) || coSheet.getRow(10).getCell(3).getStringCellValue().equals("")) {
		coSheet.getRow(CreateDocuments.findRowIndex(coSheet, "DATE:", 0)).createCell(3)
		    .setCellFormula("'CI-PL'!J3");
	
		//CONSIGNEE POPULATE
		clientRow = clientSheet.getRow(1); //first line of Consignee
		coCell = coSheet.getRow(10).createCell(3);
		while (clientRow.getCell(columnIndex) != null) { //iterates until reaching an empty cell
		    if (clientRow.getCell(columnIndex).getStringCellValue().equals("")) //if not null but is empty
			break;
	    
		    coSheet.shiftRows(coCell.getRowIndex() + 1, 100, 1); //shifts every row down one

		    coCell.setCellValue(clientRow.getCell(columnIndex).getStringCellValue());

		    coCell = coSheet.createRow(coCell.getRowIndex() + 1).createCell(3); //incrementing the row
		    if (clientSheet.getRow(clientRow.getRowNum() + 1) == null)
			clientSheet.createRow(clientRow.getRowNum() + 1);
		    clientRow = clientSheet.getRow(clientRow.getRowNum() + 1); //incrementing the row
	    
		}//while

	
		//NOTIFY POPULATE
		clientRow = clientSheet.getRow(CreateDocuments.findRowIndex(clientSheet, "Notify", 0));
		coCell = coSheet.getRow(CreateDocuments.findRowIndex(coSheet, "NOTIFY PARTY:", 0)).createCell(3);	
		while (clientRow.getCell(columnIndex) != null) { //iterates until reaching an empty cell
		    if (clientRow.getCell(columnIndex).getStringCellValue().equals("")) //if not null but is empty
			break;
	    
		    coSheet.shiftRows(coCell.getRowIndex() + 1, 100, 1); //shifts every row down one

		    coCell.setCellValue(clientRow.getCell(columnIndex).getStringCellValue());

		    coCell = coSheet.createRow(coCell.getRowIndex() + 1).createCell(3); //incrementing the row
		    if (clientSheet.getRow(clientRow.getRowNum() + 1) == null)
			clientSheet.createRow(clientRow.getRowNum() + 1);
		    clientRow = clientSheet.getRow(clientRow.getRowNum() + 1); //incrementing the row
		}//while

		String country = clientSheet.getRow(CreateDocuments.findRowIndex(clientSheet, "Notify", 0)).getCell(columnIndex)
		    .getStringCellValue();
		coSheet.getRow(CreateDocuments.findRowIndex(coSheet, "COUNTRY OF DESTINATION:", 0)).createCell(3).setCellValue(country);

		int scRowIndex = CreateDocuments.findRowIndex(piSheet, "SALES CONTRACT:", 0);
	
		coSheet.getRow(CreateDocuments.findRowIndex(coSheet, "INVOICE NUMBER:", 1)).createCell(3)
		    .setCellFormula("'CI-PL'!E" + (scRowIndex+1));
		coSheet.getRow(CreateDocuments.findRowIndex(coSheet, "CONTAINER #:", 1)).createCell(3)
		    .setCellFormula("'CI-PL'!E" + (scRowIndex+3));	
		coSheet.getRow(CreateDocuments.findRowIndex(coSheet, "CONTAINER SIZE:", 1)).createCell(3)
		    .setCellFormula("SUBSTITUTE('CI-PL'!B" + (scRowIndex+6) + ",\"1X\",\"\")");

		try {
		    BufferedOutputStream bos = new BufferedOutputStream(new FileOutputStream(file));
		    wb.write(bos);
		} catch (Exception ex) {
		    MasterLog.appendError(ex);
		}//try-catch
	    }//if
	}//if
    }//populateCO()
}
