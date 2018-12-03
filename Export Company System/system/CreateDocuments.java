/*
  Programmer: Patrick Nercessian

  Class Purpose:
     This class creates and populates a new Excel Workbook for each new order.
     It also contains helper methods to search through an Excel sheet.
 */

package system;

import java.io.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.usermodel.CellType;

import java.awt.Desktop;
import java.awt.Color;

import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.nio.file.Paths;
import java.nio.file.attribute.PosixFilePermissions;

import java.util.Date;
import java.util.Properties;
import java.text.DateFormat;
import java.text.SimpleDateFormat;

public class CreateDocuments {

    private File template = new File("ADMIN/OTHER/maven/src/main/resources/file/Purchase Order Template.xlsm");
    private final File clientDatabase = new File("ADMIN/OTHER/maven/src/main/resources/file/ClientDatabase.xlsx");
    private final File factoryDatabase = new File("ADMIN/OTHER/maven/src/main/resources/file/FactoryDatabase.xlsx");
    private File documentsFile;

    private static Properties prop = new Properties();

    private static XSSFWorkbook documentsWB;
    private static XSSFSheet salesContract;
    private static XSSFSheet purchaseOrder;
    private XSSFSheet calcSheet;

    private static XSSFWorkbook clientWB;
    private static XSSFSheet clientSheet;

    private static XSSFWorkbook factoryWB;
    private static XSSFSheet factorySheet;


    private String client;
    private String factory;
    private String modelLetters;
    private String incoterm;
    private String delivery;
    private String containerType;
    private int numContainers;

    private int desiredSC;
	
    private int id;


    /**
     * Creates necessary files, workbooks, and assigns instance varibales
     *
     * @param client the Client for this order
     * @param factory the Factory for this order
     * @param modelLetters the model for this order
     * @param incoterm the incoterm for this order
     * @param delivery the delivery destination for this order
     * @param containerType the size of the container
     * @param numContainers the number of containers
     * @param desiredSC the desired Sales Contract Number (-1 if it should be autoassigned)
     */
    public CreateDocuments(String client, String factory, String modelLetters, String incoterm, String delivery,
			   String containerType, int numContainers, int desiredSC, boolean longModel) {
	this.client = client;
	this.factory = factory;
	this.modelLetters = modelLetters;
	this.incoterm = incoterm;
	this.delivery = delivery;
	this.containerType = containerType;
	this.numContainers = numContainers;

	this.desiredSC = desiredSC;

	if (longModel) template = new File("ADMIN/OTHER/maven/src/main/resources/file/Purchase Order Template - Long.xlsm");

	
	//assigning worksbooks, sheets, and files
	try {
	    clientWB = new XSSFWorkbook(new FileInputStream(clientDatabase));
	    clientSheet = clientWB.getSheet("DB-Customers");


	    factoryWB = new XSSFWorkbook(new FileInputStream(factoryDatabase));
	    factorySheet = factoryWB.getSheet("DB-Factories");
	    
	    Date date = new Date();
	    SimpleDateFormat sdf = new SimpleDateFormat("MM-dd-yyyy");
	    documentsFile = new File("EXPORT HOME FURNISHINGS/" + this.client + "/" + this.factory + "/" + sdf.format(date) + " ERROR FILE");

	    
	    if (!documentsFile.exists())
		documentsFile.mkdirs();

	    Files.copy(template.toPath(), documentsFile.toPath(), StandardCopyOption.REPLACE_EXISTING); //creates goal file

	    documentsWB = new XSSFWorkbook(new FileInputStream(documentsFile));

	    salesContract = documentsWB.getSheet("PI");
	    purchaseOrder = documentsWB.getSheet("PO");
	    calcSheet = documentsWB.getSheet("CALC SHEET");
	    salesContract.lockObjects(false);	    
	    salesContract.disableLocking();
	    purchaseOrder.lockObjects(false);	    
	    purchaseOrder.disableLocking();
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	}//try-catch




    }//CreateDocuments()

    /** 
     * Populates the Consignee, Notify, Requirements, and Notes subsections. Also partially populates the Calc Sheet. Also loads the properties file.
     */
    public void populateWithClient() {
	XSSFRow clientRow = clientSheet.getRow(0);
	XSSFCell salesCell = salesContract.getRow(2).createCell(1);
	int columnIndex = findColumnIndex(clientSheet, this.client, 0);

	//CONSIGNEE POPULATE
	clientRow = clientSheet.getRow(1); //first line of Consignee
	while (clientRow.getCell(columnIndex) != null) { //iterates until reaching an empty cell
	    if (clientRow.getCell(columnIndex).getStringCellValue().equals("")) //if not null but is empty
		break;
	    
	    salesContract.shiftRows(salesCell.getRow().getRowNum() + 1, 200, 1); //shifts every row down one

	    salesCell.setCellValue(clientRow.getCell(columnIndex).getStringCellValue());

	    salesCell = salesContract.createRow(salesCell.getRow().getRowNum() + 1).createCell(1); //incrementing the row
	    if (clientSheet.getRow(clientRow.getRowNum() + 1) == null)
		clientSheet.createRow(clientRow.getRowNum() + 1);
	    clientRow = clientSheet.getRow(clientRow.getRowNum() + 1); //incrementing the row
	    
	}//while



	//NOTIFY POPULATE
	clientRow = clientSheet.getRow(findRowIndex(clientSheet, "Notify", 0));
	salesCell = salesContract.getRow(2).createCell(5);
	while (clientRow.getCell(columnIndex) != null) { //iterates until reaching an empty cell
	    if (clientRow.getCell(columnIndex).getStringCellValue().equals("")) //if not null but is empty
		break;
	    
	    if (salesCell.getRow() == null)
		salesContract.createRow(salesCell.getRowIndex());
	    if (salesCell.getRow().getCell(4) == null)
		salesCell.getRow().createCell(4);
			
	    if (!salesCell.getRow().getCell(4).getStringCellValue().equals("NOTIFY:")
		      && salesContract.getRow(salesCell.getRowIndex()).getCell(0) != null
		      && !salesContract.getRow(salesCell.getRowIndex()).getCell(0).getStringCellValue().equals("")) {
		salesContract.shiftRows(salesCell.getRowIndex(), 200, 1);
		salesCell = salesContract.createRow(salesCell.getRowIndex() - 1).createCell(5);
	    }

	    salesCell.setCellValue(clientRow.getCell(columnIndex).getStringCellValue());
	    if (salesContract.getRow(salesCell.getRow().getRowNum() + 1) == null)
		salesContract.createRow(salesCell.getRow().getRowNum() + 1);
	    
	    salesCell = salesContract.getRow(salesCell.getRow().getRowNum() + 1).createCell(5); //incrementing the row
	    clientRow = clientSheet.getRow(clientRow.getRowNum() + 1); //incrementing the row	    
	}//while


	//REQUIREMENTS POPULATE
	clientRow = clientSheet.getRow(findRowIndex(clientSheet, "Colums/Requirements", 0));
	try {
	    salesCell = salesContract.getRow(findRowIndex(salesContract, "MODEL", 0)).getCell(10);
	} catch (NullPointerException ex) {
	    salesCell = salesContract.getRow(findRowIndex(salesContract, "MODEL", 0)).createCell(10);
	}//try-catch
	while (clientRow.getCell(columnIndex) != null) { //iterates until reaching an empty cell
	    if (clientRow.getCell(columnIndex).getStringCellValue().equals("")) //if not null but is empty
		break;
	    
	    salesCell.setCellValue(clientRow.getCell(columnIndex).getStringCellValue());

	    if (salesContract.getRow(salesCell.getRowIndex()).getCell(salesCell.getColumnIndex() + 1) == null)
		salesCell = salesContract.getRow(salesCell.getRowIndex()).createCell(salesCell.getColumnIndex() + 1);
	    else
		salesCell = salesContract.getRow(salesCell.getRowIndex()).getCell(salesCell.getColumnIndex() + 1);	    
	    clientRow = clientSheet.getRow(clientRow.getRowNum() + 1); //incrementing the row
	}//while

	try {
	//NOTES POPULATE
	    XSSFCellStyle redStyle = documentsWB.createCellStyle();
	    XSSFFont redFont = documentsWB.createFont();
	    redFont.setColor((short)10);
	    redFont.setFontHeightInPoints((short)12);
	    redStyle.setFont(redFont);
	    
	clientRow = clientSheet.getRow(findRowIndex(clientSheet, "Notes", 0));
	salesCell = salesContract.getRow(findRowIndex(salesContract, "NOTES:", 6)).createCell(7);
	while (clientRow.getCell(columnIndex) != null) { //iterates until reaching an empty cell
	    if (clientRow.getCell(columnIndex).getStringCellValue().equals("")) //if not null but is empty
		break;

	    try {
		if (salesContract.getRow(salesCell.getRowIndex() + 2).getCell(salesCell.getColumnIndex()).getStringCellValue().equals("QTY"))
		    salesContract.shiftRows(salesCell.getRowIndex() + 2, 200, 1); //shifts every row down one	    
	    } catch (NullPointerException npe) {
		//		MasterLog.appendEntry("npe at " + salesCell.getRowIndex() + 2);
	    }
	    String note = clientRow.getCell(columnIndex).getStringCellValue();
	    if (!note.endsWith("(RED)")) {
		salesCell.setCellValue(note);
	    } else {
		salesCell.setCellValue(note.substring(0, note.indexOf("(RED)")));
		salesCell.setCellStyle(redStyle);
	    }
	    

	    if (salesContract.getRow(salesCell.getRow().getRowNum() + 1) == null)
		salesContract.createRow(salesCell.getRow().getRowNum() + 1);	    

	    salesCell = salesContract.getRow(salesCell.getRowIndex() + 1).createCell(7); //incrementing the row	    
	    clientRow = clientSheet.getRow(clientRow.getRowNum() + 1); //incrementing the row
	}//while
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	}//try-catch

	clientRow = clientSheet.getRow(findRowIndex(clientSheet, "Port of Discharge", 0));
	try {
	    salesContract.getRow(findRowIndex(salesContract, "PORT OF DISCHARGE:", 0)).createCell(1).setCellValue(clientRow.getCell(columnIndex).getStringCellValue());
	} catch (NullPointerException npe) {}

	//COUNTRY ON CALC SHEET
	String country = clientSheet.getRow(findRowIndex(clientSheet, "Country", 0)).getCell(columnIndex).getStringCellValue();
	String rep = clientSheet.getRow(findRowIndex(clientSheet, "EHF Sales Rep", 0)).getCell(columnIndex).getStringCellValue();
	calcSheet.disableLocking();
	calcSheet.getRow(findRowIndex(calcSheet, "REP/COUNTRY:", 0)).getCell(1).setCellValue(rep + " / " + country);
	calcSheet.enableLocking();

	try {
	    File file = new File("ADMIN/OTHER/id.properties");
	    if (file.createNewFile())
		prop.setProperty("id", "0");
	    else
		prop.load(new FileInputStream("ADMIN/OTHER/id.properties"));
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	}//try-catch

	

    }//populateWithClient()

    /**
     * Populates remaining information (e.g. date, sales contract number, factory, payment terms, etc.)
     */
    public void autoPopulate() {
	Date date = new Date();
	DateFormat year = new SimpleDateFormat("yyyy");
	String[] monthNames = {"January", "February", "March", "April", "May", "June", "July", "August","September",
			       "October", "November", "December"};
	int rowIndex = findRowIndex(salesContract, "SALES CONTRACT:", 0);
	String paymentTerms = clientSheet.getRow(findRowIndex(clientSheet, "Payment", 0))
	    .getCell(findColumnIndex(clientSheet, this.client, 0)).getStringCellValue();

	if (this.desiredSC == -1) { //normal increment
	    int tempId = Integer.parseInt(prop.getProperty("id")) + 1;
	    salesContract.getRow(rowIndex++).createCell(1).setCellValue(tempId + " / " + monthNames[date.getMonth()]
								    + " " + year.format(date)); //SC#, Month, Year
	} else { //custom sc
	    salesContract.getRow(rowIndex++).createCell(1).setCellValue(desiredSC + " / " + monthNames[date.getMonth()]
									+ " " + year.format(date)); //SC#, Month, Year
	}//if-else
	
	rowIndex++; //skip TAG OR PO# because it is user input
	salesContract.getRow(rowIndex++).createCell(1).setCellValue(this.factory);
	salesContract.getRow(rowIndex++).createCell(1).setCellValue(this.incoterm + " " + this.delivery);
	salesContract.getRow(rowIndex++).createCell(1).setCellValue(paymentTerms);
	salesContract.getRow(rowIndex++).createCell(1).setCellValue(this.numContainers + "X" + this.containerType);

    }//autoPopulate()

    /**
     * Fills in the Date Cell
     */
    public void fillDate() {
	DateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy");
        Date date = new Date();
	salesContract.getRow(2).createCell(9).setCellValue(dateFormat.format(date));

	//create each cell under Date to remove Currency
	for (int i = 3; i < findRowIndex(salesContract, "AMOUNT", 9); i++)
	    salesContract.getRow(i).createCell(9);

    }//fillDate()

    
    /**
     * Fills in Total/Amount Requirements, fills in Totals Row, fills in Freight Row and Subtotal Row if needed. Formulas are generated for each.
     */
    public void fillTotals() {
	try {
	    XSSFRow requireRow = salesContract.getRow(findRowIndex(salesContract, "MODEL", 0));
	    XSSFRow totalsRow = salesContract.getRow(findRowIndex(salesContract, "TOTAL:", 5));
	    int qtyCol = findColumnIndex(salesContract, "QTY", requireRow.getRowNum());
	    String firstWord;

	    XSSFCellStyle dollarStyle = documentsWB.createCellStyle();
	    XSSFDataFormat df = documentsWB.createDataFormat();
	    dollarStyle.setDataFormat(df.getFormat("$#,#0.00"));

	    int startCol = findColumnIndex(salesContract, "AMOUNT", findRowIndex(salesContract, "MODEL", 0));
	
	    //totalRow adds subtotal and freight, then becomes subtotal row (rest of the method)
	    if (this.incoterm.equalsIgnoreCase("C&F") || this.incoterm.equalsIgnoreCase("CIF")) {
		totalsRow.getCell(startCol).setCellFormula("SUM(" + Character.toUpperCase((char)(startCol + 97))
							   + (totalsRow.getRowNum() - 1) + ","
							   + Character.toUpperCase((char)(startCol + 97)) + (totalsRow.getRowNum())
							   + ")");

		totalsRow = salesContract.getRow(findRowIndex(salesContract, "TOTAL:", 5) - 2);
		if (totalsRow == null) salesContract.createRow(totalsRow.getRowNum());
		totalsRow.createCell(5).setCellValue("SUBTOTAL:");
		salesContract.createRow(totalsRow.getRowNum() + 1).createCell(5).setCellValue("FREIGHT:");	    
	    }//if
	
	
	    for (int i = startCol; i < requireRow.getLastCellNum(); i++) { //iterates through client-specific requirements
	    
		//if requirement is a Total or Amount
		firstWord = requireRow.getCell(i).getStringCellValue().split("\n")[0];
		if (firstWord.equalsIgnoreCase("total") || firstWord.equalsIgnoreCase("amount")) {

		    boolean a = firstWord.equalsIgnoreCase("amount");
		    for (int k = requireRow.getRowNum() + 1; k < totalsRow.getRowNum(); k++) { //going through each model
			if (salesContract.getRow(k) == null)
			    salesContract.createRow(k);

			salesContract.getRow(k).createCell(i);
			if (a)
			    salesContract.getRow(k).getCell(i).setCellStyle(dollarStyle);
			salesContract.getRow(k).getCell(i).setCellFormula("IF("
									  + Character.toUpperCase((char)(qtyCol + 97)) + (k + 1) + "*"
									  + Character.toUpperCase((char)(i + 96)) + (k + 1) + "=0,\"\","
									  + Character.toUpperCase((char)(qtyCol + 97)) + (k + 1) + "*"
									  + Character.toUpperCase((char)(i + 96)) + (k + 1) +")"); //multiplying QTY and Unit
		    }//for
		    if (totalsRow.getCell(i) == null)
			totalsRow.createCell(i);
		    totalsRow.getCell(i).setCellFormula("SUM(" + Character.toUpperCase((char)(i + 97)) + (requireRow.getRowNum() + 2)
							+ ":" + Character.toUpperCase((char)(i + 97)) + (totalsRow.getRowNum()) + ")");
		}//if

		if (i == startCol) //goes back to normal (total is total row [not subtotal row])
		    totalsRow = salesContract.getRow(findRowIndex(salesContract, "TOTAL:", 5));
	    }//for
	} catch (Exception ex) { 
	    MasterLog.appendError(ex);
	}//try-catch
	
    }//fillTotals()

    /**
     * Reserve a specified number of Sales Contract IDs.
     *
     * @param num the number of IDs to reserve
     */
    protected static void reserveIDs(int num) {
	try {
	    prop.load(new FileInputStream("ADMIN/OTHER/id.properties"));
	    prop.setProperty("id", String.valueOf(Integer.parseInt(prop.getProperty("id")) + num));
	    prop.store(new FileOutputStream("ADMIN/OTHER/id.properties"), null);
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	}
    }//reserveIDs(int)

    /**
     * Get the current Sales Contract Number
     *
     * @return the current Sales Contract Number
     */
    public static int getCurrentId() {
	int id = -1;
	try {
	    prop.load(new FileInputStream("ADMIN/OTHER/id.properties"));
	    id = Integer.parseInt(prop.getProperty("id"));
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	}
	return id;
    }

    /**
     * finds desired column based on 'str' and 'row'
     *
     * @param sheet the XSSFSheet to be searched
     * @param str the String to be found
     * @param row the row to be searched
     * @return the column index
     */
    public static int findColumnIndex(XSSFSheet sheet, String str, int row) {
	int lastCol = sheet.getRow(row).getLastCellNum();
	for (int col = 0; col <= lastCol; col++) {
	    //	    if (sheet.getRow(row).getCell(col) == null)
	    //		sheet.getRow(row).createCell(col);
	    try {
		if (sheet.getRow(row).getCell(col).getStringCellValue().trim().equalsIgnoreCase(str.trim()))
		    return col;
	    } catch (NullPointerException npe) { }
	}//for
	return -1; //if not found
    }//findColumnIndex(XSSFSheet, String, int)

    /**
     * finds desired row based on 'str' and 'col'
     *
     * @param sheet the XSSFSheet to be searched
     * @param str the string to be found
     * @param col the column to be searched
     * @return the row index
     */
    public static int findRowIndex(XSSFSheet sheet, String str, int col) {
	for (int row = 0; row <= sheet.getLastRowNum(); row++) {
	    if (sheet.getRow(row) == null)
		sheet.createRow(row);
	    if (sheet.getRow(row).getCell(col) == null)
		sheet.getRow(row).createCell(col);
	    XSSFCell cell = sheet.getRow(row).getCell(col);

	    try {
		if (cell.getRichStringCellValue().getString().trim().equalsIgnoreCase(str))
		    return row;
	    } catch (Exception ex) {}
		
	}//for
	return -1; //if not found
    }//findColumnIndex(XSSFSheet, String, int)


    /** 
     * Opens the the Workbook.
     */
    public void openSheet() {
	try {
	    Desktop.getDesktop().open(documentsFile);
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	}
    }//openSheet()



    /**
     * Populates the Purchase Order sheet with factory info, order info, and sets formulas to bring over models from the PI.
     * Sloppy by necessity: Apache Poi cell styles and formulas are difficult to write clean code for.
     */
    public void populatePO() {
	int rowNum = 0;
	double discount;
	XSSFRow cFactory = factorySheet.getRow(rowNum);
	
	while (!cFactory.getCell(0).getStringCellValue().equalsIgnoreCase(this.factory)) {
	    cFactory = factorySheet.getRow(++rowNum);
	    if (cFactory == null)
		return; //don't populate Purchase Order if factory is not in database
	}

	purchaseOrder.getRow(2).createCell(1).setCellValue(this.factory); //factory name
	purchaseOrder.getRow(3).createCell(1).setCellValue(cFactory.getCell(1).getStringCellValue()); //address 1
	purchaseOrder.getRow(4).createCell(1).setCellValue(cFactory.getCell(2).getStringCellValue()); //address 2

	purchaseOrder.getRow(5).createCell(0).setCellValue("CONTACT:");
	purchaseOrder.getRow(5).createCell(1).setCellValue(cFactory.getCell(3).getStringCellValue()); //contact

	//	purchaseOrder.getRow(3).createCell(4).setCellValue(this.numContainers + "X" + this.containerType); //container
	purchaseOrder.getRow(3).createCell(4).setCellFormula("IF(container<>\"\",container,\"\")");


	int poModelIndex = findRowIndex(purchaseOrder, "MODEL", 0);
	int poSubTotalIndex = findRowIndex(purchaseOrder, "SUBTOTAL:", 3);
	int piModelIndex = findRowIndex(salesContract, "MODEL", 0);
	int piTotalIndex = findRowIndex(salesContract, "SUBTOTAL:", 5);
	if (piTotalIndex == -1) piTotalIndex = findRowIndex(salesContract, "TOTAL:", 5);

	XSSFCell cell;
	for (int i = 0; i < poSubTotalIndex - (poModelIndex + 1); i++) {
	    for (int k = 0; k < findColumnIndex(purchaseOrder, "FABRIC/COLOR", findRowIndex(purchaseOrder, "MODEL", 0)); k++) {
		cell = purchaseOrder.getRow(i + (poModelIndex + 1)).createCell(k);
		try {
		    cell.setCellStyle(salesContract.getRow(piModelIndex + 1).getCell(k).getCellStyle());
		} catch (Exception ex) {}
		cell.setCellFormula("IF(PI!" + ((char)(k + 65)) + (i + (piModelIndex + 2)) + "<>\"\","
				    + "PI!" + ((char)(k + 65)) + (i + (piModelIndex + 2)) + ",\"\")");
	    }//for
	    int k = findColumnIndex(salesContract, "FABRIC/COLOR", findRowIndex(salesContract, "MODEL", 0));
	    cell = purchaseOrder.getRow(i + (poModelIndex + 1))
		.createCell(findColumnIndex(purchaseOrder, "FABRIC/COLOR", findRowIndex(purchaseOrder, "MODEL", 0)));
	    try {
		cell.setCellStyle(salesContract.getRow(piModelIndex + 1).getCell(findColumnIndex(salesContract, "FABRIC/COLOR", findRowIndex(salesContract, "MODEL", 0))).getCellStyle());
	    } catch (Exception ex) {}	    
	    cell.setCellFormula("IF(PI!" + ((char)(k + 65)) + (i + (piModelIndex + 2)) + "<>\"\","
				+ "PI!" + ((char)(k + 65)) + (i + (piModelIndex + 2)) + ",\"\")");


	    k = findColumnIndex(salesContract, "QTY", findRowIndex(salesContract, "MODEL", 0));
	    cell = purchaseOrder.getRow(i + (poModelIndex + 1))
		.createCell(findColumnIndex(purchaseOrder, "QTY", findRowIndex(purchaseOrder, "MODEL", 0)));
	    try {
		cell.setCellStyle(salesContract.getRow(piModelIndex + 1).getCell(findColumnIndex(salesContract, "QTY", findRowIndex(salesContract, "MODEL", 0))).getCellStyle());
	    } catch (Exception ex) {}	    
	    cell.setCellFormula("IF(PI!" + ((char)(k + 65)) + (i + (piModelIndex + 2)) + "<>\"\","
				+ "PI!" + ((char)(k + 65)) + (i + (piModelIndex + 2)) + ",\"\")");
	}//for
	
	//setting the value of Discount cell
	try { 
	    discount = cFactory.getCell(findColumnIndex(factorySheet, "Discount", 0)).getNumericCellValue();
	    purchaseOrder.getRow(findRowIndex(purchaseOrder, "DISCOUNT:", 3)).getCell(7).setCellValue(discount);
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	}
    }//populatePO()

    /**
     * Populates the Very Hidden Email Template sheet based on factory database
     */
    public void populateEmail() {
	String to, cc, bcc, firstName;
	
	XSSFSheet emailSheet = documentsWB.getSheet("EMAIL TEMPLATE");

	emailSheet.disableLocking();	
	int rowNum = 1;
	XSSFRow factoryRow = factorySheet.getRow(rowNum++);
	
	while (factoryRow != null) {
	    if (factoryRow.getCell(0).getStringCellValue().equals(this.factory)) {
		to = factoryRow.getCell(findColumnIndex(factorySheet, "TO:", 0)).getStringCellValue();
		cc = factoryRow.getCell(findColumnIndex(factorySheet, "CC:", 0)).getStringCellValue();
		bcc = factoryRow.getCell(findColumnIndex(factorySheet, "BCC:", 0)).getStringCellValue();
		firstName = factoryRow.getCell(findColumnIndex(factorySheet, "CONTACT", 0)).getStringCellValue().split(" ")[0];
		if (firstName.length() != 0)
		    firstName = firstName.substring(0, 1) + firstName.substring(1).toLowerCase(); //first letter capitalized
		
		emailSheet.getRow(findRowIndex(emailSheet, "TO:", 2)).createCell(3).setCellValue(to);
		emailSheet.getRow(findRowIndex(emailSheet, "CC:", 2)).createCell(3).setCellValue(cc);
		emailSheet.getRow(findRowIndex(emailSheet, "BCC:", 2)).createCell(3).setCellValue(bcc);
		
		emailSheet.getRow(findRowIndex(emailSheet, "LINE 1", 2)).createCell(3).setCellValue("Dear " + firstName + ",");
		break;
	    }
	    
	    factoryRow = factorySheet.getRow(rowNum++); //incrementing row
	}//while
    }//populateEmail()

    /**
     * Finishes populating the Calculation Sheet
     */
    public void populateCalcSheet() {
	try {
	    calcSheet.disableLocking();	
	    calcSheet.getRow(findRowIndex(calcSheet, "S/C #", 0)).getCell(1).setCellFormula("IF(SCWITHDATE<>\"\",SCWITHDATE,\"\")");
	    calcSheet.getRow(findRowIndex(calcSheet, "CUSTOMER:", 0)).getCell(1).setCellValue(this.client);
	    calcSheet.getRow(findRowIndex(calcSheet, "CL REF#:", 0)).getCell(1).setCellFormula("IF(customerPO<>\"\",customerPO,\"\")");
	    calcSheet.getRow(findRowIndex(calcSheet, "PMT TERMS:", 0)).getCell(1).setCellFormula("IF(PaymentTerms<>\"\",PaymentTerms,\"\")");
	    calcSheet.getRow(findRowIndex(calcSheet, "SHIP TERMS:", 0)).getCell(1).setCellFormula("IF(DeliveryTerms<>\"\",DeliveryTerms,\"\")");
	    
	    calcSheet.getRow(findRowIndex(calcSheet, "S/C #", 0)).getCell(6).setCellFormula("IF(PO!E3<>\"\",PO!E3,\"\")");
	    calcSheet.getRow(findRowIndex(calcSheet, "S/C #", 0)).getCell(9).setCellFormula("IF(PO!I3<>\"\",PO!I3,\"\")");	    
	    calcSheet.getRow(findRowIndex(calcSheet, "CUSTOMER:", 0)).getCell(7).setCellValue(this.factory);
	    calcSheet.getRow(findRowIndex(calcSheet, "CL REF#:", 0)).getCell(6).setCellValue(this.modelLetters);
	    calcSheet.getRow(findRowIndex(calcSheet, "REP/COUNTRY:", 0)).getCell(6).setCellFormula("IF(container<>\"\",container,\"\")");
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	}//try-catch
	calcSheet.enableLocking();	
    }//populateCalcSheet()

    /**
     * Writes to the file, then checks that it was successfully written to. If so, it renames the file from Error File to correct name and creates Log Entry
     */
    public void writeFile() {
	FileOutputStream fos = null;
	BufferedOutputStream bos = null;
	try {
	    fos = new FileOutputStream(documentsFile);
	    bos = new BufferedOutputStream(fos);
	    documentsWB.write(bos);
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	} finally {
	    try {
		documentsWB.close();
		if (bos != null) { bos.flush(); bos.close(); }
		if (fos != null) { fos.flush(); fos.close(); }
	    } catch (IOException ioe) {
		MasterLog.appendError(ioe);
	    }//try-catch
	}//try-catch-finally
	
	if (CheckCompletion.checkCreated(documentsWB)) {
	    
	    //incrementing SC# (creating property file if nonexistant)
	    try {
		if (desiredSC == -1) { //normal increment
		    prop.setProperty("id", String.valueOf(Integer.parseInt(prop.getProperty("id")) + 1));//increments id
		    this.id = Integer.parseInt(prop.getProperty("id"));
		} else { //custom sc
		    this.id = desiredSC;
		}
		MasterLog.append("(SC#" + id + ")"); //appended to MenuApp entry
		prop.store(new FileOutputStream("ADMIN/OTHER/id.properties"), null);

		//renaming file
		Date date = new Date();
		DateFormat year = new SimpleDateFormat("yyyy");
		String[] monthNames = {"January", "February", "March", "April", "May", "June", "July", "August","September",
				       "October", "November", "December"};
		String name = this.id + " - model " + this.modelLetters + " - " + monthNames[date.getMonth()]
		    + " " + year.format(date) + ".xlsm";
		Files.move(documentsFile.toPath(), documentsFile.toPath().resolveSibling(name));
		documentsFile = new File("EXPORT HOME FURNISHINGS/" + this.client + "/" + this.factory + "/" + name);

		//LOGGING CREATION
		SimpleDateFormat sdf = new SimpleDateFormat("MM/dd/yyyy HH:mm:ss");
		String logEntry = sdf.format(documentsFile.lastModified()) + " - " + System.getProperty("user.name") + " - SC#: "
		    + this.id + " - Client: " + client + " - Factory: " + factory + " - Model: " + modelLetters + " - Updated: FALSE" + "            PENDING";
		Log.updateLog(new File("ADMIN/OTHER/maven/src/main/resources/file/Log - Sales Contract.txt"), logEntry + "\n\n");
		Log.updateExcelLog(logEntry);
		
	    } catch (Exception ex) {
		MasterLog.appendError(ex);
	    }//try-catch
	}//if
    }//writeFile()

}//CreateDocuments
