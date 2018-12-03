/*
  Programmer: Patrick Nercessian

  Class Purpose:
     This class contains code pertaining to the Distribution Log, a massive
     log containing information on each sale of a model
 */

package system;

import java.io.*;
import org.apache.poi.xssf.usermodel.*;

import java.awt.Desktop;

import java.nio.file.Files;
import java.nio.file.StandardCopyOption;

import java.util.Date;
import java.util.Arrays;
import java.util.ArrayList;

import java.text.SimpleDateFormat;

public class Distribution {
    private static File file = new File("ADMIN/OTHER/maven/src/main/resources/file/DistributionLog.xlsx");
    private static File logFile = new File("ADMIN/OTHER/maven/src/main/resources/file/Log - Sales Contract.txt");
    private static XSSFWorkbook workbook;    
    private static XSSFSheet sheet;

    private static int i, column;

    private static int findFirstIndex(String scStr) {
	int start = 1;
	int mid;
	int end = sheet.getLastRowNum();
	
	int desiredIndex = -1;
	String searchScStr;
	int searchScInt;
	int compareto;
	
	while (start <= end) {
	    mid = (start + end) / 2;
	    try {
		searchScInt = (int) sheet.getRow(mid).getCell(10).getNumericCellValue();
		if (Integer.parseInt(scStr) > searchScInt) compareto = 1;
		else  if (Integer.parseInt(scStr) < searchScInt) compareto = -1;
		else compareto = 0;
	    } catch (Exception ex) {
		//		MasterLog.appendEntry(ex.toString());
		searchScStr = sheet.getRow(mid).getCell(0).getStringCellValue();
		compareto = scStr.compareTo(searchScStr);
	    }
	    if (compareto == 0) {
		desiredIndex = mid;
		break;
	    } else if (compareto < 0) {
		end = mid - 1;
	    } else {
		start = mid + 1;
	    }//if-elseif-else

	    if (start > end) {
		if (end == -1)
		    desiredIndex = 0;
		else if (scStr.compareTo("" + sheet.getRow(end).getCell(10).getNumericCellValue()) < 0)
		    desiredIndex = end;
		else
		    desiredIndex = end + 1;
	    }//if
	}//while

	try {
	    while (sheet.getRow(desiredIndex - 1).getCell(10).getNumericCellValue() == Integer.parseInt(scStr))
		desiredIndex--;
	} catch (IllegalStateException ise) {}
	return desiredIndex;
    }
    
    /**
     * Adds models to Distribution Log
     *
     * @param infoArray the 2D String Array to add the models to
     * @param scStr the Sales Contract number
     */
    public static void addModels(String[][] infoArray, String scStr) {
	MasterLog.appendEntry("Attempting to add " + infoArray.length + " models for SC#" + scStr + "...");
	try {
	    workbook = new XSSFWorkbook(new BufferedInputStream(new FileInputStream(file)));
	    sheet = workbook.getSheet("DISTRIBUTION");
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	}//try-catch
	
	int modelCol = CreateDocuments.findColumnIndex(sheet, "MODEL#", 0);
	int compCol = CreateDocuments.findColumnIndex(sheet, "COMPOSITION", 0);
	int qtyCol = CreateDocuments.findColumnIndex(sheet, "QTY/SETS", 0);
	
	int desiredIndex = findFirstIndex(scStr);
	
	sheet.disableLocking();
	String orderEntry = OrderLog.findOrderEntry(scStr);
	String client = OrderLog.findClient(orderEntry);
	XSSFCell cell;
	if (desiredIndex != -1) {
	    try {
		XSSFWorkbook workbook = OrderLog.getWB(scStr);		
		sheet.shiftRows(desiredIndex, sheet.getLastRowNum(), infoArray.length);
	
		for (int i = 0; i < infoArray.length; i++) {
		    sheet.createRow(desiredIndex + i).createCell(0).setCellValue(OrderLog.findFactory(orderEntry));
	    
		    cell = sheet.getRow(desiredIndex + i).createCell(1);
		    cell.setCellValue(infoArray[i][0]);
		    cell.setCellStyle(sheet.getRow(1).getCell(1).getCellStyle());
		    
		    cell = sheet.getRow(desiredIndex + i).createCell(2);
		    cell.setCellValue(infoArray[i][1]);
		    cell.setCellStyle(sheet.getRow(1).getCell(2).getCellStyle());		    
		    
		    cell = sheet.getRow(desiredIndex + i).createCell(3);
		    cell.setCellValue(infoArray[i][2]);
		    cell.setCellStyle(sheet.getRow(1).getCell(3).getCellStyle());
		    
		    cell = sheet.getRow(desiredIndex + i).createCell(4);
		    try {
			cell.setCellValue(Integer.parseInt(infoArray[i][3]));
		    } catch (NumberFormatException nfe) {
			cell.setCellValue(infoArray[i][3]);
		    }
		    cell.setCellStyle(sheet.getRow(1).getCell(4).getCellStyle());		    

		    cell = sheet.getRow(desiredIndex + i).createCell(5);
		    cell.setCellValue(client);
		    cell.setCellStyle(sheet.getRow(1).getCell(5).getCellStyle());		    
		    
		    cell = sheet.getRow(desiredIndex + i).createCell(6);
		    cell.setCellValue(OrderLog.findCountry(client));
		    cell.setCellStyle(sheet.getRow(1).getCell(6).getCellStyle());

		    cell = sheet.getRow(desiredIndex + i).createCell(7);
		    cell.setCellValue(OrderLog.findPoDate(workbook));
		    cell.setCellStyle(sheet.getRow(1).getCell(7).getCellStyle());

		    cell = sheet.getRow(desiredIndex + i).createCell(8);
		    cell.setCellValue(OrderLog.findShipDate(workbook));
		    cell.setCellStyle(sheet.getRow(1).getCell(8).getCellStyle());
		    
		    
		    cell = sheet.getRow(desiredIndex + i).createCell(9);
		    cell.setCellValue(Integer.parseInt(OrderLog.findEhfNum(orderEntry)));
		    cell.setCellStyle(sheet.getRow(1).getCell(9).getCellStyle());
		    
		    cell = sheet.getRow(desiredIndex + i).createCell(10);
		    cell.setCellValue(Integer.parseInt(scStr));
		    cell.setCellStyle(sheet.getRow(1).getCell(10).getCellStyle());
		}//for
	    } catch (Exception ex) {
		MasterLog.appendError(ex);
	    }//try-catch	    
	}//if
	sheet.enableLocking();

	try {
	    BufferedOutputStream bos = new BufferedOutputStream(new FileOutputStream(file));
	    workbook.write(bos);
	    Log.updateModels(scStr);
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	}//try-catch
	MasterLog.append("Added Successfully");
    }//addModels(String[][])

    /**
     * Transforms the Excel Sheet's data into a String array
     *
     * @return the String array
     */
    public static String[] toArray() {
	try {
	    workbook = new XSSFWorkbook(new BufferedInputStream(new FileInputStream(file)));
	    sheet = workbook.getSheet("DISTRIBUTION");	    
	    String[] array = new String[sheet.getLastRowNum() - 1];
	    for (int r = 0; r < array.length; r++) {
		array[r] = " --- ";
		for (int c = 0; c < 11; c++) {
		    try {
			array[r] += sheet.getRow(r+1).getCell(c).getStringCellValue() + " --- ";
		    } catch (IllegalStateException ise) {
			array[r] += (int) sheet.getRow(r+1).getCell(c).getNumericCellValue() + " --- ";
		    } catch (NullPointerException npe) { }
		}//for
	    }//for
	    return array;	
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	}
	return new String[0];
    }//toArray()

    /**
     * Searches the 2D String array for anything containing the filters
     *
     * @param all a 1D String Array of all workbook entries
     * @param filters the filters to search through the data with
     */
    public static String[][] resultingSearch(String[] all, String ... filters) {
	all = Arrays.stream(all).map(str -> str.toUpperCase()).toArray(String[]::new);
	String[][] splitArray = toSplitArray(all);
	String[][] splitArray2 = toSplitArray(all);
	String[][] splitArray3 = toSplitArray(all);	
	column = -1;

	for (i = 0; i < filters.length - 2; i++) {
	    switch (i) {
	    case 0: column = 0;break;
	    case 1: column = 1;break;
	    case 2: column = 5;break;
	    case 3: column = 6;break;
		//	    case 4: column = 6;break;
		//	    case 5: column = 6;break;		
	    }//switch
	    if (!filters[i].equals(""))
		splitArray = Arrays.stream(splitArray).filter(arr -> arr[column].contains(filters[i].toUpperCase())).toArray(String[][]::new);
	}//for

	if (!filters[4].equals("")) { //if country2
	    for (i = 0; i < filters.length - 1; i++) {
		if (i == 3) continue;
		
		switch (i) {
		case 0: column = 0;break;
		case 1: column = 1;break;
		case 2: column = 5;break;
		    
		case 4: column = 6;break;
		}//switch
		if (!filters[i].equals(""))
		    splitArray2 = Arrays.stream(splitArray2).filter(arr -> arr[column].contains(filters[i].toUpperCase())).toArray(String[][]::new);
	    }//for
	    String[][] temp = new String[splitArray.length + splitArray2.length][splitArray[0].length];
	    MasterLog.appendEntry("temp size:" + temp.length);
	    for (int x = 0; x < splitArray.length; x++)
		temp[x] = splitArray[x];
	    for (int x = 0; x < splitArray2.length; x++)
		temp[splitArray.length + x] = splitArray2[x];
	    splitArray = temp;
	}//if
	try {
	if (!filters[5].equals("")) { //if country3
	    for (i = 0; i < filters.length; i++) {
		if (i == 3 || i == 4) continue;
		
		switch (i) {
		case 0: column = 0;break;
		case 1: column = 1;break;
		case 2: column = 5;break;
		    
		case 5: column = 6;break;
		}//switch
		if (!filters[i].equals(""))
		    splitArray3 = Arrays.stream(splitArray3).filter(arr -> arr[column].contains(filters[i].toUpperCase())).toArray(String[][]::new);
	    }//for
	    String[][] temp = new String[splitArray.length + splitArray3.length][splitArray[0].length];
	    for (int x = 0; x < splitArray.length; x++)
		temp[x] = splitArray[x];
	    for (int x = 0; x < splitArray3.length; x++)
		temp[splitArray.length + x] = splitArray3[x];
	    splitArray = temp;
	}//if
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	}
	
	return splitArray;
    }//resultingSearching(String[]
    
    /**
     * Searches through the data based on either Ship Date or PO Date
     *
     * @param results the 1D array of data
     * @param minDate the minimum date
     * @param maxDate the maximum date
     * @param isPoDate determines whether PO Date or Ship Date should be searched
     */
    public static String[] resultingDateSearch(String[] results, Date minDate, Date maxDate, boolean isPoDate) {
	try {
	    results = Arrays.stream(results).filter(str -> {
		    String dateStr;
		    if (isPoDate)
			dateStr = str.substring(specificIndexOf(str, " --- ", 8) + 5, specificIndexOf(str, " --- ", 9));
		    else
			dateStr = str.substring(specificIndexOf(str, " --- ", 9) + 5, specificIndexOf(str, " --- ", 10));

		    if (isPoDate) {
			if (dateStr.trim().equals("") || dateStr.trim().equals("N/A"))
			    return true;
		    } else {
			if (dateStr.trim().equals("") || dateStr.trim().equals("N/A"))
			    return false;
		    }

			int day = Integer.parseInt(dateStr.substring(3, 5));
			int month = Integer.parseInt(dateStr.substring(0, 2)) - 1;
			int year = Integer.parseInt(dateStr.substring(6)) - 1900;
			Date date = new Date(dateStr);

			if (date.compareTo(minDate) < 0 || date.compareTo(maxDate) > 0)
			    return false;

		    return true;
		}).toArray(String[]::new);
	    return results;
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	}
	return new String[0];
    }//resultingDateSearch(String[], Date, Date)

    public static String[][] toSplitArray(String[] results) {
	String[][] arr = new String[results.length][11];
	for (int i = 0; i < results.length; i++) {
	    for (int x = 0; x < 11; x++) {
		int index1 = specificIndexOf(results[i], "---", x+1) + 4;
		int index2 = specificIndexOf(results[i], "---", x+2) - 1;
		try {
		    arr[i][x] = results[i].substring(index1, index2);
		} catch (Exception ex) {
		    MasterLog.appendError(ex);
		    MasterLog.appendEntry(results[i]);
		    MasterLog.appendEntry("" + index1);
		    MasterLog.appendEntry("" + index2);
		}
	    }//for
	}//for
	return arr;
    }//toSplitArray(String[])

    public static void exportExcel(String[][] finalResults) {
	Date date = new Date();
	SimpleDateFormat sdf = new SimpleDateFormat("MM-dd-yyyy hh mm ss");
	try {
	    File resultFile = new File("ADMIN/OTHER/maven/src/main/resources/file/Distribution Search Results/" + sdf.format(date) + " Search Result.xlsm");
	    if (!resultFile.exists())
		resultFile.mkdirs();

	    //these lines are needed because it makes a directory otherwise for some weird reason
	    File template = new File("ADMIN/OTHER/maven/src/main/resources/file/Empty Excel.xlsm");
	    Files.copy(template.toPath(), resultFile.toPath(), StandardCopyOption.REPLACE_EXISTING);

	    
	    XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(resultFile));
	    wb.setSheetName(0, "Result");
	    XSSFSheet sheet = wb.getSheet("Result");

	    String[] headers = {"FACTORY", "MODEL #", "COMPOSITION", "FABRIC/FINISH", "QTY/SETS", "CLIENT", "COUNTRY",
				"PO DATE", "SHIP DATE", "EHF#", "S/C #"};
	    XSSFRow row = sheet.createRow(0);
	    for (int c = 0; c < headers.length; c++)
		row.createCell(c).setCellValue(headers[c]);
	    for (int r = 0; r < finalResults.length; r++) {
		row = sheet.createRow(r+1);
		for (int c = 0; c < finalResults[r].length; c++) {
		    row.createCell(c).setCellValue(finalResults[r][c]);
		}//for
	    }//for
	    
	    for (int c = 0; c < finalResults[0].length; c++)
		sheet.autoSizeColumn(c);

	    wb.write(new FileOutputStream(resultFile));

	    Desktop.getDesktop().open(resultFile);
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	}//try-catch	
    }//exportExcel(String[][])

    public static void refreshExcel() {
	try {
	    workbook = new XSSFWorkbook(new BufferedInputStream(new FileInputStream(file)));
	    sheet = workbook.getSheet("DISTRIBUTION");
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	}//try-catch
	sheet.disableLocking();
	ArrayList<String> list = Log.modelLackingSCs();
	outerLoop: for (String sc : list) {
	    int desiredIndex = findFirstIndex(sc);
	    XSSFRow row = sheet.getRow(desiredIndex);
	    try {
		while (row.getCell(10).getNumericCellValue() == Integer.parseInt(sc)) {
		    //		    MasterLog.appendEntry(org.apache.poi.ss.usermodel.CellType.forInt(row.getCell(8).getCellType()) + "");
		    if (row.getCell(8).getCellType() == 2) {
			row.createCell(17).setCellValue(OrderLog.findShipDateDate(OrderLog.getWB(sc)));
			//			String formula = row.getCell(8).getCellFormula();
			//			row.getCell(8).setCellFormula(formula);
		    } else {
			row.getCell(8).setCellValue(OrderLog.findShipDate(OrderLog.getWB(sc)));
		    }
		    
		    row = sheet.getRow(++desiredIndex);
			//		    MasterLog.appendEntry((desiredIndex - 1) + "");
			//		    break outerLoop;
		}//while
	    } catch (IllegalStateException ise) {
		MasterLog.appendEntry("ise");		
	    } catch (Exception ex) {
		MasterLog.appendError(ex);
	    }//try-catch
	}//for
	new XSSFFormulaEvaluator(workbook).evaluateAll();
	sheet.enableLocking();

	try {
	    BufferedOutputStream bos = new BufferedOutputStream(new FileOutputStream(file));
	    workbook.write(bos);
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	}//try-catch	
    }//refreshExcel()

    private static int specificIndexOf(String str, String substr, int n) {
	int i = str.indexOf(substr);
	while (--n > 0 && i != -1)
	    i = str.indexOf(substr, i + 1);
	return i;
    }//originalIndexOf(String, String, int)

}
