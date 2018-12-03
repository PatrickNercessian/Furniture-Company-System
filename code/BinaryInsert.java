/*
  Programmer: Patrick Nercessian

  Class Purpose:
     Contains only one method, sortedFileInsert, which inserts a String
     into a TextFile alphabetically through a modified Binary Search
 */

package system;

import java.util.Scanner;
import java.util.ArrayList;
import java.util.List;

import java.io.*;
import java.nio.file.attribute.*;

import java.lang.StackTraceElement;

public class BinaryInsert {
    /**
     * Inserts text into a file alphabetically
     *
     * @param oldFile the file to be altered
     * @param entry the text to insert
     */
    public static void sortedFileInsert(File oldFile, String entry){
	List<String> lines = new ArrayList<String>();
	try {
	    Scanner scan = new Scanner(oldFile);
	    String line;
	    while (scan.hasNextLine()) {
		line = scan.nextLine();
		if (!line.equals(""))
		    lines.add(line);
	    }//while
	} catch (FileNotFoundException fnfe) {
	    MasterLog.appendError(fnfe);
	}//try-catch
						 
	String[] arr = lines.toArray(new String[0]);

	int desiredIndex = -1;
	int start = 0, end = arr.length -1;
	int mid;
	while (start <= end) {
	    mid = (start + end) / 2;
	    if (entry.compareTo(arr[mid]) == 0) {
		desiredIndex = mid;
		break;		
	    } else if (entry.compareTo(arr[mid]) < 0) {
		end = mid - 1;
	    } else {
		start = mid + 1;
	    }

	    if (start > end) {
		if (end == -1)
		    desiredIndex = 0;
		else if (entry.compareTo(arr[end]) < 0)
		    desiredIndex = end;
		else
		    desiredIndex = end + 1;
	    }//if
	}//while
	
	String[] newArr = new String[arr.length + 1];
	for (int i = 0; i < desiredIndex; i++)
	    newArr[i] = arr[i];
	newArr[desiredIndex] = entry;
	for (int i = desiredIndex + 1; i < newArr.length; i++)
	    newArr[i] = arr[i-1];

	
	String newText = "";
	for (int i = 0; i < newArr.length; i++)
	    newText += newArr[i] + "\n";
	
	FileWriter fw = null;
	BufferedWriter bw = null;
	try {
	    fw = new FileWriter(oldFile, false);
	    bw = new BufferedWriter(fw);
	    bw.write(newText);
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	} finally {
	    try {
		if (bw != null) bw.close();
		if (fw != null) fw.close();
	    } catch (IOException ioe) {
		MasterLog.appendError(ioe);
	    }//try-catch
	}//try-catch-finally
    }//sortedFileInsert
}
