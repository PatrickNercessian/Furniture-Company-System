/*
  Programmer: Patrick Nercessian

  Class Purpose:
     This class contains code pertaining to the Email System part of the program
     including helper methods to send email and confirm login info.
 */

package system;

import java.io.*;

import java.util.Date;
import java.util.Properties;
import java.util.ArrayList;    

import javax.activation.*;
import javax.mail.*;
import javax.mail.internet.*;

public class Email {

    private static Properties properties;
    private static Session session;
    

    public static boolean checkLogin(String username, String password) {
	Transport transport;
	properties = System.getProperties();
	session = Session.getInstance(properties,
					      new Authenticator() {
						  protected PasswordAuthentication getPasswordAuthentication() {
						      return new PasswordAuthentication(username, password);
						  }
					      });
	
	properties.setProperty("mail.smtp.host", "smtpout.secureserver.net");
	properties.setProperty("mail.smtp.user", username);
	properties.setProperty("mail.smtp.password", password);
	properties.setProperty("mail.smtp.auth", "true");
	try {
	    transport = session.getTransport("smtp");
	    transport.connect("smtpout.secureserver.net", username, password);
	    transport.close();
		
	    return true;
	} catch (MessagingException e) {
	    return false;
	}
    }

    /**
     * Sends an email
     *
     * @param recipient the email address to receive the email
     * @param subject email subject
     * @param body email body
     * @param attachment email attachment
     * @return boolean value if email sent or not
     */
    public static boolean sendEmail(String username, String recipient, String subject, String body, File ... attachment) {
	try {
	    MimeMessage msg = new MimeMessage(session);
	    msg.setFrom(new InternetAddress(username));
	    msg.addRecipient(Message.RecipientType.TO, new InternetAddress(recipient));
	    msg.setSubject(subject);

	    BodyPart messageBodyPart = new MimeBodyPart();
	    messageBodyPart.setText(body);
	    Multipart multipart = new MimeMultipart();
	    multipart.addBodyPart(messageBodyPart);

	    for (int i = 0; i < attachment.length; i++) {
		messageBodyPart = new MimeBodyPart();
		DataSource source = new FileDataSource(attachment[i]);
		messageBodyPart.setDataHandler(new DataHandler(source));
		messageBodyPart.setFileName(attachment[i].getName());
		multipart.addBodyPart(messageBodyPart);
	    }//for

	    msg.setContent(multipart);

	    Transport.send(msg);
	    return true;
	} catch (MessagingException mex) {
	    MasterLog.appendError(mex);
	}//try-catch
	return false;
    }//sendEmail(String, String, String, String, File...)

    public static boolean sendReminder(String recipient, String salesRep, Date minDate, Date maxDate) {
	if (!checkLogin("office@ehfurnishings.com", "Cattiger321"))
	    return false;
	
	File textFile = new File("ADMIN/OTHER/maven/src/main/resources/file/Log - Sales Contract.txt");	
	BufferedReader br;
	try {
	    br = new BufferedReader(new FileReader(textFile));
	    String currentOrder;
	    while ((currentOrder = br.readLine()) != null) {
		if (currentOrder.equals("") || currentOrder.contains("reserved")) //skip if line is empty
		    continue;

		
		//changing order status to SHIPPED if container number has been entered
		if (currentOrder.indexOf("SC#:") < 0 || currentOrder.indexOf(" - Client:") < 0) MasterLog.appendEntry(currentOrder);
		String sc = currentOrder.substring(currentOrder.indexOf("SC#:") + 5, currentOrder.indexOf(" - Client:"));
		if (Log.getOrderStatus(sc).equals("BOOKING")) {
		    ExistingDocuments ed = new ExistingDocuments(Log.findFile(sc));
		    try {
			if (ed.hasContainerNum()) {
			    Log.changeOrderStatus("SHIPPED", sc);
			    MasterLog.appendEntry("Changed Sales Contract #" + sc + " Status to SHIPPED");
			}//if
		    } catch (Exception ex) {
			MasterLog.appendEntry("ERROR ON SC#" + sc);
			MasterLog.appendError(ex);
		    }
		}//if
	    }//while
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	}//try-catch

	String[] pendingList = Log.compileList(salesRep, "PENDING", minDate, maxDate);
	String[] confirmedList = Log.compileList(salesRep, "CONFIRMED", minDate, maxDate);
	String[] bookingList = Log.compileList(salesRep, "BOOKING", minDate, maxDate);
	String[] shippedList = Log.compileList(salesRep, "SHIPPED", minDate, maxDate);
	String[] canceledList = Log.compileList(salesRep, "CANCELED", minDate, maxDate);
	String[] reinstatedList = Log.compileList(salesRep, "REINSTATED", minDate, maxDate);

	String[][] list = {pendingList, confirmedList, bookingList, shippedList, canceledList, reinstatedList};
	OpenReport or = new OpenReport(list, salesRep, "Open Report");
	or.createPopulate();
	
	String minDateStr = minDate.toString().substring(0, minDate.toString().indexOf(" 00:"));
	String maxDateStr = maxDate.toString().substring(0, maxDate.toString().indexOf(" 00:"));	
	String message = "Dear " + salesRep + ",\n\nBelow is your order summary from " + minDateStr + " to " + maxDateStr;
	message += Log.orderListString(pendingList, "PENDING") + "$" + or.getTotals()[0];
	message += Log.orderListString(confirmedList, "CONFIRMED") + "$" + or.getTotals()[1];
	message += Log.orderListString(bookingList, "BOOKING") + "$" + or.getTotals()[2];	
	message += Log.orderListString(shippedList, "SHIPPED") + "$" + or.getTotals()[3];
	message += Log.orderListString(canceledList, "CANCELED") + "$" + or.getTotals()[4];
	message += Log.orderListString(reinstatedList, "REINSTATED") + "$" + or.getTotals()[5];

	boolean completed = sendEmail("office@ehfurnishings.com", recipient, "Your EHF Order Summary", message, or.getReportFile());
	//	boolean completed = true;
	message = "Dear " + salesRep + ",\n\nPlease see attached Excel document for a list of orders that have not been updated with Model Information\n";

	String[] modelLackingList = Log.modelLackingOrders(salesRep);
	int numLacking = 0;
	for (int i = 0; i < modelLackingList.length; i++)
	    if (modelLackingList[i] != null) numLacking++;
	String[][] formattedList = new String[modelLackingList.length][5];
	for (int i = 0; i < modelLackingList.length; i++) {
	    String order = modelLackingList[i];
	    if (order != null) {
		formattedList[i][0] = order.substring(order.indexOf("SC#:") + 5, order.indexOf(" - Client:"));
		formattedList[i][1] = order.substring(order.indexOf("Client:") + 8, order.indexOf(" - Factory:"));
		formattedList[i][2] = order.substring(order.indexOf("Factory:") + 9, order.indexOf(" - Model:"));
		formattedList[i][3] = order.substring(order.indexOf("Model:") + 7, order.indexOf("      "));
		formattedList[i][4] = order.substring(order.indexOf("       ") + 12);
	    }//if
	}
	OpenReport reminder = new OpenReport(formattedList, salesRep, "Reminder");
	reminder.modelLackingExcel();

	if (!sendEmail("office@ehfurnishings.com", recipient, "You have " + numLacking + " orders without updated models", message, reminder.getReportFile()))
	    completed = false;

	return completed;
    }//sendReminder(String, String, Date, Date)
    
    
    public static boolean sendLeads(String country, String user, String pass, String subject, String body, File ... attachment) {
	if (!checkLogin(user, pass))
	    return false;
	boolean complete = true;
	BufferedReader br;
	ArrayList<String> listOfRecipients = new ArrayList<>();
	try {
	    br = new BufferedReader(new FileReader("ADMIN/OTHER/maven/src/main/resources/file/Lead List.txt"));
	    String current;
	    while ((current = br.readLine()) != null) {
		if (current.equals("")) //skip if line is empty
		    continue;
		if (country.equals("Select Country") || current.substring(current.indexOf(" - Country:") + 12).equalsIgnoreCase(country)) {
		    String info = current.substring(0, current.indexOf(" - Country:"));
		    String greeting = "Dear " + current.substring(current.indexOf("Contact Name: ") + 14, current.indexOf(" - Country:")) + ",\n\n";
		    try {
			if (sendEmail(user, current.substring(0, current.indexOf(" - Company Name:")), subject,
				       greeting + body, attachment)) {
			    listOfRecipients.add(info);			    
			} else {
			    complete = false;			    
			    listOfRecipients.add("FAILED: " + info);
			}//if-else
		    } catch (Exception ex) {
			complete = false;
			listOfRecipients.add("FAILED: " +  info);
		    }//try-catch
		}//if
	    }//while
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	}//try-catch
	int countFailed = 0;
	for (int i = 0; i < listOfRecipients.size(); i++)
	    if (listOfRecipients.get(i).startsWith("FAILED:")) countFailed++;
	if (country.equals("Select Country")) country = "ALL COUNTRIES";
	String confirmationMsg = "You sent a Mass Email to " + listOfRecipients.size() + " recipients from " + country + "."
	                         + countFailed + " emails failed to send. Attached files have been attached to this email."
	                         + "Below is the Mass Email and the list of recipients.\n\n\n";

	confirmationMsg += body + "\n\n\n\nRecipients:\n";

	for (int i = 0; i < listOfRecipients.size(); i++)
	    confirmationMsg += listOfRecipients.get(i) + "\n";
	
	if (!sendEmail("office@ehfurnishings.com", user, "Confirmation of Mass Email", confirmationMsg, attachment))
	    complete = false;
	
	return complete;
    }//sendLeads(String, String)
}
