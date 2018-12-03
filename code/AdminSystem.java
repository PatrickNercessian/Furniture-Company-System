/*
  Programmer: Patrick Nercessian

  Class Purpose:
     The Admin Application allows the user to send Email Reminders,
     update the Order Log, and add EHF Sales Reps to the system.
 */

package system;

import java.io.*;
import org.apache.poi.xssf.usermodel.*;

import java.lang.StackTraceElement;

import java.nio.file.Files;
import java.nio.file.attribute.PosixFilePermissions;

import java.util.Date;
import java.util.ArrayList;
import java.util.Set;
import java.time.ZoneId;
import java.time.LocalDate;

import javafx.application.Application;

import javafx.stage.Stage;
import javafx.stage.Modality;

import javafx.scene.*;
import javafx.scene.text.Font;
import javafx.scene.paint.Color;
import javafx.scene.layout.*;
import javafx.scene.control.*;

import javafx.geometry.Pos;

import javafx.event.ActionEvent;
import javafx.event.EventHandler;

public class AdminSystem extends Application {

    private Stage mainStage;
    private Stage stage = null; //used in remindScene(File)

    private Scene mainMenu;

    private int x; //for lamda expression

    private Button backBtn = new Button("Back to Main Menu");

    @Override
    public void start(Stage stage) {
	mainStage = stage;

	VBox outer = new VBox(20.0);
	outer.setStyle("-fx-background-color: linear-gradient(to bottom right, Plum, MediumOrchid); "
			  + " -fx-padding: 40px;");

	Label title = new Label("EHF Admin System");
	title.setFont(new Font(30));
	title.setTextFill(Color.INDIGO);

	Button remindBtn = new Button("Send Email Reminders");
	remindBtn.setOnAction(e -> {
		mainStage.setScene(remindScene(new File("ADMIN/OTHER/maven/src/main/resources/file/Sales Rep List.txt")));
		mainStage.sizeToScene();
		mainStage.show();
	    });

	Button orderLogBtn = new Button("Update Order Log");
	orderLogBtn.setOnAction(e -> {
		mainStage.setScene(updateOrderLogScene());
		mainStage.sizeToScene();
		mainStage.show();
	    });	    

	Button addRepBtn = new Button("Add EHF Sales Rep to System");
	addRepBtn.setOnAction(e -> {
		mainStage.setScene(addRepScene());
		mainStage.sizeToScene();
		mainStage.show();
	    });

	outer.getChildren().addAll(title, remindBtn, orderLogBtn, addRepBtn);
	
	mainMenu = new Scene(outer);
	
	mainStage.setScene(mainMenu);
	mainStage.sizeToScene();
	mainStage.show();


	//not shown in Main Menu
	backBtn.setOnAction(e -> {
		mainStage.setScene(mainMenu);
		mainStage.sizeToScene();
		mainStage.show();
	    });
		
    }//start(Stage)

    @Override
    public void stop() {
	Set<Thread> threadSet = Thread.getAllStackTraces().keySet();
	Thread[] threadArray = threadSet.toArray(new Thread[threadSet.size()]);
	for (int i = 0; i < threadArray.length; i++) {
	    if (threadArray[i].getName().startsWith("Error Email Thread")) {
		MasterLog.appendEntry("Waiting for Email Thread to die");
		try {
		    threadArray[i].join();
		} catch (InterruptedException ie) {
		    MasterLog.appendError(ie);
		}
		MasterLog.append("Thread has died.");
	    }//if
	}//for
    }//stop()

    /**
     * Creates a Scene to generate open reports and remind EHF reps
     *
     * @param f text file which contains EHF reps and their emails
     * @return the scene
     */
    private Scene remindScene(File f) {
	Scene scene;
	String currentRep;
	String[] reps;
	

	VBox outer = new VBox(20);
	outer.setStyle("-fx-background-color: linear-gradient(to bottom right,  Plum, MediumOrchid); "
			  + " -fx-padding: 40px;");	

	Label title = new Label("Email Open Order Reports");
	title.setFont(new Font(30));
	title.setTextFill(Color.INDIGO);

	outer.getChildren().addAll(backBtn, title);

	//counting number of reps (to create arrays)
	int count = 0;
	BufferedReader br;
	try {
	    br = new BufferedReader(new FileReader(f));
	    while ((currentRep = br.readLine()) != null) {
		if (!currentRep.equals(""))
		    count++;
	    }//while
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	}//try-catch

	reps = new String[count];
	HBox[] salesReps = new HBox[count];
	CheckBox[] checkBoxes = new CheckBox[count];


	//Adding labels and checkboxes for each rep
	int i = 0;
	try {
	    br = new BufferedReader(new FileReader(f));
	    while ((currentRep = br.readLine()) != null) {
		if (currentRep.equals(""))
		    continue;
		reps[i] = currentRep;
		salesReps[i] = new HBox(10);
		checkBoxes[i] = new CheckBox();
		salesReps[i].getChildren().addAll(new Label(currentRep.substring(0, currentRep.indexOf(" - "))), checkBoxes[i]);
		outer.getChildren().add(salesReps[i]);
		i++;
	    }//while
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	}//try-catch

	//setting up Date Pickers
	HBox dateHBox = new HBox(20);
	DatePicker minDatePicker = new DatePicker();
	minDatePicker.setValue(LocalDate.now().minusWeeks(1));
	DatePicker maxDatePicker = new DatePicker();
	maxDatePicker.setValue(LocalDate.now());	
	dateHBox.getChildren().addAll(new Label("From:"), minDatePicker, new Label("To:"), maxDatePicker);
	
	Button emailAllBtn = new Button("Email Open Order to All EHF Sales Reps");
	emailAllBtn.setOnAction(e -> {
		MasterLog.appendEntry("Emailing Open Order to all Sales Reps");
		boolean loggedIn = Email.checkLogin("office@ehfurnishings.com", "Cattiger321");
		for (int x = 0; x < reps.length; x++) {
		    String emailAddress = reps[x].substring(reps[x].indexOf(" - ") + 3);
		    String salesRep = reps[x].substring(0, reps[x].indexOf(" - "));
		    Date minDate = Date.from(minDatePicker.getValue().atStartOfDay(ZoneId.systemDefault()).toInstant());
		    Date maxDate = Date.from(maxDatePicker.getValue().atStartOfDay(ZoneId.systemDefault()).toInstant());
		    if (loggedIn && Email.sendReminder(emailAddress, salesRep, minDate, maxDate)) {
			completeStage("All EHF Sales Reps have been emailed.");
			MasterLog.append(salesRep + " has been emailed.");			
		    }//if
		}//for
	    });

	Button emailSelectedBtn = new Button("Email Open Order to Selected EHF Sales Reps");
	emailSelectedBtn.setOnAction(e -> {
		MasterLog.appendEntry("Emailing Open Order to selected Sales Reps");	
		boolean loggedIn = Email.checkLogin("office@ehfurnishings.com", "Cattiger321");	
		for (x = 0; x < checkBoxes.length; x++) {
		    if (checkBoxes[x].isSelected()) {
			String emailAddress = reps[x].substring(reps[x].indexOf(" - ") + 3);
			String salesRep = reps[x].substring(0, reps[x].indexOf(" - "));
			Date minDate = Date.from(minDatePicker.getValue().atStartOfDay(ZoneId.systemDefault()).toInstant());
			Date maxDate = Date.from(maxDatePicker.getValue().atStartOfDay(ZoneId.systemDefault()).toInstant());
			if (loggedIn && Email.sendReminder(emailAddress, salesRep, minDate, maxDate)) {
			    completeStage(salesRep + " has been emailed.");
			    MasterLog.append(salesRep + " has been emailed.");
			}//if
		    }//if
		}//for
	    });

	

	outer.getChildren().add(1, emailAllBtn);
	outer.getChildren().addAll(dateHBox, emailSelectedBtn);
	

	scene = new Scene(outer);
	return scene;
    }//remindScene()

    /**
     * Creates a Scene to add an EHF Rep
     *
     * @return the Scene
     */
    private Scene addRepScene() {
	Scene scene;

	VBox outer = new VBox(20);
	outer.setStyle("-fx-background-color: linear-gradient(to bottom right,  Plum, MediumOrchid); "
			  + " -fx-padding: 40px;");	


	Label title = new Label("Add New EHF Sales Rep");
	title.setFont(new Font(30));
	title.setTextFill(Color.INDIGO);
	
	
	HBox repBox = new HBox(40);
	TextField nameField = new TextField();
	nameField.setPromptText("New Sales Rep's Name");

	TextField emailField = new TextField();
	emailField.setPromptText("New Sales Rep's Email");

	repBox.getChildren().addAll(nameField, emailField);

	
	Button addBtn = new Button("Add EHF Rep");
	addBtn.setOnAction(e -> {
		File textFile = new File("ADMIN/OTHER/maven/src/main/resources/file/Sales Rep List.txt");
		BinaryInsert.sortedFileInsert(textFile,nameField.getText() + " - " + emailField.getText() + "\n");
		completeStage("Added " + nameField.getText() + " at " + emailField.getText() + " to the list of EHF Sales Reps. If this was a mistake, please see Chris Nercessian to remove the Sales Rep from the text document");
		MasterLog.appendEntry("Added " + nameField.getText() + " at " + emailField.getText() + " to the list of EHF Sales Reps.");
	    });
	

	outer.getChildren().addAll(backBtn, title, repBox, addBtn);

	scene = new Scene(outer);
	return scene;
    }

    /**
     * Creates a Scene to update the Order Log
     *
     * @return the scene.
     */
    public Scene updateOrderLogScene() {
	VBox outer = new VBox(20);
	outer.setStyle("-fx-background-color: linear-gradient(to bottom right,  Plum, MediumOrchid); "
			  + " -fx-padding: 40px;");
	outer.setAlignment(Pos.CENTER);
	
	Label title = new Label("Update Order Log");
	title.setFont(new Font(30));
	title.setTextFill(Color.INDIGO);

	Button updateBtn = new Button("Update Order Log");
	updateBtn.setOnAction(e -> {
		try {
		    /*
		    String[] list = OrderLog.listOfSC();
		    for (int i = 0; i < list.length; i++) {
			MasterLog.appendEntry(list[i]);
			OrderLog.populateOrder(list[i]);
		    }
		    */
		    OrderLog.populateAllOrders();
		    completeStage("Completed Update");
		} catch (Exception ex) {
		    MasterLog.appendError(ex);
		    completeStage("FAILED: SEE MASTER LOG");
		}//try-catch
	    });
	
	outer.getChildren().addAll(backBtn, title, updateBtn);
	return new Scene(outer);
    }

    /**
     * Creates a stage which activates the EventHandler when yesBtn is pressed
     *
     * @param action The Action to complete when Yes is pressed
     * @param text The message to show.
     * @return the Stage
     */
    private Stage confirmation(EventHandler<ActionEvent> action, String text) {
	Stage stage = new Stage();
	VBox outer = new VBox(20);

	outer.setStyle("-fx-background-color: linear-gradient(to bottom right,  Plum, MediumOrchid); "
			  + "-fx-spacing: 20px; -fx-padding: 40px;");	
	
	Label label = new Label(text);

	HBox yesno = new HBox(30);
	Button yes = new Button("Yes");
	yes.setOnAction(action);
	
	Button no = new Button("No");
	no.setOnAction(e -> stage.close());
	yesno.getChildren().addAll(yes, no);
	
	outer.getChildren().addAll(label, yesno);
	
	Scene scene = new Scene(outer);
	stage.setScene(scene);
	stage.initModality(Modality.APPLICATION_MODAL);
	stage.sizeToScene();
	
	return stage;
    }//confirmation()

    /**
     * Shows a stage to confirm completion of a task
     *
     * @param text The message to show.
     */    
    private void completeStage(String text) {
	Stage stage = new Stage();
	VBox vbox = new VBox(30);
	vbox.setStyle("-fx-background-color: linear-gradient(to bottom right,  Plum, MediumOrchid);"
		      + "-fx-padding: 40px;");
	Scene s = new Scene(vbox);
				
	Button ok = new Button("OK");
	ok.setOnAction(ev -> {
		stage.close();
		mainStage.setScene(mainMenu);
		mainStage.sizeToScene();
		mainStage.show();
	    });
	vbox.getChildren().addAll(new Label(text), ok);
	stage.setScene(s);
	stage.sizeToScene();
	stage.show();
    }//completeStage()    

    public static void main(String[] args) {
	Application.launch(args);
    }
    
}//AdminSystem
