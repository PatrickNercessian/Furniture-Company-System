/*
  Programmer: Patrick Nercessian

  Class Purpose:
     The Main JavaFX Application for the program. Nearly all parts of the program
     can be utilized through this App. The few that cannot are used through the
     AdminSystem class.
 */

package system;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;

import org.apache.poi.xssf.usermodel.*;

import java.lang.StackTraceElement;

import java.awt.Desktop;
import java.awt.GraphicsDevice;
import java.awt.GraphicsEnvironment;

import javafx.application.Application;
import javafx.application.Platform;

import javafx.stage.Stage;
import javafx.stage.Modality;
import javafx.stage.FileChooser;

import javafx.scene.Scene;
import javafx.scene.Group;
import javafx.scene.layout.*;
import javafx.scene.control.*;
import javafx.scene.text.*;
import javafx.scene.paint.Color;

import javafx.event.ActionEvent;
import javafx.event.EventHandler;

import java.util.Date;
import java.util.ArrayList;
import java.util.Set;
import java.text.SimpleDateFormat;

import java.time.ZoneId;
import java.time.LocalDate;

import javafx.geometry.Pos;

import javafx.concurrent.Task;

public class MenuApp extends Application {

    private Stage mainStage;    
    private Stage stage = null; //used for confirmation() scene
    private String[][] finalResults = null; //used in searchDistributionScene() lamda expression
    private File selectedFile = null; //used in emailLeadsScene()
    private ArrayList<File> attachedFiles = new ArrayList<File>();

    private int screenHeight;

    private Scene mainMenu;

    private Button yes = new Button("Yes"); //used in confirmation() so that it can be disarmed in other methods
    private Button backBtn = new Button("Back to Main Menu");

    @Override
    public void stop() {
	MasterLog.appendEntry("Main System Exited");	
	Set<Thread> threadSet = Thread.getAllStackTraces().keySet();
	Thread[] threadArray = threadSet.toArray(new Thread[threadSet.size()]);
	for (int i = 0; i < threadArray.length; i++) {
	    if (threadArray[i].getName().startsWith("Error Email Thread")) {
		MasterLog.append("Waiting for Email Thread to die...");
		try {
		    threadArray[i].join();
		} catch (InterruptedException ie) {
		    MasterLog.appendError(ie);
		}
		MasterLog.append("Thread has died.");
	    }//if
	}//for
    }//stop()

    @Override
    public void start(Stage stage) {
	GraphicsDevice gd = GraphicsEnvironment.getLocalGraphicsEnvironment().getDefaultScreenDevice();
	screenHeight = gd.getDisplayMode().getHeight();	
	
	mainStage = stage;

	VBox mainVBox = new VBox(20.0);
	mainVBox.setStyle("-fx-background-color: linear-gradient(to bottom right, NavajoWhite, BurlyWood); "
			  + "-fx-spacing: 20px; -fx-padding: 40px;");
	mainVBox.setAlignment(Pos.CENTER);


	Label title = new Label("EHF System");
	title.setFont(new Font(50));
	title.setTextFill(Color.INDIGO);

	Button createSC = new Button("Create an Order");
	createSC.setOnAction(e -> {
		mainStage.setScene(salesContractScene(false));
		mainStage.sizeToScene();
		mainStage.show();
	    });

	Button createCustomSC = new Button("Reissue an Order");
	createCustomSC.setOnAction(e -> {
		mainStage.setScene(salesContractScene(true));
		mainStage.sizeToScene();
		mainStage.show();
	    });

	Button reserveID = new Button("Reserve Sales Contract IDs");
	reserveID.setOnAction(e -> {
		mainStage.setScene(reserveScene());
		mainStage.sizeToScene();
		mainStage.show();
	    });

	Button statusChange = new Button("Update Order Status");
	statusChange.setOnAction(e -> {
		mainStage.setScene(statusScene());
		mainStage.sizeToScene();
		mainStage.show();
	    });

	Button assignPO = new Button("Assign EHF#'s");
	assignPO.setOnAction(e -> {
		mainStage.setScene(assignScene());
		mainStage.sizeToScene();
		mainStage.show();
	    });

	Button openSC = new Button("Open an Order");
	openSC.setOnAction(e -> {
		mainStage.setScene(openSCScene());
		mainStage.sizeToScene();
		mainStage.show();
	    });		

	Button addClient = new Button("Add New Client");
	addClient.setOnAction(e -> {
		mainStage.setScene(addClientScene(true));
		mainStage.sizeToScene();
		mainStage.show();
	    });

	Button updateClient = new Button("Update Existing Client");
	updateClient.setOnAction(e -> {
		mainStage.setScene(addClientScene(false));
		mainStage.sizeToScene();
		mainStage.show();
	    });

	Button addFactory = new Button("Add New Factory");
	addFactory.setOnAction(e -> {
		mainStage.setScene(addFactoryScene(true));
		mainStage.sizeToScene();
		mainStage.show();
	    });

	Button updateFactory = new Button("Update Existing Factory");
	updateFactory.setOnAction(e -> {
		mainStage.setScene(addFactoryScene(false));
		mainStage.sizeToScene();
		mainStage.show();
	    });	

	Button openLog = new Button("Open Sales Contract Log");
	openLog.setOnAction(e -> {
		try {
		    File old = new File("ADMIN/OTHER/maven/src/main/resources/file/BackupLog.xlsx");
		    File copy = new File("ADMIN/OTHER/maven/src/main/resources/file/BackupLog Copy.xlsx");
		    Files.copy(old.toPath(), copy.toPath(), StandardCopyOption.REPLACE_EXISTING);
		    Desktop.getDesktop().open(new File("ADMIN/OTHER/maven/src/main/resources/file/BackupLog Copy.xlsx"));
		} catch (Exception ex) {
		    MasterLog.appendError(ex);
		}//try-catch
	    });

	Button updateModels = new Button("Update an Order's Models");
	updateModels.setOnAction(e -> {
		mainStage.setScene(updateModelsScene());
		mainStage.sizeToScene();
		mainStage.show();
	    });	

	Button searchDistribution = new Button("Search Distribution Log");
	searchDistribution.setOnAction(e -> {
		mainStage.setScene(searchDistributionScene());
		mainStage.sizeToScene();
		mainStage.show();
	    });

	Button listOrders = new Button("List Orders");
	listOrders.setOnAction(e -> {
		mainStage.setScene(listOrdersScene());
		mainStage.sizeToScene();
		mainStage.show();
	    });

	Button addLead = new Button("Add Lead");
	addLead.setOnAction(e -> {
		mainStage.setScene(addLeadScene());
		mainStage.sizeToScene();
		mainStage.show();
	    });

	Button emailLeads = new Button("Email Leads");
	emailLeads.setOnAction(e -> {
		mainStage.setScene(emailLeadsScene());
		mainStage.sizeToScene();
		mainStage.show();
	    });	


	VBox orderVbox = new VBox(20, createSC, statusChange, assignPO, openSC);
	orderVbox.setAlignment(Pos.TOP_CENTER);	
	VBox outerOrder = new VBox(40, orderVbox, createCustomSC);
	outerOrder.setStyle("-fx-padding: 20px;");
	outerOrder.setAlignment(Pos.TOP_CENTER);	
	
	VBox databaseVbox = new VBox(20, addClient, updateClient, addFactory, updateFactory);
	databaseVbox.setAlignment(Pos.TOP_CENTER);
	VBox outerDatabase = new VBox(40, databaseVbox, openLog);
	outerDatabase.setStyle("-fx-padding: 20px;");
	outerDatabase.setAlignment(Pos.TOP_CENTER);	

	VBox reportVbox = new VBox(20, updateModels, searchDistribution, listOrders);
	reportVbox.setStyle("-fx-padding: 20px;");
	reportVbox.setAlignment(Pos.TOP_CENTER);

	VBox emailListVbox = new VBox(20, addLead, emailLeads);
	emailListVbox.setStyle("-fx-padding: 20px;");
	emailListVbox.setAlignment(Pos.TOP_CENTER);		

	Tab orderTab = new Tab("Orders", outerOrder);
	orderTab.setClosable(false);
	Tab databaseTab = new Tab("Databases", outerDatabase);
	databaseTab.setClosable(false);
	Tab reportTab = new Tab("Reports", reportVbox);
	reportTab.setClosable(false);
	Tab emailListTab = new Tab("Email List", emailListVbox);
	emailListTab.setClosable(false);
	TabPane tabPane = new TabPane(orderTab, databaseTab, reportTab, emailListTab);
	tabPane.setTabMinWidth(56);
	//	tabPane.setMaxWidth(325);
	
	mainVBox.getChildren().addAll(title, tabPane);

	mainMenu = new Scene(mainVBox);
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
    
    /**
     * Creates the scene for Generating new Sales Contracts
     * @return the scene
     */
    private Scene salesContractScene(boolean customSC) {
	Scene scene;
	
	VBox outerBox = new VBox();
	outerBox.setStyle("-fx-background-color: linear-gradient(to bottom right, NavajoWhite, BurlyWood); "
			  + "-fx-spacing: 30px; -fx-padding: 40px;");
	outerBox.setAlignment(Pos.CENTER);


	//Title
	Label title = new Label("Create an Order");
	title.setFont(new Font(40));
	title.setTextFill(Color.INDIGO);

	
	VBox scVBox = new VBox(20);
	scVBox.setAlignment(Pos.CENTER);	

	TextField scTextField = new TextField();
	scTextField.setPromptText("Enter custom SC#");

	ComboBox<String> customerCB = new ComboBox<>();
	new AutoCompleteComboBox(customerCB);
	customerCB.getItems().add("Select Customer");
	customerCB.getSelectionModel().selectFirst();
	try { //adding all client names to choicebox
	    BufferedReader br = new BufferedReader(new FileReader("ADMIN/OTHER/maven/src/main/resources/file/Client List.txt"));
	    String client;
	    while ((client = br.readLine()) != null)
		customerCB.getItems().add(client);
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	}//try-catch

	ComboBox<String> factoryCB = new ComboBox<>();
	new AutoCompleteComboBox(factoryCB);
	factoryCB.getItems().add("Select Factory");
	factoryCB.getSelectionModel().selectFirst();
	try { //adding all factory names to choicebox
	    BufferedReader br = new BufferedReader(new FileReader("ADMIN/OTHER/maven/src/main/resources/file/Factory List.txt"));
	    String factory;
	    while ((factory = br.readLine()) != null)
		factoryCB.getItems().add(factory);
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	}//try-catch

	
	TextField modelsText = new TextField();
	modelsText.setPromptText("Models");
	
	
	HBox incotermBox = new HBox(10);
	ComboBox<String> incotermCB = new ComboBox<>();
	new AutoCompleteComboBox(incotermCB);
	incotermCB.getItems().addAll("Select Incoterm", "EXW", "FOB","C&F", "CIF");
	incotermCB.getSelectionModel().selectFirst();
	
	TextField whereText = new TextField();
	whereText.setPromptText("Where?");
	
	incotermBox.getChildren().addAll(incotermCB, whereText);


	
	HBox containersBox = new HBox(10);
	ComboBox<String> containersCB = new ComboBox<>();
	new AutoCompleteComboBox(containersCB);
	containersCB.getItems().addAll("Select Container Size", "LTL", "20'", "40'STD", "40'HC", "53'TRL");
	containersCB.getSelectionModel().selectFirst();
	
	TextField numContainers = new TextField();
	numContainers.setPromptText("Number of Containers");
	
	containersBox.getChildren().addAll(containersCB, numContainers);

	ToggleGroup tGroup = new ToggleGroup();
	ToggleButton shortModel = new ToggleButton("25 Models");
	ToggleButton longModel = new ToggleButton("125 Models");
	shortModel.setToggleGroup(tGroup);
	longModel.setToggleGroup(tGroup);
	HBox modelLengthHBox = new HBox(shortModel, longModel);
	modelLengthHBox.setAlignment(Pos.CENTER);

	Button generateButton = new Button("Generate Order");
	generateButton.setDefaultButton(true);
	generateButton.setOnAction(e -> {
		
		stage = confirmation(ev -> {
			yes.disarm();
			stage.setHeight(225);

			boolean invalid = (customerCB.getValue().equals("Select Customer")
					   || factoryCB.getValue().equals("Select Factory")
					   || modelsText.getText().equals("") || modelsText.getText().contains("/")
					   || modelsText.getText().contains(" ")
					   || incotermCB.getValue().equals("Select Incoterm")
					   || whereText.getText().equals("")
					   || containersCB.getValue().equals("Select Container Size")
					   || !validInput(customerCB) || !validInput(factoryCB) || !validInput(incotermCB)
					   || !validInput(containersCB));
			
			int scInt = -1;
			if (customSC) {
			    try {
				scInt = Integer.parseInt(scTextField.getText());
				if (!Log.isUniqueSC(scTextField.getText())) //can remove this after reissue button is removed
				    invalid = true;
			    } catch (NumberFormatException nfe) {
				invalid = true;
			    }
			}//if

			if (invalid) {

			    stage.close();
			    scVBox.getChildren().add(new Label("Make sure to input all fields."));
			    scVBox.getChildren().add(new Label("Make sure no \"/\"s or spaces are in Model."));
			    if (customSC)
				scVBox.getChildren().add(new Label("Make sure this SC# has not been issued before."));
			} else {
			    MasterLog.appendEntry("Creating New Order...");				
			    int num = -1; //numContainers
			    try {
				num = Integer.parseInt(numContainers.getText());
			    } catch (Exception ex) {
				MasterLog.appendError(ex);
				scVBox.getChildren().add(new Label("Only enter a number for \"Number of Containers\""));
			    }//try-catch

			    String models = modelsText.getText().replaceAll("\\.", "").trim();
			    CreateDocuments cd = new CreateDocuments(customerCB.getValue(), factoryCB.getValue(),
								     models, incotermCB.getValue(),
								     whereText.getText().trim(), containersCB.getValue(),
								     num, scInt, longModel.isSelected());
			

			    Thread t = new Thread(() -> {
				    try {
					cd.populatePO();
					cd.populateWithClient();
					cd.autoPopulate();
					cd.fillDate();
					cd.fillTotals();
					cd.populateCalcSheet();
					cd.populateEmail();
					cd.writeFile();
					cd.openSheet();
					
					Platform.exit();
					MasterLog.append("Order Created");
					
					Platform.runLater(() -> {
						stage.close();
						yes.arm();
					    });
				    } catch (Exception ex) {
					MasterLog.appendError(ex);
					yes.arm();
				    }//try-catch
				});
			    t.setDaemon(true);
			    t.start();
			}//if-else
		    }
		    , "Are you sure you want to generate this order?");
		stage.show();
	    });

	
	if (customSC)
	    scVBox.getChildren().add(scTextField);

	scVBox.getChildren().addAll(customerCB, factoryCB, modelsText, incotermBox, containersBox, modelLengthHBox);
	outerBox.getChildren().addAll(backBtn, title, scVBox, generateButton);

	scene = new Scene(outerBox);
	return scene;
    }//salesContractScene()

    public static boolean validInput(ComboBox cb) {
	for (int i = 0; i < cb.getItems().size(); i++) { //ensure entry is one of the items
	    if (cb.getValue().equals(cb.getItems().get(i)))
		return true;
	}//for
	return false;
    }//validInput(ComboBox)


    private Scene reserveScene() {
	VBox vbox = new VBox(20);
	vbox.setStyle("-fx-background-color: linear-gradient(to bottom right, NavajoWhite, BurlyWood); -fx-spacing: 30px; -fx-padding: 40px;");
	Scene scene = new Scene(vbox);

	Label title = new Label("Sales Contract ID Reservation");
	title.setFont(new Font(30));
	title.setTextFill(Color.INDIGO);
	
	TextField reserveText = new TextField();
	reserveText.setPromptText("Number of Sales Contract IDs to Reserve");

	Button reserveButton = new Button("Reserve");
	reserveButton.setOnAction(e -> {
		MasterLog.appendEntry("Reserving Sales Contracts...");
		try {
		    int numToReserve = Integer.parseInt(reserveText.getText());
		    int id = CreateDocuments.getCurrentId();
		    SimpleDateFormat sdf = new SimpleDateFormat("MM/dd/yyyy HH:mm:ss");
		    
		    CreateDocuments.reserveIDs(numToReserve);
		    String logEntry = sdf.format(new Date()) + " - " + id + " through "
			+ (id + numToReserve) + " reserved by " + System.getProperty("user.name");
		    Log.updateLog(new File("ADMIN/OTHER/maven/src/main/resources/file/Log - Sales Contract.txt"), logEntry + "\n\n");
		    Log.updateExcelLog(logEntry);
		    
		    completeStage("Reserved Sales Contracts " + id + " through " + (id + numToReserve));
		    MasterLog.append("Reserved Sales Contracts " + id + " through " + (id + numToReserve));
		} catch (Exception ex) {
		    MasterLog.appendError(ex);
		    Label error = new Label("Make sure you enter a number!");
		    vbox.getChildren().add(error);
		}//try-catch
		mainStage.sizeToScene();
	    });

	vbox.getChildren().addAll(backBtn, title, reserveText, reserveButton);

	return scene;
    }//reserveScene()

    private Scene statusScene() {
	VBox outer = new VBox(30);
	outer.setStyle("-fx-background-color: linear-gradient(to bottom right, NavajoWhite, BurlyWood); "
			  + "-fx-spacing: 20px; -fx-padding: 40px;");
	outer.setAlignment(Pos.CENTER);	
	HBox allButTitle = new HBox(30);	

	VBox scvbox = new VBox(5);
	TextField sc = new TextField();
	sc.setPromptText("Sales Contract Number");
	scvbox.getChildren().addAll(new Label("Enter SC#:"), sc);

	
	VBox choices = new VBox(30);
	
	Button confirmBtn = new Button("Confirm");
	confirmBtn.setOnAction(e -> {
		String scStr = sc.getText();
		String status = Log.getOrderStatus(scStr);
		if (status.equals("PENDING") || status.equals("REINSTATED") || status.equals("CONFIRMED")) {
		    stage = confirmation(ev -> {
			    yes.disarm();
			    MasterLog.appendEntry("Confirming Sales Contract #" + scStr + "...");
			    stage.setHeight(250);

			    Thread t = new Thread(() -> {
				    try {
					File f = Log.findFile(scStr);
					ExistingDocuments ed = new ExistingDocuments(f);
					
					ed.confirmedCopyPI();
				    } catch (FileNotFoundException fnfe) {
					Platform.runLater(() -> completeStage("Sales Contract #" + scStr + " file could not be found"));
				    }//try-catch
				    
				    Platform.runLater(() -> {
					    stage.close();
					    yes.arm();
					    if (CheckCompletion.checkConfirm(scStr)) {
						Log.changeOrderStatus("CONFIRMED", scStr);
						completeStage("Sales Contract #" + scStr + " has been CONFIRMED");
						MasterLog.append("Sales Contract #" + scStr + " has been CONFIRMED");
					    } else {
						completeStage("ERROR: CONFIRMATION FAILED");
						MasterLog.append("ERROR: CONFIRMATION FAILED");
					    }//if-else
					});
				});
			    t.setDaemon(true);
			    t.start();
			}
			,"Are you sure you want to CONFIRM Sales Contract #"+ scStr);
				  
		    stage.show();

		} else if (status.equals("")) {
		    completeStage("Sales Contract # " + scStr + " CANNOT BE FOUND. Double check that this is the correct SC#");
		} else {
		    completeStage("You cannot CONFIRM a " + status + " order!");
		}//if-elseif-else
	    });

	Button bookingBtn = new Button("Booking");
	bookingBtn.setOnAction(e -> {
		String scStr = sc.getText();
		String status = Log.getOrderStatus(scStr);
		if (status.equals("CONFIRMED")) {
		    stage = confirmation(ev -> {
			    yes.disarm();
			    MasterLog.appendEntry("Booking Sales Contract #" + scStr + "...");
			    stage.setHeight(250);

			    Thread t = new Thread(() -> {
				    try {
					File f = Log.findFile(scStr);
					ExistingDocuments ed = new ExistingDocuments(f);
					if (ed.containsPO()) {
					    ed.bookingCopyPI();
					    ed.populateCO(scStr);
					} else {
					    Platform.runLater(() -> {
						    stage.close();
						    completeStage("ERROR: Sales Contract #" + scStr + "has not been assigned an EHF#");
						});
					}//if-else
				    } catch (FileNotFoundException fnfe) {
					Platform.runLater(() -> completeStage("Sales Contract #" + scStr + " file could not be found"));
				    } catch (NullPointerException npe){
				    }//try-catch
					
				    Platform.runLater(() -> {
					    stage.close();
					    yes.arm();
					    if (CheckCompletion.checkBooking(scStr)) {
						Log.changeOrderStatus("BOOKING", scStr);
						completeStage("Sales Contract #" + scStr + " is now in BOOKING");
						MasterLog.append("Sales Contract #" + scStr + " is now in BOOKING");
					    } else {
						completeStage("ERROR: BOOKING FAILED");
						MasterLog.append("ERROR: BOOKING FAILED");
					    }//if-else
					});
				});
			    t.setDaemon(true);
			    t.start();				
			}
			, "Are you sure you want to place Sales Contract #"+ sc.getText() + " for BOOKING?");

		    stage.show();
		} else if (status.equals("")) {
		    completeStage("Sales Contract # " + scStr + " CANNOT BE FOUND. Double check that this is the correct SC#");
		} else {
		    completeStage("You cannot place a " + status + " order for BOOKING!");
		}//if-elseif-else
	    });
						 
	Button cancelBtn = new Button("Cancel");
	cancelBtn.setOnAction(e -> {
		String scStr = sc.getText();
		String status = Log.getOrderStatus(scStr);
		if (status.equals("PENDING") || status.equals("REINSTATED") || status.equals("CONFIRMED")) {
		    stage = confirmation(ev -> {
			    yes.disarm();
			    MasterLog.appendEntry("Canceling Sales Contract #" + scStr + "...");
			    stage.setHeight(250);
			    stage.show();
			    Thread t = new Thread(() -> {
				    Log.changeOrderStatus("CANCELED", scStr);
				    Platform.runLater(() -> {
					    stage.close();
					    yes.arm();
					    completeStage("Sales Contract #" + scStr + " has been CANCELED");
					    MasterLog.append("Sales Contract #" + scStr + " has been CANCELED");
					});
				});
			    t.setDaemon(true);
			    t.start();
			}
			, "Are you sure you want to CANCEL Sales Contract #"+ scStr);
		    stage.show();

		} else if (status.equals("")) {
		    completeStage("Sales Contract # " + scStr + " CANNOT BE FOUND. Double check that this is the correct SC#");
		} else {
		    completeStage("You cannot CANCEL a " + status + " order!");
		}//if-elseif-else
	    });

	Button reinstateBtn = new Button("Reinstate");
	reinstateBtn.setOnAction(e -> {
		String scStr = sc.getText();
		String status = Log.getOrderStatus(scStr);
		if (status.equals("CANCELED")) {
		    stage = confirmation(ev -> {
			    yes.disarm();
			    MasterLog.appendEntry("Reinstating Sales Contract #" + scStr + "...");
			    stage.setHeight(250);
			    stage.show();			    
			    Thread t = new Thread(() -> {			    
				    Log.changeOrderStatus("REINSTATED", sc.getText());
				    Platform.runLater(() -> {
					    stage.close();
					    yes.arm();
					    completeStage("Sales Contract #" + sc.getText() + " has been REINSTATED");
					    MasterLog.append("Sales Contract #" + sc.getText() + " has been REINSTATED");
					});
				});
			    t.setDaemon(true);
			    t.start();
			}
			, "Are you sure you want to REINSTATE Sales Contract #"+ sc.getText());
		    stage.show();
		    
		} else if (status.equals("")) {
		    completeStage("Sales Contract # " + scStr + " CANNOT BE FOUND. Double check that this is the correct SC#");
		} else {
		    completeStage("You cannot REINSTATE a " + status + " order!");
		}//if-elseif-else
	    });


	Button unbookingBtn = new Button("Unbook");
	unbookingBtn.setOnAction(e -> {
		String scStr = sc.getText();
		String status = Log.getOrderStatus(scStr);
		
		PasswordField passField = new PasswordField();
		passField.setPromptText("Password");
		Button checkPassBtn = new Button("Enter");		
		VBox passVBox = new VBox(20, new Label("Enter Password"), passField, checkPassBtn);
		passVBox.setStyle("-fx-background-color: linear-gradient(to bottom right, NavajoWhite, BurlyWood); "
				   + "-fx-spacing: 20px; -fx-padding: 40px;");
		Stage passStage = new Stage();
		passStage.setScene(new Scene(passVBox));
		passStage.sizeToScene();
		passStage.initModality(Modality.APPLICATION_MODAL);
		passStage.show();
		checkPassBtn.setOnAction(ev -> {
			if (passField.getText().equals("NEW")) {
			    if (status.equals("BOOKING")) {
				try {
				    File f = Log.findFile(scStr);
				    XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(f));
				    boolean areSheetsRemoved = true;
				    for (int i = 0; i < wb.getNumberOfSheets(); i++) {
					String sheetName = wb.getSheetAt(i).getSheetName();
					if (sheetName.startsWith("PI BOOKING") || sheetName.equals("CI-PL")) {
					    passStage.close();					    
					    completeStage("You must delete the following sheets before unbooking: \"PI BOOKING,\" \"CI-PL\".");
					    areSheetsRemoved = false;
					    break;
					}//if
				    }//for
			    
				    if (areSheetsRemoved) {
					stage = confirmation(eve -> {
						yes.disarm();
						MasterLog.appendEntry("Unbooking Sales Contract #" + scStr + "...");
						stage.setHeight(250);
						stage.show();

						Thread t = new Thread(() -> {			    
							Log.changeOrderStatus("CONFIRMED", sc.getText());
							Platform.runLater(() -> {
								stage.close();
								passStage.close();
								yes.arm();
								completeStage("Sales Contract #" + sc.getText() + " is now CONFIRMED. NOTE: CO must be manually updated in the future.");
								MasterLog.append("Sales Contract #" + sc.getText() + " is now CONFIRMED");
							    });
						    });
						t.setDaemon(true);
						t.start();
					    }
					    , "Are you sure you want to UNBOOK Sales Contract #"+ scStr);
					stage.show();
				    }
				} catch (IOException ioe) {
				    passStage.close();
				    completeStage("Sales Contract #" + scStr + " file could not be found");
				}//try-catch
		    

			    } else if (status.equals("")) {
				passStage.close();
				completeStage("Sales Contract # " + scStr + " CANNOT BE FOUND. Double check that this is the correct SC#");
			    } else {
				passStage.close();
				completeStage("You cannot UNBOOK a " + status + " order!");
			    }//if-elseif-else
			} else {
			    passStage.close();
			    completeStage("Incorrect Password");
			}//if-else
		    });
	    });


	Button createCIBtn = new Button("Create CI-PL");
	createCIBtn.setOnAction(e -> {
		String scStr = sc.getText();
		String status = Log.getOrderStatus(scStr);
		
		PasswordField passField = new PasswordField();
		passField.setPromptText("Password");
		Button checkPassBtn = new Button("Enter");		
		VBox passVBox = new VBox(20, new Label("Enter Password"), passField, checkPassBtn);
		passVBox.setStyle("-fx-background-color: linear-gradient(to bottom right, NavajoWhite, BurlyWood); "
				   + "-fx-spacing: 20px; -fx-padding: 40px;");
		Stage passStage = new Stage();
		passStage.setScene(new Scene(passVBox));
		passStage.sizeToScene();
		passStage.initModality(Modality.APPLICATION_MODAL);
		passStage.show();
		checkPassBtn.setOnAction(ev -> {
			if (passField.getText().equals("NEW") || passField.getText().equals("TEMP")) {
			    if (status.equals("BOOKING")) {
				
				    stage = confirmation(eve -> {
					    try {
						yes.disarm();
						stage.setHeight(250);
						MasterLog.appendEntry("Creating CI-PL for Sales Contract #" + scStr + "...");
						File f = Log.findFile(scStr);
						ExistingDocuments ed = new ExistingDocuments(f);
						ed.createCI();
						stage.close();
						passStage.close();
						completeStage("CI-PL created for Sales Contract #" + sc.getText());
						Desktop.getDesktop().open(f);
						MasterLog.append("CI-PL created for Sales Contract #" + sc.getText());
					    } catch (IOException ioe) {
						stage.close();
						passStage.close();
						completeStage("Sales Contract #" + scStr + " file could not be found");
					    } finally {
						yes.arm();
					    }//try-catch-finally
					}
					, "Are you sure you want to Create a CI-PL for SC#"+ scStr);
				    stage.show();
			    } else if (status.equals("")) {
				passStage.close();
				completeStage("Sales Contract # " + scStr + " CANNOT BE FOUND. Double check that this is the correct SC#");
			    } else {
				passStage.close();
				completeStage("You cannot create a CI-PL for a " + status + " order!");
			    }//if-elseif-else
			} else {
			    passStage.close();
			    completeStage("Incorrect Password");
			}//if-else
		    });
	    });	
		    
		    

	HBox top = new HBox(10);
	top.getChildren().addAll(confirmBtn, bookingBtn, unbookingBtn);
	HBox bot = new HBox(10);
	bot.getChildren().addAll(cancelBtn, reinstateBtn, createCIBtn);
	choices.getChildren().addAll(top, bot);

	allButTitle.getChildren().addAll(scvbox, choices);

	Label title = new Label("Update Order Status");
	title.setFont(new Font(40));
	title.setTextFill(Color.INDIGO);
	
	outer.getChildren().addAll(backBtn, title, allButTitle);
	return new Scene(outer);
    }//statusScene()


    public Scene assignScene() {
	Scene scene;
	Button 	assignBtn = new Button("Assign EHF #'s");

	VBox outer = new VBox(20);
	outer.setStyle("-fx-background-color: linear-gradient(to bottom right, NavajoWhite, BurlyWood); "
		       + "-fx-spacing: 20px; -fx-padding: 40px;");
	outer.setAlignment(Pos.CENTER);

	Label title = new Label("Assign EHF#'s");
	title.setFont(new Font(40));
	title.setTextFill(Color.INDIGO);

	Button anotherBtn = new Button("Add Row");
	anotherBtn.setOnAction(e -> {
		outer.getChildren().removeAll(anotherBtn, assignBtn);
		outer.getChildren().addAll(userEntryHBox(), anotherBtn, assignBtn);
		mainStage.sizeToScene();
	    });

	assignBtn.setDefaultButton(true);
	assignBtn.setOnAction(e -> {
		stage = confirmation(ev -> {
			yes.disarm();
			stage.setHeight(250);
			MasterLog.appendEntry("Assigning EHF#s...");
			boolean complete = true;
			for (int i = 0; i < outer.getChildren().size(); i++) {
			    HBox hbox = null;
			    TextField scText = null, poText = null;

		    
			    //allowing access to the textfields if multiple
			    if (outer.getChildren().get(i) instanceof HBox)
				hbox = (HBox)outer.getChildren().get(i);
			    if (hbox == null)
				continue;
		    
			    if (hbox.getChildren().get(1) instanceof TextField)
				poText = (TextField)hbox.getChildren().get(1);
			    if (hbox.getChildren().get(3) instanceof TextField)
				scText = (TextField)hbox.getChildren().get(3);

			    String po = poText.getText();
			    String sc = scText.getText();
			    try {
				if (Log.isValidPO(po, sc)) {
				    ExistingDocuments ed = new ExistingDocuments(Log.findFile(sc));
				    if (ed.insertPO(po).equals("")) {
				
				    } else {
					outer.getChildren().add(new Label(ed.insertPO(po)));
					mainStage.sizeToScene();
					complete = false;
				    }//if-else
				} else {
				    outer.getChildren().add(new Label("Invalid EHF#! This could be caused by: duplicate EHF#'s, too large/small EHF#, or the order is not CONFIRMED"));
				    mainStage.sizeToScene();			    
				    complete = false;
				}//if-else
				if (!CheckCompletion.checkAssign(sc))
				    complete = false;
				if (complete) {
				    Log.appendEHF(po, sc);
				    MasterLog.append("EHF " + po + " assigned to SC#" + sc);
				}//if
			    } catch (Exception ex) {
				MasterLog.appendError(ex);
				complete = false;
			    }//try-catch

			}//for
			stage.close();
			yes.arm();
			if (complete) {
			    completeStage("EHF#'s have been assigned.");
			    MasterLog.append("EHF#'s have been assigned.");		    
			} else { 
			    completeStage("ERROR: EHF#'s MAY NOT HAVE BEEN ASSIGNED");
			}//if-else
		    }, "Are you sure you want to assign these EHF#s?");
		stage.show();
	    });
	outer.getChildren().addAll(backBtn, title, userEntryHBox(), anotherBtn, assignBtn);
       
	scene = new Scene(outer);
	return scene;
    }//assignScene()

    private HBox userEntryHBox() {
	HBox hbox = new HBox(10);
	
	TextField po = new TextField();
	po.setPromptText("PO #");
	
	TextField sc = new TextField();
	sc.setPromptText("SC #");
	
	hbox.getChildren().addAll(new Label("Assign"), po, new Label("to"), sc);

	return hbox;
    }//userEntryHBox()

    public Scene addClientScene(boolean addingNew) {
	String[] consigneeArr=new String[9], notifyArr=new String[9], requirementArr = new String[18];
	BufferedReader br;
	VBox outer = new VBox(15);
	outer.setStyle("-fx-background-color: linear-gradient(to bottom right, NavajoWhite, BurlyWood); -fx-padding: 20px;");
	outer.setAlignment(Pos.CENTER);
	ScrollPane scrollPane = new ScrollPane(outer);
	scrollPane.setVbarPolicy(ScrollPane.ScrollBarPolicy.ALWAYS);
	scrollPane.setHbarPolicy(ScrollPane.ScrollBarPolicy.NEVER);

	Label title;
	if (addingNew) title = new Label("Add Client");
	else title = new Label("Update Client");
	title.setFont(new Font(30));
	title.setTextFill(Color.INDIGO);

	ComboBox<String> customerCB = new ComboBox<>();
	new AutoCompleteComboBox(customerCB);
	customerCB.getItems().add("Select Customer");
	customerCB.getSelectionModel().selectFirst();
	try { //adding all client names to choicebox
	    br = new BufferedReader(new FileReader("ADMIN/OTHER/maven/src/main/resources/file/Client List.txt"));
	    String client;
	    while ((client = br.readLine()) != null)
		customerCB.getItems().add(client);
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	}//try-catch
	
	ComboBox<String> countryCB = new ComboBox<>();
	new AutoCompleteComboBox(countryCB);	
	countryCB.getItems().add("Select Country");
	countryCB.getSelectionModel().selectFirst();
	try {
	    br = new BufferedReader(new FileReader("ADMIN/OTHER/maven/src/main/resources/file/Country List.txt"));
	    String currentCountry;
	    while ((currentCountry = br.readLine()) != null) {
		if (currentCountry.equals(""))
		    continue;
		countryCB.getItems().add(currentCountry);
	    }//while
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	}//try-catch

	TextField portField = new TextField();
	portField.setPromptText("Port of Discharge");

	ChoiceBox<String> repCB = new ChoiceBox<>();
	repCB.getItems().add("Select EHF Sales Rep");
	repCB.getSelectionModel().selectFirst();
	try {
	    br = new BufferedReader(new FileReader("ADMIN/OTHER/maven/src/main/resources/file/Sales Rep List.txt"));
	    String currentRep;
	    while ((currentRep = br.readLine()) != null) {
		if (currentRep.equals(""))
		    continue;
		repCB.getItems().add(currentRep.substring(0, currentRep.indexOf(" - ")));
	    }//while
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	}//try-catch

	TextField nameField = new TextField();
	nameField.setPromptText("Client Name");
	
	HBox top = new HBox(95);

	VBox consignee = new VBox(5);
	consignee.getChildren().add(new Label("Consignee:"));
	TextField[] consigneeFields = new TextField[9];
	for (int i = 0; i < consigneeFields.length; i++) {
	    consigneeFields[i] = new TextField();
	    if (i == 0)
		consigneeFields[i].setPromptText("Company Name");
	    else
		consigneeFields[i].setPromptText("Address Line " + i);
	    
	    consignee.getChildren().add(consigneeFields[i]);
	}//for

	VBox notify = new VBox(5);
	Button copyConsigneeBtn = new Button("Copy Consignee");
	copyConsigneeBtn.setFont(new Font(8));
	HBox header = new HBox(20, new Label("Notify:"), copyConsigneeBtn);
	notify.getChildren().add(header);	
	TextField[] notifyFields = new TextField[9];
	for (int i = 0; i < notifyFields.length; i++) {
	    notifyFields[i] = new TextField();
	    if (i == 0)
		notifyFields[i].setPromptText("Company Name");
	    else
		notifyFields[i].setPromptText("Address Line " + i);
	    notify.getChildren().add(notifyFields[i]);
	}//for
	copyConsigneeBtn.setOnAction(e -> {
		for (int i = 0; i < notifyFields.length; i++)
		    notifyFields[i].setText(consigneeFields[i].getText());
	    });	

	
	VBox requirements = new VBox(5);
	requirements.getChildren().add(new Label("Column Requirements:           Total"));
	HBox[] hboxes = new HBox[9];
	TextField[] requirementFields = new TextField[9];
	CheckBox[] totalChecked = new CheckBox[9];
	for (int i = 0; i < requirementFields.length; i++) {
	    requirementFields[i] = new TextField();
	    requirementFields[i].setPromptText("Requirement");

	    totalChecked[i] = new CheckBox();
	    
	    hboxes[i] = new HBox(20);

	    //no check box on last
	    if (i < 8) hboxes[i].getChildren().addAll(requirementFields[i], totalChecked[i]);
	    else hboxes[i].getChildren().add(requirementFields[i]);
	    
	    requirements.getChildren().add(hboxes[i]);
	}//for

	top.getChildren().addAll(consignee, notify, requirements);
	

	VBox notes = new VBox(5);
	HBox notesLabels = new HBox(664, new Label("Notes"));
	notes.getChildren().add(notesLabels);
	
	ArrayList<TextField> notesFields = new ArrayList<>();
	ArrayList<CheckBox> notesCheckBoxes = new ArrayList<>();	
	ArrayList<HBox> notesHboxes = new ArrayList<>();
	
	Button addNotesBtn = new Button("Add Notes");
	addNotesBtn.setOnAction(e -> {
		if (notesFields.size() == 0)
		    notesLabels.getChildren().add(new Label("Red"));
		for (int i = 0; i < 5; i++) {
		    TextField newField = new TextField();
		    CheckBox newCheckBox = new CheckBox();
		    HBox newHBox = new HBox(20);

		    newField.setPromptText("Note Line " + (notesFields.size() + 1));
		    newField.setMinWidth(portField.getWidth() - 50);
		    notesFields.add(newField);
		    notesCheckBoxes.add(newCheckBox);
		    newHBox.getChildren().addAll(newField, newCheckBox);
		    notesHboxes.add(newHBox);
		    
		    notes.getChildren().add(notes.getChildren().size() - 1, newHBox);
		}
	    });
	notes.getChildren().add(addNotesBtn);
	
		  
	TextField paymentField = new TextField();
	paymentField.setPromptText("Specify Payment Terms");

	if (!addingNew) {
	    customerCB.setOnAction(e -> {
		    if (!customerCB.getValue().equals("Select Customer") && validInput(customerCB)) {
			try {
			    XSSFSheet db = new XSSFWorkbook(new FileInputStream("ADMIN/OTHER/maven/src/main/resources/file/ClientDatabase.xlsx"))
				.getSheet("DB-Customers");
		
			    int column = CreateDocuments.findColumnIndex(db, customerCB.getValue(), 0);
		
			    XSSFCell cell;
			    cell = db.getRow(CreateDocuments.findRowIndex(db, "EHF Sales Rep", 0)).getCell(column);
			    repCB.setValue(cell.getStringCellValue());

			    cell = db.getRow(CreateDocuments.findRowIndex(db, "Country", 0)).getCell(column);
			    countryCB.setValue(cell.getStringCellValue());

			    try {
				cell = db.getRow(CreateDocuments.findRowIndex(db, "Port of Discharge", 0)).getCell(column);
				portField.setText(cell.getStringCellValue());
			    } catch (NullPointerException npe) {} //if not filled out
			    
		
			    for (int i = 0; i < consigneeFields.length; i++) {
				try {
				    cell = db.getRow(i+1).getCell(column);
				    consigneeFields[i].setText(cell.getStringCellValue());
				} catch (NullPointerException npe) {
				    consigneeFields[i].clear();
				}//try-catch
			    }
			    for (int i = 0; i < notifyFields.length; i++) {
				try {
				    cell = db.getRow(i + CreateDocuments.findRowIndex(db, "Notify", 0)).getCell(column);
				    notifyFields[i].setText(cell.getStringCellValue());
				} catch (NullPointerException npe) {
				    notifyFields[i].clear();
				}//try-catch			
			    }
			    int totalCount = 0;
			    for (int i = 0; i < requirementFields.length * 2; i++) {
				if (i - totalCount < requirementFields.length)
				    totalChecked[i-totalCount].setSelected(false);
			    
				try {
				    cell = db.getRow(i + CreateDocuments.findRowIndex(db, "Colums/Requirements", 0)).getCell(column);
				    if (cell.getStringCellValue().startsWith("TOTAL\n")) {
					totalCount++;
					totalChecked[i-totalCount].setSelected(true);
				    } else {
					if (i - totalCount >= requirementFields.length)
					    break;
					if (!cell.getStringCellValue().equals(""))
					    requirementFields[i-totalCount].setText(cell.getStringCellValue());
					else
					    requirementFields[i-totalCount].setText("");
				    }//if-else
				} catch (NullPointerException npe) {
				    if (i - totalCount >= requirementFields.length)
					break;
				    requirementFields[i-totalCount].clear();
				}//try-catch			    
			    }

			    cell = db.getRow(CreateDocuments.findRowIndex(db, "Notes", 0)).getCell(column);
			    int i = 0;
			    while (cell != null && cell.getRow() != null && !cell.getStringCellValue().equals("")) {
				if (notesFields.size() == 0)
				    addNotesBtn.fire();
				if (i == notesFields.size())
				    addNotesBtn.fire();				
				
				String note = cell.getStringCellValue();
				if (!note.endsWith("(RED)")) {
				    notesFields.get(i).setText(note);
				} else {
				    notesFields.get(i).setText(note.substring(0, note.indexOf("(RED)")));
				    notesCheckBoxes.get(i).setSelected(true);
				}
				    
				i++;
				cell = db.getRow(i + CreateDocuments.findRowIndex(db, "Notes", 0)).getCell(column);
			    }//while
			    for (int x = 0; (x + i) < notesFields.size(); x++) {
				notesFields.get(x + i).clear();
			    }//for
			
			    try {
				cell = db.getRow(CreateDocuments.findRowIndex(db, "Payment", 0)).getCell(column);
				paymentField.setText(cell.getStringCellValue());
			    } catch (NullPointerException npe) {
				paymentField.clear();
			    }//try-catch
			} catch (Exception ex) {
			    MasterLog.appendEntry(ex.toString());
			    StackTraceElement[] arr = ex.getStackTrace();
			    for (int i = 0; i < arr.length; i++)
				MasterLog.append(arr[i].toString());
			}//try-catch
		    }//if
		});
	}//if
	
	Button add;
	if (addingNew) add = new Button("Add Client");
	else add = new Button("Update Client");
	add.setDefaultButton(true);		
	add.setOnAction(e -> {
		String confirmationMsg = "Are you sure you want to add this Client?";
		if (!addingNew) confirmationMsg = "Are you sure you want to update this Client?";
		stage = confirmation(ev -> {
			yes.disarm();
			stage.setHeight(250);
			boolean allFilled = true;
		
			String name;
			if (addingNew) {
			    name = nameField.getText().replaceAll("\\.", "").trim();;
			    MasterLog.appendEntry("Adding New Client...");		    		    
			} else {
			    name = customerCB.getValue();
			    MasterLog.appendEntry("Updating Client " + name + "...");
			}//if-else

			if (!validInput(customerCB)) {
			    outer.getChildren().add(new Label("Please select a Valid Customer."));
			    allFilled = false;
			}
		
			if (repCB.getValue().equals("Select EHF Sales Rep")) {
			    outer.getChildren().add(new Label("Please Enter an EHF Sales Rep"));
			    allFilled = false;
			}
		
			if (name.equals("") || name.equals("Select Customer") || name.startsWith(" ") || name.endsWith(" ")) {
			    allFilled = false;
			    outer.getChildren().add(new Label("Please select a Client Name (Do not have trailing spaces)"));
			}

			if (allFilled) {
			    ExistingDocuments ed = new ExistingDocuments(new File("ADMIN/OTHER/maven/src/main/resources/file/ClientDatabase.xlsx"));
		
			    for (int i = 0; i < consigneeFields.length; i++)
				consigneeArr[i] = consigneeFields[i].getText();
			    for (int i = 0; i < notifyFields.length; i++)
				notifyArr[i] = notifyFields[i].getText();
		
			    for (int i = 0, j = 0; i < requirementFields.length; i++, j++) {//i is array index, j is textfield index
				requirementArr[j] = requirementFields[i].getText();
				if (totalChecked[i].isSelected())
				    requirementArr[++j] = "TOTAL\n" + requirementFields[i].getText();
			    }//for

			    String[] notesArr = new String[notesFields.size()];
			    for (int i = 0; i < notesFields.size(); i++)
				notesArr[i] = notesFields.get(i).getText();

			    boolean[] redArr = new boolean[notesCheckBoxes.size()];
			    for (int i = 0; i < redArr.length; i++)
				redArr[i] = notesCheckBoxes.get(i).isSelected();
		
		    
			    ed.addClient(addingNew, name, repCB.getValue(), countryCB.getValue(), portField.getText(), consigneeArr, notifyArr, requirementArr, notesArr, redArr, paymentField.getText());

			    if (addingNew) {
				File textFile = new File("ADMIN/OTHER/maven/src/main/resources/file/Client List.txt");
				BinaryInsert.sortedFileInsert(textFile, nameField.getText().replaceAll("\\.","").trim());
			    }
			    
			    if (addingNew) {
				completeStage("Added Client \"" + name + "\" to Database.");
				MasterLog.append("Added Client \"" + name + "\" to Database.");			
			    } else {
				completeStage("Updated Client \"" + name + "\" in Database.");
				MasterLog.append("Updated Client \"" + name + "\" in Database.");
			    }
			}//if
			stage.close();
			yes.arm();
		    }, confirmationMsg);
		stage.show();
	    });

	outer.getChildren().addAll(backBtn, title);
	if (addingNew)
	    outer.getChildren().add(nameField);
	else
	    outer.getChildren().add(customerCB);
	outer.getChildren().addAll(repCB, countryCB, portField, top, notes, paymentField, add);

	return new Scene(scrollPane);
    }//addClientScene()


    private Scene addFactoryScene(boolean addingNew) {
	Scene scene;

	VBox outer = new VBox(20);
	outer.setStyle("-fx-background-color: linear-gradient(to bottom right, NavajoWhite, BurlyWood); "
		       + "-fx-spacing: 20px; -fx-padding: 40px;");
	outer.setAlignment(Pos.CENTER);

	ComboBox<String> factoryCB = new ComboBox<>();
	new AutoCompleteComboBox(factoryCB);
	factoryCB.getItems().add("Select Factory");
	factoryCB.getSelectionModel().selectFirst();
	try { //adding all factory names to choicebox
	    BufferedReader br = new BufferedReader(new FileReader("ADMIN/OTHER/maven/src/main/resources/file/Factory List.txt"));
	    String factory;
	    while ((factory = br.readLine()) != null)
		factoryCB.getItems().add(factory);
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	}//try-catch

	

	Label title;
	if (addingNew) title = new Label("Add New Factory");
	else title = new Label("Update Existing Factory");
	title.setFont(new Font(30));
	title.setTextFill(Color.INDIGO);

	TextField nameField = new TextField();
	nameField.setPromptText("Enter Factory Name");

	TextField addressOneField = new TextField();
	addressOneField.setPromptText("Address 1");

	TextField addressTwoField = new TextField();
	addressTwoField.setPromptText("Address 2");

	TextField contactField = new TextField();
	contactField.setPromptText("Contact Name");

	ToggleGroup tGroup = new ToggleGroup();
	ToggleButton local = new ToggleButton("LOCAL");
	ToggleButton direct = new ToggleButton("DIRECT");
	local.setToggleGroup(tGroup);
	direct.setToggleGroup(tGroup);
	HBox shipPointHBox = new HBox(local, direct);




	TextField discountField = new TextField();
	discountField.setPromptText("Enter Discount (only number)");
	
	TextField toField = new TextField();
	toField.setPromptText("To Email Address");
	
	TextField ccField = new TextField();
	ccField.setPromptText("CC Email Address");
	
	TextField bccField = new TextField();
	bccField.setPromptText("BCC Email Address");
	
	VBox left = new VBox(20, addressOneField, addressTwoField, contactField, shipPointHBox);
	VBox right = new VBox(20, discountField, toField, ccField, bccField);

	HBox middle = new HBox(40, left, right);

	
	factoryCB.setOnAction(e -> {
		if (!factoryCB.getValue().equals("Select Factory") && validInput(factoryCB)) {
		    try {
			XSSFSheet db = new XSSFWorkbook(new FileInputStream("ADMIN/OTHER/maven/src/main/resources/file/FactoryDatabase.xlsx")).getSheet("DB-Factories");
			int row = CreateDocuments.findRowIndex(db, factoryCB.getValue(), 0);
			XSSFCell cell;
		
			cell = db.getRow(row).getCell(CreateDocuments.findColumnIndex(db, "ADDRESS 1", 0));
			addressOneField.setText(cell.getStringCellValue());

			cell = db.getRow(row).getCell(CreateDocuments.findColumnIndex(db, "ADDRESS 2", 0));
			addressTwoField.setText(cell.getStringCellValue());

			cell = db.getRow(row).getCell(CreateDocuments.findColumnIndex(db, "CONTACT", 0));
			contactField.setText(cell.getStringCellValue());

			cell = db.getRow(row).getCell(CreateDocuments.findColumnIndex(db, "SHIP-POINT", 0));
			if (cell.getStringCellValue().equals("LOCAL"))
			    local.setSelected(true);
			if (cell.getStringCellValue().equals("DIRECT"))
			    direct.setSelected(true);

			cell = db.getRow(row).getCell(CreateDocuments.findColumnIndex(db, "DISCOUNT", 0));
			discountField.setText("" + cell.getNumericCellValue());

			cell = db.getRow(row).getCell(CreateDocuments.findColumnIndex(db, "TO:", 0));
			toField.setText(cell.getStringCellValue());

			cell = db.getRow(row).getCell(CreateDocuments.findColumnIndex(db, "CC:", 0));
			ccField.setText(cell.getStringCellValue());

			cell = db.getRow(row).getCell(CreateDocuments.findColumnIndex(db, "BCC:", 0));
			bccField.setText(cell.getStringCellValue());
		    
		    } catch (Exception ex) {
			MasterLog.appendError(ex);
		    }//try-catch
		}//if
	    });

	Button addBtn;
	if (addingNew) addBtn = new Button("Add Factory");
	else addBtn = new Button("Update Factory");
	addBtn.setDefaultButton(true);		
	addBtn.setOnAction(e -> {
		String confirmationMsg = "Are you sure you want to add this Factory?";
		if (!addingNew) confirmationMsg = "Are you sure you want to update this Factory?";
		stage = confirmation(ev -> {
			yes.disarm();
			stage.setHeight(250);
			boolean allFilled = true;

			if (!MenuApp.validInput(factoryCB)) {
			    outer.getChildren().add(new Label("Please select a valid Factory."));
			    allFilled = false;
			}
		
			String name;
			if (addingNew) name = nameField.getText().replaceAll("\\.","").trim();
			else name = factoryCB.getValue();

			if (name.equals("") || name.equals("Select Factory") || name.startsWith(" ") || name.endsWith(" ")) {
			    outer.getChildren().add(new Label("Please select a Factory (Do not leave trailing spaces)"));
			    allFilled = false;
			}

			String shipPoint = null;
			if (local.isSelected())
			    shipPoint = "LOCAL";
			if (direct.isSelected())
			    shipPoint = "DIRECT";
			if (!local.isSelected() && !direct.isSelected()) {
			    outer.getChildren().add(new Label("Please select Local or Direct"));
			    allFilled = false;
			}

			double discount = 0;
			try {
			    discount = Double.parseDouble(discountField.getText());
			    if (discount > 1) {
				outer.getChildren().add(new Label("Please enter your percentage in decimal format"));
				allFilled = false;
			    }
			} catch (NumberFormatException nfe) {
			    outer.getChildren().add(new Label("Please enter only a number for Discount"));
			    allFilled = false;
			}//try-catch

			if (toField.getText().equals("")) {
			    outer.getChildren().add(new Label("Please enter a To Email Adress"));
			    allFilled = false;
			}

			if (allFilled) {
			    ExistingDocuments ed = new ExistingDocuments(new File("ADMIN/OTHER/maven/src/main/resources/file/FactoryDatabase.xlsx"));
			    ed.addFactory(addingNew, name, addressOneField.getText(), addressTwoField.getText(),
					  contactField.getText(), shipPoint, discount, toField.getText(), ccField.getText(),
					  bccField.getText());

			    if (addingNew) {
				File textFile = new File("ADMIN/OTHER/maven/src/main/resources/file/Factory List.txt");
				BinaryInsert.sortedFileInsert(textFile, name);
			    }//if
			    
			    stage.close();
			    yes.arm();
			    
			    if (addingNew) {
				completeStage("Added Factory \"" + name + "\" to Database");
				MasterLog.appendEntry("Added Factory \"" + name + "\" to Database");
			    } else {
				completeStage("Updated Factory \"" + name + "\" in Database.");
				MasterLog.appendEntry("Updated Factory \"" + name + "\" in Database.");
			    }
			}//if

		    }, confirmationMsg);
		stage.show();
	    });
	
	outer.getChildren().addAll(backBtn, title);
	if (addingNew) outer.getChildren().add(nameField);
	else outer.getChildren().add(factoryCB);
	
	outer.getChildren().addAll(middle, addBtn);

	scene = new Scene(outer);
	return scene;
	    
    }//addFactory(String, String, String, String, double, String, String, String)

    

    public Scene openSCScene() {
	VBox outer = new VBox(30);
	outer.setStyle("-fx-background-color: linear-gradient(to bottom right, NavajoWhite, BurlyWood); -fx-padding: 40px;");
	outer.setAlignment(Pos.CENTER);
	
	Label title = new Label("Open a Sales Contract");
	title.setFont(new Font(40));
	title.setTextFill(Color.INDIGO);
	
	TextField tf = new TextField();
	tf.setPromptText("Enter a SC# or a PO#");
	tf.setMaxWidth(200);
	
	Button openBtn = new Button("Open File");
	openBtn.setOnAction(e -> {
		try {
		    Desktop.getDesktop().open(Log.findFile(tf.getText()));
		} catch (IOException ex) {
		    MasterLog.appendError(ex);
		} catch (NullPointerException npe) {
		    completeStage("Sales Contract " + tf.getText() + " could not be found");
		}
	    });
	openBtn.setDefaultButton(true);	
	outer.getChildren().addAll(backBtn, title, tf, openBtn);
	return new Scene(outer);
    }

    public Scene updateModelsScene() {
	VBox outer = new VBox(20);
	outer.setAlignment(Pos.CENTER);
	outer.setStyle("-fx-background-color: linear-gradient(to bottom right, NavajoWhite, BurlyWood); -fx-padding: 40px;");

	ScrollPane scrollPane = new ScrollPane(outer);
	scrollPane.setVbarPolicy(ScrollPane.ScrollBarPolicy.ALWAYS);
	scrollPane.setHbarPolicy(ScrollPane.ScrollBarPolicy.NEVER);

	Label title = new Label("Update an Order's Models");
	title.setFont(new Font(40));
	title.setTextFill(Color.INDIGO);

	TextField scTextField = new TextField();
	scTextField.setPromptText("Sales Contract #");
	scTextField.setMaxWidth(150);
	
	ArrayList<HBox> modelHboxes = new ArrayList<HBox>();
	ArrayList<TextField[]> modelTextFields = new ArrayList<TextField[]>();

	modelHboxes.add(new HBox(20));
	
	modelTextFields.add(new TextField[4]);
	TextField[] tf1 = modelTextFields.get(0);
	tf1[0] = new TextField();
	tf1[0].setPromptText("MODEL NUMBER");
	tf1[1] = new TextField();
	tf1[1].setPromptText("COMPOSITION");
	tf1[2] = new TextField();
	tf1[2].setPromptText("FABRIC/FINISH");	
	tf1[3] = new TextField();
	tf1[3].setPromptText("QUANTITY/SETS");
	
	modelHboxes.get(0).getChildren().addAll(tf1[0], tf1[1], tf1[2], tf1[3]);
	
	Button addRowBtn = new Button("Add Row");
	addRowBtn.setOnAction(e -> {
		modelHboxes.add(new HBox(20));
		
		modelTextFields.add(new TextField[4]);
		TextField[] tf = modelTextFields.get(modelTextFields.size() - 1);
		tf[0] = new TextField();
		tf[0].setPromptText("MODEL NUMBER");
		tf[1] = new TextField();
		tf[1].setPromptText("COMPOSITION");
		tf[2] = new TextField();
		tf[2].setPromptText("FABRIC/FINISH");
		tf[3] = new TextField();
		tf[3].setPromptText("QUANTITY/SETS");

		HBox hbox = modelHboxes.get(modelHboxes.size() - 1);
		hbox.getChildren().addAll(tf[0], tf[1], tf[2], tf[3]);

		outer.getChildren().add(outer.getChildren().size() - 2, hbox); //need to add to index

		if (mainStage.getHeight() < (screenHeight - 200))
		    mainStage.sizeToScene();
	    });

	Button enterBtn = new Button("Enter Models");
	enterBtn.setDefaultButton(true); 	
	enterBtn.setOnAction(e -> {
		stage = confirmation(ev -> {
			yes.disarm();
			stage.setHeight(250);
			try {
			    int count = 0;
			    for (int i = 0; i < modelTextFields.size(); i++)
				if (!modelTextFields.get(i)[0].equals("")) count++;
			    String[][] infoArray = new String[count][4];
			    TextField[] t;
			    for (int i = 0; i < infoArray.length; i++) {
				t = modelTextFields.get(i);
				infoArray[i][0] = t[0].getText();
				infoArray[i][1] = t[1].getText();
				infoArray[i][2] = t[2].getText();
				infoArray[i][3] = t[3].getText();
			    }

			    String scStr = scTextField.getText();
			    Distribution.addModels(infoArray, scStr);
			    
			    completeStage("Added Models to Distribution Log");
			} catch (Exception ex) {
			    MasterLog.appendError(ex);
			}//try-catch
			stage.close();
			yes.arm();
		    }, "Are you sure you want to update these models?");
		stage.show();
	    });

	outer.getChildren().addAll(backBtn, title, scTextField, modelHboxes.get(0), addRowBtn, enterBtn);

	return new Scene(scrollPane);
    }//updateModelsScene()

    public Scene searchDistributionScene() {
	VBox outer = new VBox(20);
	outer.setAlignment(Pos.CENTER);
	outer.setStyle("-fx-background-color: linear-gradient(to bottom right, NavajoWhite, BurlyWood); -fx-padding: 40px;");

	ScrollPane scrollPane = new ScrollPane(outer);
	scrollPane.setVbarPolicy(ScrollPane.ScrollBarPolicy.ALWAYS);
	scrollPane.setHbarPolicy(ScrollPane.ScrollBarPolicy.NEVER);

	Scene scene = new Scene(scrollPane);

	Label title = new Label("Search Distribution Log");
	title.setFont(new Font(40));
	title.setTextFill(Color.INDIGO);

	Button updateShipDateBtn = new Button("Update Ship Dates");
	updateShipDateBtn.setOnAction(e -> {
		stage = confirmation(ev -> {
			yes.disarm();
			stage.setHeight(250);
			MasterLog.appendEntry("Updating Ship Dates in Distribution Log...");
			Distribution.refreshExcel();
			stage.close();
			completeStage("Updated Ship Dates in Distribution Log");
			MasterLog.appendEntry("Updated Ship Dates in Distribution Log");
			yes.arm();
		    }
		    , "This may take a few minutes, are you sure?");
		stage.show();
	    });


	ComboBox<String> customerCB = new ComboBox<>();
	new AutoCompleteComboBox(customerCB);
	customerCB.getItems().add("Select Customer");
	customerCB.getSelectionModel().selectFirst();
	try { //adding all client names to choicebox
	    BufferedReader br = new BufferedReader(new FileReader("ADMIN/OTHER/maven/src/main/resources/file/Client List.txt"));
	    String client;
	    while ((client = br.readLine()) != null)
		customerCB.getItems().add(client);
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	}//try-catch
	
	ComboBox<String> factoryCB = new ComboBox<>();
	new AutoCompleteComboBox(factoryCB);
	factoryCB.getItems().add("Select Factory");
	factoryCB.getSelectionModel().selectFirst();
	try { //adding all factory names to choicebox
	    BufferedReader br = new BufferedReader(new FileReader("ADMIN/OTHER/maven/src/main/resources/file/Factory List.txt"));
	    String factory;
	    while ((factory = br.readLine()) != null)
		factoryCB.getItems().add(factory);
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	}//try-catch

	

	HBox countryHBox = new HBox(20);

	ComboBox<String> countryCB1 = new ComboBox<String>();
	ComboBox<String> countryCB2 = new ComboBox<String>();
	ComboBox<String> countryCB3 = new ComboBox<String>();
	countryCB1.setMaxWidth(175);
	countryCB2.setMaxWidth(175);
	countryCB3.setMaxWidth(175);	

	BufferedReader br;
	new AutoCompleteComboBox(countryCB1);
	new AutoCompleteComboBox(countryCB2);
	new AutoCompleteComboBox(countryCB3);		    
	countryCB1.getItems().add("Select Country");
	countryCB2.getItems().add("Select Country");
	countryCB3.getItems().add("Select Country");

	countryCB1.getSelectionModel().selectFirst();
	countryCB2.getSelectionModel().selectFirst();
	countryCB3.getSelectionModel().selectFirst();	    
	    
	countryHBox.getChildren().addAll(countryCB1, countryCB2, countryCB3);
	countryHBox.setAlignment(Pos.CENTER);
	try {
	    br = new BufferedReader(new FileReader("ADMIN/OTHER/maven/src/main/resources/file/Country List.txt"));
	    String currentCountry;
	    while ((currentCountry = br.readLine()) != null) {
		if (currentCountry.equals(""))
		    continue;
		countryCB1.getItems().add(currentCountry);
		countryCB2.getItems().add(currentCountry);
		countryCB3.getItems().add(currentCountry);		    
	    }//while
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	}//try-catch

	customerCB.setOnAction(e -> { //populate first country ComboBox
		if (!customerCB.getValue().equals("Select Customer") && validInput(customerCB)) {
		    try {
			XSSFSheet db = new XSSFWorkbook(new FileInputStream("ADMIN/OTHER/maven/src/main/resources/file/ClientDatabase.xlsx"))
			    .getSheet("DB-Customers");
		
			int column = CreateDocuments.findColumnIndex(db, customerCB.getValue(), 0);
		
			XSSFCell cell;
			cell = db.getRow(CreateDocuments.findRowIndex(db, "Country", 0)).getCell(column);
			countryCB1.setValue(cell.getStringCellValue());
		    } catch (Exception ex) {
			MasterLog.appendError(ex);
		    }//try-catch
		}//if
	    });

	TextField modelTextField = new TextField();
	modelTextField.setPromptText("Model");
	modelTextField.setMaxWidth(200);

	HBox dateHBox = new HBox(20);
	dateHBox.setAlignment(Pos.CENTER);
	DatePicker minDatePicker = new DatePicker();
	minDatePicker.setValue(LocalDate.now().minusMonths(6));
	DatePicker maxDatePicker = new DatePicker();
	maxDatePicker.setValue(LocalDate.now());
	ToggleGroup tGroup = new ToggleGroup();
	ToggleButton po = new ToggleButton("PO Date");
	ToggleButton ship = new ToggleButton("Ship Date");
	po.setToggleGroup(tGroup);
	ship.setToggleGroup(tGroup);
	po.setMinWidth(80); po.setMaxWidth(80);
	ship.setMinWidth(80); ship.setMaxWidth(80);
	po.setSelected(true);
	VBox choiceVBox = new VBox(po, ship);
	dateHBox.getChildren().addAll(new Label("From:"), minDatePicker, new Label("To:"), maxDatePicker, choiceVBox);
	


	Button exportBtn = new Button("Export to Excel File");
	exportBtn.setOnAction(e -> {
		if (finalResults != null && finalResults.length > 0)
		    Distribution.exportExcel(finalResults);
		else
		    outer.getChildren().add(new Label("Make sure your Search has results before Exporting!"));
	    });
	
	Button searchBtn = new Button("Search");
	searchBtn.setDefaultButton(true);	
	searchBtn.setOnAction(e -> {
		try {
		MasterLog.appendEntry("Searching Distribution Log...");

		try {
		    outer.getChildren().remove(9, outer.getChildren().size());
		} catch (Exception ex) {
		}//try-catch
		outer.getChildren().add(exportBtn);		
		
		String customer = "", factory = "", country1 = "", country2 = "", country3 = "";
		if (!customerCB.getValue().equals("Select Customer"))
		    customer = customerCB.getValue();
		if (!factoryCB.getValue().equals("Select Factory"))
		    factory = factoryCB.getValue();
		if (!countryCB1.getValue().equals("Select Country"))
		    country1 = countryCB1.getValue();
		if (!countryCB2.getValue().equals("Select Country"))
		    country2 = countryCB2.getValue();
		if (!countryCB3.getValue().equals("Select Country"))
		    country3 = countryCB3.getValue();

		if (country1.equals("") && (!country2.equals("") || !country3.equals(""))) {
		    outer.getChildren().add(new Label("Please select first Country if you are searching for multiple"));
		    mainStage.sizeToScene();
		} else {
		    String[] dateResults = Distribution.resultingDateSearch(Distribution.toArray(),
									    Date.from(minDatePicker.getValue()
										      .atStartOfDay(ZoneId.systemDefault())
										      .toInstant()),
									    Date.from(maxDatePicker.getValue()
										      .atStartOfDay(ZoneId.systemDefault())
										      .toInstant()),
									    po.isSelected());		
		
		    finalResults = Distribution.resultingSearch(dateResults, factory, modelTextField.getText(),
									   customer, country1, country2, country3);
		

		    try {
			if (finalResults.length != 0) {
			    HBox hbox = new HBox(15);
			    VBox vbox;
			    for (int i = 0; i < finalResults[0].length; i++) {
				vbox = new VBox(15);
				String header = "";
				switch (i) {
				case 0: header = "FACTORY";break;
				case 1: header = "MODEL #";break;
				case 2: header = "COMPOSITION";break;
				case 3: header = "FABRIC/FINISH";break;
				case 4: header = "QTY/SETS";break;
				case 5: header = "CLIENT";break;
				case 6: header = "COUNTRY";break;
				case 7: header = "PO DATE";break;
				case 8: header = "SHIP DATE";break;
				case 9: header = "EHF#";break;
				case 10: header = "S/C #";break;
				}//switch
				vbox.getChildren().add(new Label(header));
				for (int x = 0; x < finalResults.length; x++) {
				    vbox.getChildren().add(new Label(finalResults[x][i]));
				}//for
				vbox.setAlignment(Pos.CENTER);
				hbox.getChildren().add(vbox);
			    }//for
			    outer.getChildren().add(hbox);

			    mainStage.sizeToScene();		    
			    if (mainStage.getHeight() > (screenHeight - 200))
				mainStage.setHeight(screenHeight - 200);
			} else {
			    outer.getChildren().add(new Label("Found 0 results. If you think this may be an error, please notify someone."));
			}//if-else
			MasterLog.append("Found " + finalResults.length + " results");
		    } catch (Exception ex) {
			MasterLog.appendError(ex);
		    }//try-catch
		}//if-else
		} catch (Exception exxxx) {
		    MasterLog.appendError(exxxx);
		}
	    });

	outer.getChildren().addAll(backBtn, title, updateShipDateBtn, customerCB, factoryCB, countryHBox, modelTextField, dateHBox, searchBtn);

	return scene;
    }//searchDistributionScene()

    public Scene listOrdersScene() {
	VBox outer = new VBox(20);
	outer.setAlignment(Pos.CENTER);
	outer.setStyle("-fx-background-color: linear-gradient(to bottom right, NavajoWhite, BurlyWood); -fx-padding: 40px;");

	ScrollPane scrollPane = new ScrollPane(outer);
	scrollPane.setVbarPolicy(ScrollPane.ScrollBarPolicy.ALWAYS);
	scrollPane.setHbarPolicy(ScrollPane.ScrollBarPolicy.NEVER);

	Scene scene = new Scene(scrollPane);	

	Label title = new Label("List Orders");
	title.setFont(new Font(40));
	title.setTextFill(Color.INDIGO);

	ComboBox<String> statusCB = new ComboBox<>();
	statusCB.getItems().add("Select Order Status");
	statusCB.getSelectionModel().selectFirst();

	statusCB.getItems().addAll("CANCELED", "REINSTATED", "PENDING", "CONFIRMED", "BOOKING", "SHIPPED");
	
	statusCB.setOnAction(e -> {
		MasterLog.appendEntry("Listing Orders...");
		if (outer.getChildren().size() == 4)
		    outer.getChildren().remove(3, 4);
		VBox resultVBox = new VBox(20);
		ArrayList<String> list = Log.listOfOrders(statusCB.getValue());
		for (String order : list)
		    resultVBox.getChildren().add(new Label(order));
		outer.getChildren().add(resultVBox);
		mainStage.sizeToScene();
	    });

	outer.getChildren().addAll(backBtn, title, statusCB);
	return scene;
    }

    public Scene addLeadScene() {
	File textFile = new File("ADMIN/OTHER/maven/src/main/resources/file/Lead List.txt");
	
	VBox outer = new VBox(20);
	outer.setAlignment(Pos.CENTER);
	outer.setStyle("-fx-background-color: linear-gradient(to bottom right, NavajoWhite, BurlyWood); -fx-padding: 40px;");

	ScrollPane scrollPane = new ScrollPane(outer);
	scrollPane.setVbarPolicy(ScrollPane.ScrollBarPolicy.ALWAYS);
	scrollPane.setHbarPolicy(ScrollPane.ScrollBarPolicy.NEVER);

	Label title = new Label("Add Leads");
	title.setFont(new Font(40));
	title.setTextFill(Color.INDIGO);

	ArrayList<HBox> leadHboxes = new ArrayList<HBox>();
	ArrayList<TextField[]> leadTextFields = new ArrayList<TextField[]>();
	ArrayList<ComboBox<String>> countryCBs = new ArrayList<>();

	leadHboxes.add(new HBox(20));
	
	leadTextFields.add(new TextField[3]);
	TextField[] tf1 = leadTextFields.get(0);
	tf1[0] = new TextField();
	tf1[0].setPromptText("EMAIL ADDRESS");
	tf1[1] = new TextField();
	tf1[1].setPromptText("COMPANY NAME");
	tf1[2] = new TextField();
	tf1[2].setPromptText("CONTACT NAME");

	ComboBox<String> countryCB = new ComboBox<>();
	new AutoCompleteComboBox(countryCB);	
	countryCB.getItems().add("Select Country");
	countryCB.getSelectionModel().selectFirst();
	try {
	    BufferedReader br = new BufferedReader(new FileReader("ADMIN/OTHER/maven/src/main/resources/file/Country List.txt"));
	    String currentCountry;
	    while ((currentCountry = br.readLine()) != null) {
		if (currentCountry.equals(""))
		    continue;
		countryCB.getItems().add(currentCountry);
	    }//while
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	}//try-catch		
	countryCB.setMaxWidth(175);
	countryCBs.add(countryCB);
	
	leadHboxes.get(0).getChildren().addAll(tf1[0], tf1[1], tf1[2], countryCB);
	
	Button addRowsBtn = new Button("Add Rows");
	addRowsBtn.setOnAction(e -> {
		for (int i = 0; i < 5; i++) {
		    leadHboxes.add(new HBox(20));
		
		    leadTextFields.add(new TextField[3]);
		    TextField[] tf = leadTextFields.get(leadTextFields.size() - 1);
		    tf[0] = new TextField();
		    tf[0].setPromptText("EMAIL ADDRESS");
		    tf[1] = new TextField();
		    tf[1].setPromptText("COMPANY NAME");
		    tf[2] = new TextField();
		    tf[2].setPromptText("CONTACT NAME");

		    HBox hbox = leadHboxes.get(leadHboxes.size() - 1);
		    
		    ComboBox<String> countryCBCopy = new ComboBox<>();
		    new AutoCompleteComboBox(countryCBCopy);
		    for (int x = 0; x < countryCB.getItems().size(); x++)
			countryCBCopy.getItems().add(countryCB.getItems().get(x));
		    countryCBCopy.getSelectionModel().selectFirst();
		    countryCBCopy.setMaxWidth(175);
		    countryCBs.add(countryCBCopy);

		    
		    hbox.getChildren().addAll(tf[0], tf[1], tf[2], countryCBCopy);

		    outer.getChildren().add(outer.getChildren().size() - 2, hbox); //need to add to index

		    if (mainStage.getHeight() < (screenHeight - 200))
			mainStage.sizeToScene();
		}
	    });

	Button enterBtn = new Button("Enter Leads");
	enterBtn.setDefaultButton(true);	
	enterBtn.setOnAction(e -> {
		stage = confirmation(ev -> {
			yes.disarm();
			stage.setHeight(250);
			try {
			    int count = 0;
			    for (int i = 0; i < leadTextFields.size(); i++)
				if (!leadTextFields.get(i)[2].getText().equals("")) count++;
			    String[][] infoArray = new String[count][4];
			    TextField[] t;
			    for (int i = 0; i < infoArray.length; i++) {
				t = leadTextFields.get(i);
				infoArray[i][0] = t[0].getText();
				infoArray[i][1] = t[1].getText();
				infoArray[i][2] = t[2].getText();
				infoArray[i][3] = countryCBs.get(i).getValue();
				if (!infoArray[i][0].equals("") && !infoArray[i][2].equals("") && validInput(countryCBs.get(i))) {
				    BinaryInsert.sortedFileInsert(textFile, infoArray[i][0] + " - Company Name: " + infoArray[i][1]
								  + " - Contact Name: " + infoArray[i][2] + " - Country: "
								  + infoArray[i][3]);
				    completeStage("Added Leads to Database");
				    MasterLog.appendEntry("Added Leads to Database");
				} else {
				    completeStage("Some Leads were not added (those without an email address, contact name, or valid country");
				    MasterLog.appendEntry("Some Leads were not added (those without an email address, contact name, or valid country");

				}//if-else
			    }
			} catch (Exception ex) {
			    MasterLog.appendError(ex);
			    completeStage("Failed to add Leads to Database. Text document may be empty");
			}//try-catch
			stage.close();
			yes.arm();
		    }, "Are you sure you want to add these Leads?");
		stage.show();
	    });
	outer.getChildren().addAll(backBtn, title, leadHboxes.get(0), addRowsBtn, enterBtn);

	return new Scene(scrollPane);	
    }//addLeadScene()

    public Scene emailLeadsScene() {
	VBox outer = new VBox(20);
	outer.setAlignment(Pos.CENTER);
	outer.setStyle("-fx-background-color: linear-gradient(to bottom right, NavajoWhite, BurlyWood); -fx-padding: 40px;");

	Label title = new Label("Email Leads");
	title.setFont(new Font(40));
	title.setTextFill(Color.INDIGO);

	ComboBox<String> countryCB = new ComboBox<>();
	new AutoCompleteComboBox(countryCB);	
	countryCB.getItems().add("Select Country");
	countryCB.getSelectionModel().selectFirst();
	try {
	    BufferedReader br = new BufferedReader(new FileReader("ADMIN/OTHER/maven/src/main/resources/file/Country List.txt"));
	    String currentCountry;
	    while ((currentCountry = br.readLine()) != null) {
		if (currentCountry.equals(""))
		    continue;
		countryCB.getItems().add(currentCountry);
	    }//while
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	}//try-catch		

	VBox msgVBox = new VBox(5);
	TextField subjectTextField = new TextField();
	subjectTextField.setPromptText("Subject Line");
	TextArea msgTextArea = new TextArea();
	msgTextArea.setPromptText("Message Body");
	msgTextArea.setMinHeight(250);
	Text fileName = new Text();
	Button removeAllBtn = new Button("Remove All Attachments");
	removeAllBtn.setFont(new Font(12));
	removeAllBtn.setOnAction(e -> {
		attachedFiles = new ArrayList<File>();
		msgVBox.getChildren().remove(3, msgVBox.getChildren().size());
		mainStage.sizeToScene();
	    });
	msgVBox.getChildren().addAll(subjectTextField, new Label("Dear _______,"), msgTextArea);

	Button selectFileBtn = new Button("Select Attachment");
	selectFileBtn.setOnAction(e -> {
		Stage stage = new Stage();
		stage.initModality(Modality.WINDOW_MODAL);
		FileChooser fileChooser = new FileChooser();
		fileChooser.setInitialDirectory(new File(".."));
		attachedFiles.add(fileChooser.showOpenDialog(stage));
		Label l = new Label("Attached File: " + attachedFiles.get(attachedFiles.size()-1).getName());
		if (msgVBox.getChildren().size() == 3) //if first file, add removeAll button
		    msgVBox.getChildren().addAll(l, removeAllBtn);
		else
		    msgVBox.getChildren().add(msgVBox.getChildren().size() - 1, l);
		mainStage.sizeToScene();
	    });


	Button sendBtn = new Button("Send Emails");
	sendBtn.setDefaultButton(true);	
	sendBtn.setOnAction(e -> {
		stage = confirmation(ev -> {
			yes.disarm();
			stage.setHeight(250);
			VBox vbox = new VBox(20);
			Label top = new Label("Enter Email Address and Password");
			title.setFont(new Font(20));
			title.setTextFill(Color.INDIGO);
			TextField userField = new TextField();
			userField.setPromptText("Email Address");
			PasswordField passField = new PasswordField();
			passField.setPromptText("Password");
			Button sendEmailBtn = new Button("Send Email");

			vbox.getChildren().addAll(top, userField, passField, sendEmailBtn);
			vbox.setAlignment(Pos.CENTER);
			vbox.setStyle("-fx-background-color: linear-gradient(to bottom right, NavajoWhite, BurlyWood); -fx-padding: 40px;");
		
			StackPane stackPane = new StackPane(vbox);
			Scene accountScene = new Scene(stackPane);
			Stage accountStage = new Stage();
			accountStage.setScene(accountScene);
			accountStage.initModality(Modality.APPLICATION_MODAL);
		
			sendEmailBtn.setOnAction(eve -> {
				if (vbox.getChildren().size() == 5) //has label of Incorrect Password
				    vbox.getChildren().remove(4, 5);
				stackPane.getChildren().add(new ProgressIndicator());
				Thread t = new Thread(() -> {
					MasterLog.appendEntry("Attempting to send Mass Email");				
					File[] arr = attachedFiles.toArray(new File[attachedFiles.size()]);
					long sizeSum = 0;
					for (int i = 0; i < arr.length; i++)
					    sizeSum += arr[i].length();
				
					String confirmation;				
					if (sizeSum < 15000000) {
					    if (Email.checkLogin(userField.getText(), passField.getText())) {
						if (Email.sendLeads(countryCB.getValue(), userField.getText(),
								    passField.getText(), subjectTextField.getText(),
								    msgTextArea.getText(), arr))
						    confirmation = "Leads have been emailed. Confirmation email sent to " + userField.getText();
						else
						    confirmation = "Some Leads have not been emailed. Confirmation email sent to " + userField.getText();
					    } else {
						confirmation = "Incorrect Username or Password";
					    }//if-else
					} else {
					    confirmation = "The attached files are greater than 15MB";    
					}//if-else
					Platform.runLater(() -> {
						if (confirmation.startsWith("Leads have been emailed.")) {
						    accountStage.close();						
						    completeStage(confirmation);
						} else if (confirmation.equals("Incorrect Username or Password")) {
						    stackPane.getChildren().remove(1, 2); //removing progress indicator
						    vbox.getChildren().add(new Label(confirmation));
						    accountStage.sizeToScene();
						} else {
						    accountStage.close();
						    outer.getChildren().add(new Label(confirmation));
						    mainStage.sizeToScene();						
						}//if-elseif-else
					    });
					MasterLog.append(confirmation);
				    });
				t.setDaemon(true);
				t.start();
			    });
		
			accountStage.sizeToScene();
			accountStage.show();
			stage.close();
			yes.arm();
		    }, "Are you sure you want to send this email?");
		stage.show();
	    });
	
	
	outer.getChildren().addAll(backBtn, title, countryCB, msgVBox, selectFileBtn, sendBtn);
	
	return new Scene(outer);
    }//emailLeadsScene()



    private Stage confirmation(EventHandler<ActionEvent> action, String text) {
	Stage stage = new Stage();
	VBox confirmationOuter = new VBox(20);
	confirmationOuter.setAlignment(Pos.CENTER);	

	confirmationOuter.setStyle("-fx-background-color: linear-gradient(to bottom right, NavajoWhite, BurlyWood); "
			  + "-fx-spacing: 20px; -fx-padding: 40px;");	
	
	Label label = new Label(text);

	HBox yesno = new HBox(30);
	yes.setOnAction(action);
	
	Button no = new Button("No");
	no.setOnAction(e -> stage.close());
	yesno.getChildren().addAll(yes, no);
	yesno.setAlignment(Pos.CENTER);
	
	confirmationOuter.getChildren().addAll(label, yesno, new ProgressIndicator());
	
	Scene scene = new Scene(confirmationOuter);
	stage.setScene(scene);
	stage.initModality(Modality.APPLICATION_MODAL);

	stage.setHeight(150);
	stage.setWidth(500);
	
	return stage;
    }//confirmation()

    private void completeStage(String text) {
	Stage stage = new Stage();
	VBox vbox = new VBox(30);
	vbox.setStyle("-fx-background-color: linear-gradient(to bottom right, NavajoWhite, BurlyWood);"
		      + "-fx-padding: 40px;");
	vbox.setAlignment(Pos.CENTER);
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
	try {
	    MasterLog.appendEntry("Main System Launched");	    
	    Application.launch(args);
	} catch (Exception ex) {
	    MasterLog.appendError(ex);
	}//try-catch
    }//main(String[])
    
}//MenuApp
