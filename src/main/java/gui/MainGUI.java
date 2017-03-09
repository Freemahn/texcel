package gui;

import console.ProjectForecastParser;
import javafx.application.Application;
import javafx.geometry.Insets;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.image.Image;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.Pane;
import javafx.scene.layout.VBox;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import java.io.File;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;

public class MainGUI extends Application {
    private ProjectForecastParser parser;

    private Logger logger = Logger.getLogger(MainGUI.class.getName());
    @Override
    public void start(Stage primaryStage) throws Exception{
        parser = new ProjectForecastParser();
        primaryStage.setTitle("Texcel");
        primaryStage.getIcons().add(new Image(getClass().getResource("../excel.png").openStream()));

        FileChooser fileChooser = new FileChooser();
        FileChooser.ExtensionFilter fileExtensions =
                new FileChooser.ExtensionFilter(
                        "Excel files", "*.xlsx", "*.xlsm", "*.xls");
        fileChooser.getExtensionFilters().add(fileExtensions);

        Button openButton = new Button("Load project world forecast");
        Button openMultipleButton = new Button("Load projects forecasts");
        openButton.setOnAction(
                e -> {
                    File file = fileChooser.showOpenDialog(primaryStage);
                    if (file != null) {
                        logger.log( Level.INFO , "Opened forecast projects world ", file);
                        ;//parse project forecast world
                    }
                });
        openMultipleButton.setOnAction(
                e -> {
                    List<File> list =
                            fileChooser.showOpenMultipleDialog(primaryStage);
                    logger.log( Level.INFO , "Opened forecast projects files ", list);
                    if (list != null) {
                        parser.parseForecastProjects(list);


                    }
                });


        GridPane inputGridPane = new GridPane();

        GridPane.setConstraints(openButton, 0, 0);
        GridPane.setConstraints(openMultipleButton, 1, 0);
        inputGridPane.setHgap(6);
        inputGridPane.setVgap(6);
        inputGridPane.getChildren().addAll(openButton, openMultipleButton);

        final Pane rootGroup = new VBox(12);
        rootGroup.getChildren().addAll(inputGridPane);
        rootGroup.setPadding(new Insets(12, 12, 12, 12));

        primaryStage.setScene(new Scene(rootGroup, 600, 400));
        primaryStage.show();
        logger.log( Level.SEVERE , "App started");
    }


    public static void main(String[] args) {
        launch(args);
    }


}
