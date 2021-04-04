package com.leticia.ui;

import com.leticia.core.WordToXML;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.VBox;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import java.io.File;

public class Landing extends BorderPane {

    private static File file;

    public static File getFile() {
        return file;
    }

    public static void setFile(File file) {
        Landing.file = file;
    }

    public Landing(Stage stage) {

        this.setTop(instructionsTop());

        FileChooser fileChooser = new FileChooser();
        fileChooser.getExtensionFilters().addAll(
                new FileChooser.ExtensionFilter("Word document", "*.doc","*.docx")

        );
        fileChooser.setTitle("Choose word doc file");

        Button btnChooseFile = new Button("Select file");
        Button btnConvertWordToXml = new Button("Convert to XML");
        btnConvertWordToXml.setDisable(true);

        HBox boxBtns = new HBox(10);
        boxBtns.setAlignment(Pos.CENTER);
        boxBtns.getChildren().add(btnChooseFile);
        boxBtns.getChildren().add(btnConvertWordToXml);
        boxBtns.setPadding(new Insets(10,10,10,10));

        VBox vBox = new VBox(10);
        vBox.setAlignment(Pos.CENTER);
        vBox.getChildren().add(instructions());
        vBox.getChildren().add(boxBtns);

        this.setCenter(vBox);
        btnChooseFile.setOnAction(e -> {
            File chosenFile = fileChooser.showOpenDialog(stage);
           Landing.setFile(chosenFile);
           if (!Landing.getFile().getAbsolutePath().isEmpty()){
               btnConvertWordToXml.setDisable(false);
           }
            System.out.println("absolute file: " + Landing.getFile().getAbsoluteFile());
            System.out.println("absolute path: " + Landing.getFile().getAbsolutePath());
            System.out.println("path: " + Landing.getFile().getPath());
            System.out.println("name: " + Landing.getFile().getName());
            System.out.println("parent: " + Landing.getFile().getParentFile().getAbsolutePath());
        });

        btnConvertWordToXml.setOnAction(e -> {


//            String wordFile, String parentPath, String outName, String wordFileExtension
            File file = Landing.getFile();
            String theFile = Landing.getFile().getName();
            String fileExtension = theFile.substring(theFile.lastIndexOf(".") + 1, Landing.getFile().getName().length());
            System.out.println(">> fileExtension: " + fileExtension);
            String fileNameOnly = file.getName().substring(0,file.getName().indexOf("."));
            System.out.println("name only: " + fileNameOnly);

            WordToXML.convert(file.getAbsolutePath(), file.getParentFile().getAbsolutePath(), fileNameOnly,fileExtension);

            btnConvertWordToXml.setDisable(true);

        });

    }

    private HBox instructions() {
        Label label = new Label("Choose word document file to convert to XML");
        HBox box = new HBox();
        box.setPadding(new Insets(10, 10, 10, 10));
        box.getChildren().add(label);
        box.setAlignment(Pos.CENTER);
        return box;
    }

    private HBox instructionsTop() {
        Label label = new Label("The destination folder opens automatically upon conversion.");
        HBox box = new HBox();
        box.setPadding(new Insets(10, 10, 10, 10));
        box.getChildren().add(label);
        box.setAlignment(Pos.CENTER);
        return box;
    }

}
