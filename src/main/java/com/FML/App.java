package com.FML;

import javafx.application.Application;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.layout.VBox;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import java.io.File;

public class App extends Application {
    private static String fileName1;
    public static String fileName2;

    @Override
    public void start(Stage primaryStage){
        try {
            FileChooser fil_chooser = new FileChooser();

            Label label1 = new Label("Не выбран ни один файл");
            Label label2 = new Label("Не выбран ни один файл");

            Button button = new Button("Выберите PDF-файл");

            EventHandler<ActionEvent> event =
                    new EventHandler<ActionEvent>() {
                        public void handle(ActionEvent e)
                        {
                            File file = fil_chooser.showOpenDialog(primaryStage);
                            if (file != null) {
                                label1.setText(file.getAbsolutePath());
                                fileName1 = file.getAbsolutePath();
                            }
                        }
                    };

            button.setOnAction(event);
            Button button1 = new Button("Выберите xls файл");

            EventHandler<ActionEvent> event1 =
                    new EventHandler<ActionEvent>() {
                        public void handle(ActionEvent e)
                        {
                            File file = fil_chooser.showSaveDialog(primaryStage);
                            if (file != null) {
                                label2.setText(file.getAbsolutePath());
                                fileName2 = file.getAbsolutePath();
                            }
                        }
                    };

            button1.setOnAction(event1);

            Button button2 = new Button("Ок");

            EventHandler<ActionEvent> event2 =
                    new EventHandler<ActionEvent>() {
                        public void handle(ActionEvent e)
                        {
                            primaryStage.close();
                        }
                    };
            button2.setOnAction(event2);

            Label label3 = new Label("После нажатия кнопки 'ок' будет сформирован файл 'new', содержащий сравнение прикреплённых файлов. Файл располагается в той же папке, что и excel-документ");
            label3.setWrapText(true);

            VBox vbox = new VBox(30, label1, button, label2, button1, label3, button2);
            vbox.setAlignment(Pos.CENTER);

            Scene scene = new Scene(vbox, 400, 450);

            primaryStage.setScene(scene);
            primaryStage.show();
        }

        catch (Exception e) {
            System.out.println(e.getMessage());
        }
    }
    public static void main(String[] args) {
        launch(args);
        String bigString = PDF.StringFromPDF(App.fileName1);
        System.out.println(bigString);
        ExcelData.parsingExcel(App.fileName2, bigString);
    }

//    public static void run()
//    {
//        launch("");
//    }
}
