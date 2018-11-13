import javafx.application.Application;
import javafx.beans.property.SimpleStringProperty;
import javafx.beans.value.ObservableValue;
import javafx.collections.ObservableList;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.geometry.Insets;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.control.Hyperlink;
import javafx.scene.control.Label;
import javafx.scene.control.Menu;
import javafx.scene.control.MenuBar;
import javafx.scene.control.MenuItem;
import javafx.scene.layout.*;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import javafx.util.Callback;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.awt.*;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.SQLException;
import java.util.ArrayList;


public class TableViewDemo extends Application {

    private Desktop desktop = Desktop.getDesktop();

    public void start(final Stage stage) throws Exception {

        final FileChooser fileChooser = new FileChooser();
        configuringFileChooser(fileChooser);

        final MenuBar menuBar = new MenuBar();
        final BorderPane root = new BorderPane();

        //Create menus
        Menu fileMenu = new Menu("File");

        //Create menu items
        MenuItem openFileItem = new MenuItem("Open File");
        MenuItem helpMenu = new MenuItem("Help");
        MenuItem aboutMenu = new MenuItem("About");
        MenuItem exitItem = new MenuItem("Exit");

        fileMenu.getItems().addAll(openFileItem, helpMenu, aboutMenu, exitItem);
        menuBar.getMenus().addAll(fileMenu);

        openFileItem.setOnAction(new EventHandler<ActionEvent>() {
            public void handle(ActionEvent event) {
                final HBox hBox = new HBox();
                final VBox dbBox = new VBox();
                final VBox tnBox = new VBox();
                final Label label = new Label();
                File file = fileChooser.showOpenDialog(stage);
                if (file != null){
                    openFile(file);

                    label.setText(file.getName());
                    openDB(file);
                    label.setPadding(new Insets(10));
                    dbBox.getChildren().add(label);

                    ArrayList nameTable = null;
                    try {
                        nameTable = DataBase.getNameTables();
                    } catch (SQLException e) {
                        e.printStackTrace();
                    }
                    final Hyperlink[] hlink = new Hyperlink[nameTable.size()];

                    for (int i = 0; i < nameTable.size(); i++) {
                        hlink[i] = new Hyperlink(nameTable.get(i).toString());
                        hlink[i].setPadding(new Insets(10));
                    }

                    for (int i = 0; i < nameTable.size() ; i++) {
                        tnBox.getChildren().add(hlink[i]);
                    }

                    for (int i = 0; i < nameTable.size(); i++) {
                        final int finalI = i;
                        hlink[i].setOnAction(new EventHandler<ActionEvent>() {
                            @Override
                            public void handle(ActionEvent event) {

                                final MenuBar menuBar = new MenuBar();
                                Menu fileMenu = new Menu("File");
                                MenuItem saveFileItem = new MenuItem("Save File");
                                MenuItem exitItem = new MenuItem("Exit");

                                fileMenu.getItems().addAll(saveFileItem, exitItem);
                                menuBar.getMenus().addAll(fileMenu);

                                final TableView table= new TableView();
                                ArrayList arrays = null;
                                try {
                                    arrays = DataBase.getColNameTable(hlink[finalI].getText());
                                } catch (SQLException e) {
                                    e.printStackTrace();
                                }
                                final TableColumn[] columns = new TableColumn[arrays.size()];
                                    //Create columns table
                                try {
                                    createColumnsTable(table, columns, hlink[finalI].getText(), arrays);
                                } catch (SQLException e) {
                                    e.printStackTrace();
                                }


                                AnchorPane anchorPane = new AnchorPane();

                                AnchorPane.setLeftAnchor(menuBar, 0.0);
                                AnchorPane.setRightAnchor(menuBar, 0.0);

                                anchorPane.getChildren().addAll(menuBar, table);

                                AnchorPane.setTopAnchor(table, ( 40.0));
                                AnchorPane.setLeftAnchor(table, 10.0);
                                AnchorPane.setRightAnchor(table, 10.0);
                                AnchorPane.setBottomAnchor(table, 10.0);

                                Scene newScene = new Scene(anchorPane, 400, 300);

                                // New window (Stage)
                                final Stage newWindow = new Stage();
                                newWindow.setTitle(hlink[finalI].getText());
                                newWindow.setScene(newScene);

                                // Set position of second window, related to primary window.
                                newWindow.setX(stage.getX() + 200);
                                newWindow.setY(stage.getY() + 100);

                                saveFileItem.setOnAction(new EventHandler<ActionEvent>() {
                                    @Override
                                    public void handle(ActionEvent event) {
                                        exportTableToXls(file, hlink, finalI, table, columns);

                                    }
                                });
                                exitItem.setOnAction(new EventHandler<ActionEvent>() {
                                    @Override
                                    public void handle(ActionEvent event) {
                                        newWindow.close();
                                    }
                                });

                                newWindow.show();
                            }
                        });
                    }
                    hBox.getChildren().addAll(dbBox,tnBox);
                    root.setCenter(hBox);

                }
            }
        });

        //click Exit
        exitItem.setOnAction(new EventHandler<ActionEvent>() {
            public void handle(ActionEvent event) {
                DataBase.disconnect();
                System.exit(0);
            }
        });

        aboutMenu.setOnAction(new EventHandler<ActionEvent>() {
            public void handle(ActionEvent event) {
                Alert alert = new Alert(Alert.AlertType.INFORMATION);
                alert.setTitle("About");
                alert.setContentText("Created by Drew.");
                alert.show();
            }
        });

        helpMenu.setOnAction(new EventHandler<ActionEvent>() {
            public void handle(ActionEvent event) {
                Alert alert = new Alert(Alert.AlertType.INFORMATION);
                alert.setTitle("Help");
                alert.setContentText("Эта программа создана для подключений к базам данных," +
                        " для просмотра их таблиц с последующим сохранением их в xls файл.");
                alert.showAndWait();
            }
        });

        root.setTop(menuBar);

        stage.setTitle("DataBase loader");
        final Scene scene = new Scene(root, 250, 500);
        stage.setScene(scene);
        stage.show();

    }

    private void exportTableToXls(File file, Hyperlink[] hlink, int finalI, TableView table, TableColumn[] columns) {
        Workbook workbook = null;
        HSSFSheet spreadsheet;
        File filexls = new File(file.getName().replace("db", "xls"));

        if (filexls.exists()){
            try {
                workbook = WorkbookFactory.create(filexls);
            } catch (IOException e) {
                e.printStackTrace();
            } catch (InvalidFormatException e) {
                e.printStackTrace();
            }
            for (int i = 0; i < workbook.getNumberOfSheets() ; i++) {
                if (workbook.getSheetName(i).equals(hlink[finalI].getText())){
                    workbook.removeSheetAt(i);
                }
            }
            spreadsheet = (HSSFSheet) workbook.createSheet(hlink[finalI].getText());


        }
        else{
            workbook = new HSSFWorkbook();
            spreadsheet = (HSSFSheet) workbook.createSheet(hlink[finalI].getText());
        }

        Row row = spreadsheet.createRow(0);
        HSSFCell cell;
        HSSFCellStyle style = createStyleForTitle((HSSFWorkbook) workbook);

        for (int j = 0; j < table.getColumns().size(); j++) {
            cell = (HSSFCell) row.createCell(j);
            cell.setCellValue(columns[j].getText());
            cell.setCellStyle(style);

        }

        for (int i = 0; i < table.getItems().size(); i++) {
            row = spreadsheet.createRow(i+1);
            for (int j = 0; j < table.getColumns().size(); j++) {
                if(columns[j].getCellObservableValue(i).getValue().toString() != null) {
                    row.createCell(j).setCellValue(columns[j].getCellObservableValue(i).getValue().toString());
                }
                else {
                    row.createCell(j).setCellValue("");
                }
            }
        }
        for (int i = 0; i < table.getColumns().size() ; i++) {
            spreadsheet.autoSizeColumn(i);
        }

        FileOutputStream fileOut;
        try {
            fileOut = new FileOutputStream(filexls);
            workbook.write(fileOut);
            fileOut.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    private TableView createColumnsTable(TableView table, TableColumn[] columns, String nameTable, ArrayList arrays) throws SQLException {

        for (int i = 0; i < arrays.size() ; i++) {
            final int j = i;
            columns[i] = new TableColumn((String) arrays.get(i));
            columns[i].setCellValueFactory(new Callback<TableColumn.CellDataFeatures<ObservableList, String>, ObservableValue<String>>() {
                public ObservableValue<String> call(TableColumn.CellDataFeatures<ObservableList, String> param) {
                    return new SimpleStringProperty(param.getValue().get(j).toString());
                }
            });
            table.getColumns().add(columns[i]);

            columns[i].setSortable(false);

        }
        ObservableList<ObservableList> data = DataBase.getFullTable(nameTable);
        table.setItems(data);
        table.setColumnResizePolicy(TableView.CONSTRAINED_RESIZE_POLICY);
        return table;
    }

    private void openDB(File file)  {
        try {
            DataBase.connection(file.getAbsolutePath());
        } catch (SQLException e) {
            e.printStackTrace();
        }
    }

    private void configuringFileChooser(FileChooser fileChooser) {
        fileChooser.setTitle("Select DataBase");

        fileChooser.getExtensionFilters().addAll(
                new FileChooser.ExtensionFilter("DB", "*.db"),
                new FileChooser.ExtensionFilter("All Files", "*.*"));
    }

    private void openFile(File file) {
        try {
            this.desktop.open(file);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static HSSFCellStyle createStyleForTitle(HSSFWorkbook workbook) {
        HSSFFont font = workbook.createFont();
        font.setBold(true);
        HSSFCellStyle style = workbook.createCellStyle();
        style.setFont(font);
        style.setAlignment(CellStyle.ALIGN_CENTER);
        return style;
    }

    public static void main(String[] args) {
        launch(args);
    }

    @Override
    public void stop() {
        DataBase.disconnect();
    }
}

