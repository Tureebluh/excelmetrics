package excelmetrics;

import java.io.File;
import java.io.FileOutputStream;
import java.net.URL;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.ResourceBundle;
import java.util.TreeMap;
import javafx.application.Platform;
import javafx.beans.property.SimpleIntegerProperty;
import javafx.beans.property.SimpleStringProperty;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.Button;
import javafx.scene.control.ScrollPane;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.scene.layout.BorderPane;
import javafx.stage.FileChooser;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class FXMLDocumentController implements Initializable {
    
    @FXML   private BorderPane borderPane;
    @FXML   private ScrollPane tableViewScrollPane;
    @FXML   private Button openFile;
            private File excelFile;
            private FileChooser fileChooser;
            private Boolean secondRow, firstRow;
            private ObservableList<DataRow> tableRows;
            private ObservableList<TableColumn> tableColumns;
            private TableView tableView;
            private TreeMap<String, List> dates;
            private int rowSize;
            private XSSFCellStyle style1,style2,style3,style4,style5,style6,style7,style8,style9,style10;
    
            
    @FXML
    private void calculateMetrics(ActionEvent event){
        if(excelFile != null){
            try{
                XSSFWorkbook workbook = new XSSFWorkbook(excelFile);
                XSSFSheet dataSheet = workbook.createSheet("Metrics");
                createStyles(dataSheet);
                createColumns(dataSheet);
                getUniqueDates(dataSheet);
                getAverageByDate(dataSheet);
                int extPos = excelFile.getPath().lastIndexOf(".");
                String fileNameMinusExt = excelFile.getPath().substring(0, extPos);
                String tempString =  fileNameMinusExt + "-with-metrics";
                FileOutputStream out = new FileOutputStream(tempString + ".xlsx");
                workbook.write(out);
                out.close();
                
                Alert alert = new Alert(AlertType.INFORMATION);
                alert.setTitle("Metrics Exported Successfully!");
                alert.setContentText(tempString + ".xlsx was exported successfully.");
                alert.showAndWait();
            }catch(Exception e){e.printStackTrace();}
        } else {
            try{
                openExcelFile(event);
            }catch(Exception e){e.printStackTrace();}
        }
    }
    
    @FXML
    private void openExcelFile(ActionEvent event) {
        //Open File Chooser Dialog and reset data
        excelFile = fileChooser.showOpenDialog(openFile.getScene().getWindow());
        tableRows = FXCollections.observableArrayList();
        tableColumns = FXCollections.observableArrayList();
        dates = new TreeMap<>();
        
        if(excelFile != null){
            try{
                XSSFWorkbook workbook = new XSSFWorkbook(excelFile);
                XSSFSheet dataSheet = workbook.getSheetAt(5);
                
                Iterator<Row> iterator = dataSheet.iterator();
                
                //Iterate through rows
                while (iterator.hasNext()) {
                    
                    ObservableList<String> row = FXCollections.observableArrayList();
                    Row nextRow = iterator.next();
                    Iterator<Cell> cellIterator = nextRow.cellIterator();
                    int columnCounter = 1;
                    
                    //Iterate through cells in row
                    while (cellIterator.hasNext()) {
                        Cell cell = cellIterator.next();
                        
                        //Add Columns to Table
                        if(secondRow){
                            TableColumn tempCol;
                            String tempString = " ";
                            
                            switch (cell.getCellType()) {
                                case Cell.CELL_TYPE_STRING:
                                    tempString = cell.getStringCellValue();
                                    break;
                                case Cell.CELL_TYPE_BOOLEAN:
                                    tempString = "" + cell.getBooleanCellValue();
                                    break;
                                case Cell.CELL_TYPE_NUMERIC:
                                    tempString = "" + cell.getNumericCellValue();
                                    break;
                                default:
                                    break;
                            }
                            tempCol = new TableColumn(tempString);
                            tempCol.setCellValueFactory(new PropertyValueFactory<>("column" + columnCounter));
                            tableColumns.add(tempCol);
                            columnCounter++;
                            
                            //Add Data to Table
                        } else if(!firstRow){
                            String tempString = "";
                            switch (cell.getCellType()) {
                                case Cell.CELL_TYPE_STRING:
                                    tempString = cell.getStringCellValue();
                                    break;
                                case Cell.CELL_TYPE_BOOLEAN:
                                    tempString = "" + cell.getBooleanCellValue();
                                    break;
                                case Cell.CELL_TYPE_NUMERIC:
                                    tempString = "" + cell.getNumericCellValue();
                                    break;
                                default:
                                    break;
                            }
                            row.add(tempString);
                        }
                    }//Cell iterator
                    
                    if(firstRow){
                        secondRow = true;
                        firstRow = false;
                    } else if(secondRow) {
                        secondRow = false;
                    } else {
                        DataRow tempRow = new DataRow();
                        rowSize = row.size();
                        
                        for(int i = 0; i < rowSize; i++){
                            switch(i + 1){
                                case 1:
                                    tempRow.setColumn1(row.get(i));
                                    break;
                                case 2:
                                    tempRow.setColumn2(row.get(i));
                                    break;
                                case 3:
                                    tempRow.setColumn3(row.get(i));
                                    break;
                                case 4:
                                    tempRow.setColumn4(row.get(i));
                                    break;
                                case 5:
                                    tempRow.setColumn5(row.get(i));
                                    break;
                                case 6:
                                    tempRow.setColumn6(row.get(i));
                                    break;
                                case 7:
                                    tempRow.setColumn7(row.get(i));
                                    break;
                                case 8:
                                    tempRow.setColumn8(row.get(i));
                                    break;
                                case 9:
                                    tempRow.setColumn9(row.get(i));
                                    break;
                                case 10:
                                    tempRow.setColumn10(row.get(i));
                                    break;
                                case 11:
                                    tempRow.setColumn11(row.get(i));
                                    break;
                                case 12:
                                    tempRow.setColumn12(row.get(i));
                                    break;
                                case 13:
                                    tempRow.setColumn13(row.get(i));
                                    break;
                                case 14:
                                    tempRow.setColumn14(row.get(i));
                                    break;
                                case 15:
                                    tempRow.setColumn15(row.get(i));
                                    break;
                                case 16:
                                    tempRow.setColumn16(row.get(i));
                                    break;
                                case 17:
                                    tempRow.setColumn17(row.get(i));
                                    break;
                                case 18:
                                    tempRow.setColumn18(row.get(i));
                                    break;
                                case 19:
                                    tempRow.setColumn19(row.get(i));
                                    break;
                                case 20:
                                    tempRow.setColumn20(row.get(i));
                                    break;
                                case 21:
                                    tempRow.setColumn21(row.get(i));
                                    break;
                                case 22:
                                    tempRow.setColumn22(row.get(i));
                                    break;
                                case 23:
                                    tempRow.setColumn23(row.get(i));
                                    break;
                                case 24:
                                    tempRow.setColumn24(row.get(i));
                                    break;
                                default:
                                    break;
                            }
                        }
                        tableRows.add(tempRow);
                    }
                }//Row iterator
                Platform.runLater(()->{
                    tableView.setItems(tableRows);
                    tableView.getColumns().addAll(tableColumns);
                });
                
                workbook.close();
            }catch(Exception e) {e.printStackTrace();}
        }
    }
    
    public void getUniqueDates(XSSFSheet sheet){
        for(int k = 0; k < tableRows.size(); k++){
            String temp = "";
            temp = tableRows.get(k).getColumn1();
            if(dates.containsKey(tableRows.get(k).getColumn1())){
                dates.get(tableRows.get(k).getColumn1()).add(tableRows.get(k));
            } else {
                ArrayList tempList = new ArrayList();
                tempList.add(tableRows.get(k));
                dates.put(tableRows.get(k).getColumn1(), tempList);
            }
        }
    }
    
    
    public void getAverageByDate(XSSFSheet sheet){
        XSSFWorkbook workbook = sheet.getWorkbook();
        
        XSSFCellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        
        XSSFCellStyle dateStyle = workbook.createCellStyle();
        dateStyle.setAlignment(HorizontalAlignment.CENTER);
        dateStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        
        DataMetrics dataMetrics = new DataMetrics();
        
        int i = 0;
        for(List<DataRow> tempList : dates.values()){ //Iterate through each map of dates
            
            XSSFRow row = sheet.createRow(i+1);
            
            for(int k = 0; k < rowSize; k++){ //Iterate through each row column
                
                Boolean doAvgColumn = false;
                int columnCounter = 0;
                int columnTotal = 0;
                String otherValues = "";
                
                for(DataRow tempRow : tempList){//Iterate through each date arraylist
                    
                    String temp = "";
                    
                    switch(k+1){
                        case 1:
                            temp = tempRow.getColumn1();
                            break;
                        case 2:
                            temp = tempRow.getColumn2();
                            break;
                        case 3:
                            temp = tempRow.getColumn3();
                            break;
                        case 4:
                            temp = tempRow.getColumn4();
                            break;
                        case 5:
                            temp = tempRow.getColumn5();
                            break;
                        case 6:
                            temp = tempRow.getColumn6();
                            break;
                        case 7:
                            temp = tempRow.getColumn7();
                            break;
                        case 8:
                            temp = tempRow.getColumn8();
                            break;
                        case 9:
                            temp = tempRow.getColumn9();
                            break;
                        case 10:
                            temp = tempRow.getColumn10();
                            break;
                        case 11:
                            temp = tempRow.getColumn11();
                            break;
                        case 12:
                            temp = tempRow.getColumn12();
                            break;
                        case 13:
                            temp = tempRow.getColumn13();
                            break;
                        case 14:
                            temp = tempRow.getColumn14();
                            break;
                        case 15:
                            temp = tempRow.getColumn15();
                            break;
                        case 16:
                            temp = tempRow.getColumn16();
                            break;
                        case 17:
                            temp = tempRow.getColumn17();
                            break;
                        case 18:
                            temp = tempRow.getColumn18();
                            break;
                        case 19:
                            temp = tempRow.getColumn19();
                            break;
                        case 20:
                            temp = tempRow.getColumn20();
                            break;
                        case 21:
                            temp = tempRow.getColumn21();
                            break;
                        case 22:
                            temp = tempRow.getColumn22();
                            break;
                        case 23:
                            temp = tempRow.getColumn23();
                            break;
                        case 24:
                            temp = tempRow.getColumn24();
                            break;
                    }
                    if(temp.toLowerCase().equals("y") || temp.toLowerCase().equals("n")){
                        doAvgColumn = true;
                        columnTotal++;
                        if(temp.toLowerCase().equals("y")){
                            columnCounter++;
                        }
                    } else {
                        otherValues = temp;
                    }
                }//Iterate through each object for that column
                
                Cell cell = row.createCell(k);
                cell.setCellStyle(style);
                double average = 0;
                if(doAvgColumn){
                    average = ((float)columnCounter / columnTotal) * 100;
                    cell.setCellValue((String)(columnCounter + "/" + columnTotal + "  " + (int)average + "%"));
                }else if(k == 0) { // Column 1
                    cell.setCellType(CellType.NUMERIC);
                    short df = workbook.createDataFormat().getFormat("dd-mmm");
                    dateStyle.setDataFormat(df);
                    cell.setCellValue((int)(Double.parseDouble(otherValues))); //Change this to Time
                    cell.setCellStyle(dateStyle);
                } else { // Column 2 & 3 respectively
                    cell.setCellValue("");
                }
                if(columnTotal != 0){
                    switch(k + 1){
                        case 9:
                            dataMetrics.setColumn9Avg((int)average);
                            dataMetrics.setColumn9Count(columnCounter);
                            dataMetrics.setColumn9Total(columnTotal);
                            break;
                        case 10:
                            dataMetrics.setColumn10Avg((int)average);
                            dataMetrics.setColumn10Count(columnCounter);
                            dataMetrics.setColumn10Total(columnTotal);
                            break;
                        case 11:
                            dataMetrics.setColumn11Avg((int)average);
                            dataMetrics.setColumn11Count(columnCounter);
                            dataMetrics.setColumn11Total(columnTotal);
                            break;
                        case 12:
                            dataMetrics.setColumn12Avg((int)average);
                            dataMetrics.setColumn12Count(columnCounter);
                            dataMetrics.setColumn12Total(columnTotal);
                            break;
                        case 13:
                            dataMetrics.setColumn13Avg((int)average);
                            dataMetrics.setColumn13Count(columnCounter);
                            dataMetrics.setColumn13Total(columnTotal);
                            break;
                        case 14:
                            dataMetrics.setColumn14Avg((int)average);
                            dataMetrics.setColumn14Count(columnCounter);
                            dataMetrics.setColumn14Total(columnTotal);
                            break;
                        case 15:
                            dataMetrics.setColumn15Avg((int)average);
                            dataMetrics.setColumn15Count(columnCounter);
                            dataMetrics.setColumn15Total(columnTotal);
                            break;
                        case 18:
                            dataMetrics.setColumn18Avg((int)average);
                            dataMetrics.setColumn18Count(columnCounter);
                            dataMetrics.setColumn18Total(columnTotal);
                            break;
                        case 19:
                            dataMetrics.setColumn19Avg((int)average);
                            dataMetrics.setColumn19Count(columnCounter);
                            dataMetrics.setColumn19Total(columnTotal);
                            break;
                        case 21:
                            dataMetrics.setColumn21Avg((int)average);
                            dataMetrics.setColumn21Count(columnCounter);
                            dataMetrics.setColumn21Total(columnTotal);
                            break;
                        case 22:
                            dataMetrics.setColumn22Avg((int)average);
                            dataMetrics.setColumn22Count(columnCounter);
                            dataMetrics.setColumn22Total(columnTotal);
                            break;
                        case 23:
                            dataMetrics.setColumn23Avg((int)average);
                            dataMetrics.setColumn23Count(columnCounter);
                            dataMetrics.setColumn23Total(columnTotal);
                            break;
                        default:
                            break;
                    }
                }   
            }//endof Iterate through each column
            dataMetrics.setColumn4Total(dataMetrics.getColumn9Total() + dataMetrics.getColumn10Total() + dataMetrics.getColumn15Total() +
                    dataMetrics.getColumn21Total() + dataMetrics.getColumn22Total() + dataMetrics.getColumn23Total());
            
            dataMetrics.setColumn4Count(dataMetrics.getColumn9Count() + dataMetrics.getColumn10Count() + dataMetrics.getColumn15Count() +
                    dataMetrics.getColumn21Count() + dataMetrics.getColumn22Count() + dataMetrics.getColumn23Count());
            
            dataMetrics.setColumn4Avg((dataMetrics.getColumn9Avg() + dataMetrics.getColumn10Avg() + dataMetrics.getColumn15Avg() +
                    dataMetrics.getColumn21Avg() + dataMetrics.getColumn22Avg() + dataMetrics.getColumn23Avg()) / 6);
            
            dataMetrics.setColumn6Total(dataMetrics.getColumn11Total() + dataMetrics.getColumn13Total() + dataMetrics.getColumn18Total() + 
                    dataMetrics.getColumn19Total() + dataMetrics.getColumn21Total() + dataMetrics.getColumn22Total() + dataMetrics.getColumn23Total());
            
            dataMetrics.setColumn6Count(dataMetrics.getColumn11Count() + dataMetrics.getColumn13Count() + dataMetrics.getColumn18Count() + 
                    dataMetrics.getColumn19Count() + dataMetrics.getColumn21Count() + dataMetrics.getColumn22Count() + dataMetrics.getColumn23Count());
            
            dataMetrics.setColumn6Avg((dataMetrics.getColumn11Avg() + dataMetrics.getColumn13Avg() + dataMetrics.getColumn18Avg() + 
                    dataMetrics.getColumn19Avg() + dataMetrics.getColumn21Avg() + dataMetrics.getColumn22Avg() + dataMetrics.getColumn23Avg()) / 7);
            
            row.getCell(3).setCellValue(dataMetrics.getColumn4Count() + "/" + dataMetrics.getColumn4Total() + "  " + dataMetrics.getColumn4Avg() + "%");
            row.getCell(5).setCellValue(dataMetrics.getColumn6Count() + "/" + dataMetrics.getColumn6Total() + "  " + dataMetrics.getColumn6Avg() + "%");
            
            i++;
            
        }//endof Iterate through date collections
    }//endof getAverageByDate
    
    /*
                
    */
    
    public void createColumns(XSSFSheet sheet){
        XSSFRow row = sheet.createRow(0);
        XSSFWorkbook workbook = sheet.getWorkbook();
        
        XSSFFont font = workbook.createFont();
        font.setFontHeightInPoints((short)12);
        font.setFontName("IMPACT");
        
        XSSFCellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setWrapText(true);
        style.setFont(font);
        
        row.setHeightInPoints(50);
        int cellId = 0;
        for(TableColumn column : tableColumns){
            Cell cell = row.createCell(cellId);
            cell.setCellValue((String)column.getText());
            sheet.setColumnWidth(cellId, 4000);
            
            switch(cellId){
                case 0:
                    cell.setCellStyle(style1);
                    cell.getCellStyle().setFont(font);
                    break;
                case 3:
                    cell.setCellStyle(style2);
                    cell.getCellStyle().setFont(font);
                    break;
                case 4:
                    cell.setCellStyle(style3);
                    cell.getCellStyle().setFont(font);
                    break;
                case 5:
                case 10:
                case 12:
                case 17:
                case 18:
                case 20:
                case 21:
                case 22:
                    cell.setCellStyle(style4);
                    cell.getCellStyle().setFont(font);
                    break;
                case 6:
                case 7:
                    cell.setCellStyle(style5);
                    cell.getCellStyle().setFont(font);
                    break;
                case 8:
                case 9:
                case 11:
                    cell.setCellStyle(style6);
                    cell.getCellStyle().setFont(font);
                    break;
                case 13:
                    cell.setCellStyle(style7);
                    cell.getCellStyle().setFont(font);
                    break;
                case 14:
                case 15:
                case 16:
                    cell.setCellStyle(style8);
                    cell.getCellStyle().setFont(font);
                    break;
                case 19:
                    cell.setCellStyle(style9);
                    cell.getCellStyle().setFont(font);
                    break;
                case 23:
                    cell.setCellStyle(style10);
                    cell.getCellStyle().setFont(font);
                    break;
                default:
                    cell.setCellStyle(style);
                    cell.getCellStyle().setFont(font);
                    break;
            }
            
            cellId++;
        }
    }
    
    public void createStyles(XSSFSheet sheet){
        style1 = sheet.getWorkbook().createCellStyle();
        style1.setAlignment(HorizontalAlignment.CENTER);
        style1.setVerticalAlignment(VerticalAlignment.CENTER);
        style1.setFillForegroundColor(IndexedColors.AQUA.getIndex());
        style1.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style1.setWrapText(true);
        
        style2 = sheet.getWorkbook().createCellStyle();
        style2.setAlignment(HorizontalAlignment.CENTER);
        style2.setVerticalAlignment(VerticalAlignment.CENTER);
        style2.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        style2.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style2.setWrapText(true);
        
        style3 = sheet.getWorkbook().createCellStyle();
        style3.setAlignment(HorizontalAlignment.CENTER);
        style3.setVerticalAlignment(VerticalAlignment.CENTER);
        style3.setFillForegroundColor(IndexedColors.PLUM.getIndex());
        style3.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style3.setWrapText(true);
        
        style4 = sheet.getWorkbook().createCellStyle();
        style4.setAlignment(HorizontalAlignment.CENTER);
        style4.setVerticalAlignment(VerticalAlignment.CENTER);
        style4.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        style4.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style4.setWrapText(true);
        
        style5 = sheet.getWorkbook().createCellStyle();
        style5.setAlignment(HorizontalAlignment.CENTER);
        style5.setVerticalAlignment(VerticalAlignment.CENTER);
        style5.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
        style5.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style5.setWrapText(true);
        
        style6 = sheet.getWorkbook().createCellStyle();
        style6.setAlignment(HorizontalAlignment.CENTER);
        style6.setVerticalAlignment(VerticalAlignment.CENTER);
        style6.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
        style6.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style6.setWrapText(true);
        
        style7 = sheet.getWorkbook().createCellStyle();
        style7.setAlignment(HorizontalAlignment.CENTER);
        style7.setVerticalAlignment(VerticalAlignment.CENTER);
        style7.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style7.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style7.setWrapText(true);
        
        style8 = sheet.getWorkbook().createCellStyle();
        style8.setAlignment(HorizontalAlignment.CENTER);
        style8.setVerticalAlignment(VerticalAlignment.CENTER);
        style8.setFillForegroundColor(IndexedColors.ROSE.getIndex());
        style8.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style8.setWrapText(true);
        
        style9 = sheet.getWorkbook().createCellStyle();
        style9.setAlignment(HorizontalAlignment.CENTER);
        style9.setVerticalAlignment(VerticalAlignment.CENTER);
        style9.setFillForegroundColor(IndexedColors.BROWN.getIndex());
        style9.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style9.setWrapText(true);
        
        style10 = sheet.getWorkbook().createCellStyle();
        style10.setAlignment(HorizontalAlignment.CENTER);
        style10.setVerticalAlignment(VerticalAlignment.CENTER);
        style10.setFillForegroundColor(IndexedColors.LIGHT_TURQUOISE.getIndex());
        style10.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style10.setWrapText(true);
    }
    
    @Override
    public void initialize(URL url, ResourceBundle rb) {
        dates = new TreeMap<>();
        tableView = new TableView<>();
        tableView.setEditable(false);
        borderPane.setCenter(tableView);
        tableRows = FXCollections.observableArrayList();
        firstRow = true;
        secondRow = false;
        fileChooser = new FileChooser();
        fileChooser.setTitle("Choose Excel File...");
        fileChooser.getExtensionFilters().addAll(
            new FileChooser.ExtensionFilter("XLSX", "*.xlsx")
        );
    }

    
    public static class DataMetrics {
        private SimpleIntegerProperty column4Total,column4Count,column4Avg;
        private SimpleIntegerProperty column6Total,column6Count,column6Avg;
        private SimpleIntegerProperty column9Total,column9Count,column9Avg;
        private SimpleIntegerProperty column10Total,column10Count,column10Avg;
        private SimpleIntegerProperty column11Total,column11Count,column11Avg;
        private SimpleIntegerProperty column12Total,column12Count,column12Avg;
        private SimpleIntegerProperty column13Total,column13Count,column13Avg;
        private SimpleIntegerProperty column14Total,column14Count,column14Avg;
        private SimpleIntegerProperty column15Total,column15Count,column15Avg;
        private SimpleIntegerProperty column16Total,column16Count,column16Avg;
        private SimpleIntegerProperty column17Total,column17Count,column17Avg;
        private SimpleIntegerProperty column18Total,column18Count,column18Avg;
        private SimpleIntegerProperty column19Total,column19Count,column19Avg;
        private SimpleIntegerProperty column20Total,column20Count,column20Avg;
        private SimpleIntegerProperty column21Total,column21Count,column21Avg;
        private SimpleIntegerProperty column22Total,column22Count,column22Avg;
        private SimpleIntegerProperty column23Total,column23Count,column23Avg;
        private SimpleIntegerProperty column24Total,column24Count,column24Avg;
        
        private DataMetrics(){
            
            column4Total = new SimpleIntegerProperty(0);
            column4Count = new SimpleIntegerProperty(0);
            column4Avg = new SimpleIntegerProperty(0);
            
            column6Total = new SimpleIntegerProperty(0);
            column6Count = new SimpleIntegerProperty(0);
            column6Avg = new SimpleIntegerProperty(0);
            
            column9Total = new SimpleIntegerProperty(0);
            column9Count = new SimpleIntegerProperty(0);
            column9Avg = new SimpleIntegerProperty(0);
            
            column10Total = new SimpleIntegerProperty(0);
            column10Count = new SimpleIntegerProperty(0);
            column10Avg = new SimpleIntegerProperty(0);
            
            column11Total = new SimpleIntegerProperty(0);
            column11Count = new SimpleIntegerProperty(0);
            column11Avg = new SimpleIntegerProperty(0);
            
            column12Total = new SimpleIntegerProperty(0);
            column12Count = new SimpleIntegerProperty(0);
            column12Avg = new SimpleIntegerProperty(0);
            
            column13Total = new SimpleIntegerProperty(0);
            column13Count = new SimpleIntegerProperty(0);
            column13Avg = new SimpleIntegerProperty(0);
            
            column14Total = new SimpleIntegerProperty(0);
            column14Count = new SimpleIntegerProperty(0);
            column14Avg = new SimpleIntegerProperty(0);
            
            column15Total = new SimpleIntegerProperty(0);
            column15Count = new SimpleIntegerProperty(0);
            column15Avg = new SimpleIntegerProperty(0);
            
            column16Total = new SimpleIntegerProperty(0);
            column16Count = new SimpleIntegerProperty(0);
            column16Avg = new SimpleIntegerProperty(0);
            
            column17Total = new SimpleIntegerProperty(0);
            column17Count = new SimpleIntegerProperty(0);
            column17Avg = new SimpleIntegerProperty(0);
            
            column18Total = new SimpleIntegerProperty(0);
            column18Count = new SimpleIntegerProperty(0);
            column18Avg = new SimpleIntegerProperty(0);
            
            column19Total = new SimpleIntegerProperty(0);
            column19Count = new SimpleIntegerProperty(0);
            column19Avg = new SimpleIntegerProperty(0);
            
            column20Total = new SimpleIntegerProperty(0);
            column20Count = new SimpleIntegerProperty(0);
            column20Avg = new SimpleIntegerProperty(0);
            
            column21Total = new SimpleIntegerProperty(0);
            column21Count = new SimpleIntegerProperty(0);
            column21Avg = new SimpleIntegerProperty(0);
            
            column22Total = new SimpleIntegerProperty(0);
            column22Count = new SimpleIntegerProperty(0);
            column22Avg = new SimpleIntegerProperty(0);
            
            column23Total = new SimpleIntegerProperty(0);
            column23Count = new SimpleIntegerProperty(0);
            column23Avg = new SimpleIntegerProperty(0);
            
            column24Total = new SimpleIntegerProperty(0);
            column24Count = new SimpleIntegerProperty(0);
            column24Avg = new SimpleIntegerProperty(0);
        }
        
        public Integer getColumn4Total(){
            return column4Total.get();
        }
        public void setColumn4Total(Integer column4Total){
            this.column4Total.set(column4Total);
        }
        public Integer getColumn4Count(){
            return column4Count.get();
        }
        public void setColumn4Count(Integer column4Count){
            this.column4Count.set(column4Count);
        }
        public Integer getColumn4Avg(){
            return column4Avg.get();
        }
        public void setColumn4Avg(Integer column4Avg){
            this.column4Avg.set(column4Avg);
        }
        
        public Integer getColumn6Total(){
            return column6Total.get();
        }
        public void setColumn6Total(Integer column6Total){
            this.column6Total.set(column6Total);
        }
        public Integer getColumn6Count(){
            return column6Count.get();
        }
        public void setColumn6Count(Integer column6Count){
            this.column6Count.set(column6Count);
        }
        public Integer getColumn6Avg(){
            return column6Avg.get();
        }
        public void setColumn6Avg(Integer column6Avg){
            this.column6Avg.set(column6Avg);
        }
        
        public Integer getColumn9Total(){
            return column9Total.get();
        }
        public void setColumn9Total(Integer column9Total){
            this.column9Total.set(column9Total);
        }
        public Integer getColumn9Count(){
            return column9Count.get();
        }
        public void setColumn9Count(Integer column9Count){
            this.column9Count.set(column9Count);
        }
        public Integer getColumn9Avg(){
            return column9Avg.get();
        }
        public void setColumn9Avg(Integer column9Avg){
            this.column9Avg.set(column9Avg);
        }
        
        public Integer getColumn10Total(){
            return column10Total.get();
        }
        public void setColumn10Total(Integer column10Total){
            this.column10Total.set(column10Total);
        }
        public Integer getColumn10Count(){
            return column10Count.get();
        }
        public void setColumn10Count(Integer column10Count){
            this.column10Count.set(column10Count);
        }
        public Integer getColumn10Avg(){
            return column10Avg.get();
        }
        public void setColumn10Avg(Integer column10Avg){
            this.column10Avg.set(column10Avg);
        }
        
        public Integer getColumn11Total(){
            return column11Total.get();
        }
        public void setColumn11Total(Integer column11Total){
            this.column11Total.set(column11Total);
        }
        public Integer getColumn11Count(){
            return column11Count.get();
        }
        public void setColumn11Count(Integer column11Count){
            this.column11Count.set(column11Count);
        }
        public Integer getColumn11Avg(){
            return column11Avg.get();
        }
        public void setColumn11Avg(Integer column11Avg){
            this.column11Avg.set(column11Avg);
        }
        
        public Integer getColumn12Total(){
            return column12Total.get();
        }
        public void setColumn12Total(Integer column12Total){
            this.column12Total.set(column12Total);
        }
        public Integer getColumn12Count(){
            return column12Count.get();
        }
        public void setColumn12Count(Integer column12Count){
            this.column12Count.set(column12Count);
        }
        public Integer getColumn12Avg(){
            return column12Avg.get();
        }
        public void setColumn12Avg(Integer column12Avg){
            this.column12Avg.set(column12Avg);
        }
        
        public Integer getColumn13Total(){
            return column13Total.get();
        }
        public void setColumn13Total(Integer column13Total){
            this.column13Total.set(column13Total);
        }
        public Integer getColumn13Count(){
            return column13Count.get();
        }
        public void setColumn13Count(Integer column13Count){
            this.column13Count.set(column13Count);
        }
        public Integer getColumn13Avg(){
            return column13Avg.get();
        }
        public void setColumn13Avg(Integer column13Avg){
            this.column13Avg.set(column13Avg);
        }
        
        public Integer getColumn14Total(){
            return column14Total.get();
        }
        public void setColumn14Total(Integer column14Total){
            this.column14Total.set(column14Total);
        }
        public Integer getColumn14Count(){
            return column14Count.get();
        }
        public void setColumn14Count(Integer column14Count){
            this.column14Count.set(column14Count);
        }
        public Integer getColumn14Avg(){
            return column14Avg.get();
        }
        public void setColumn14Avg(Integer column14Avg){
            this.column14Avg.set(column14Avg);
        }
        
        public Integer getColumn15Total(){
            return column15Total.get();
        }
        public void setColumn15Total(Integer column15Total){
            this.column15Total.set(column15Total);
        }
        public Integer getColumn15Count(){
            return column15Count.get();
        }
        public void setColumn15Count(Integer column15Count){
            this.column15Count.set(column15Count);
        }
        public Integer getColumn15Avg(){
            return column15Avg.get();
        }
        public void setColumn15Avg(Integer column15Avg){
            this.column15Avg.set(column15Avg);
        }
        
        public Integer getColumn18Total(){
            return column18Total.get();
        }
        public void setColumn18Total(Integer column18Total){
            this.column18Total.set(column18Total);
        }
        public Integer getColumn18Count(){
            return column18Count.get();
        }
        public void setColumn18Count(Integer column18Count){
            this.column18Count.set(column18Count);
        }
        public Integer getColumn18Avg(){
            return column18Avg.get();
        }
        public void setColumn18Avg(Integer column18Avg){
            this.column18Avg.set(column18Avg);
        }
        
        public Integer getColumn19Total(){
            return column19Total.get();
        }
        public void setColumn19Total(Integer column19Total){
            this.column19Total.set(column19Total);
        }
        public Integer getColumn19Count(){
            return column19Count.get();
        }
        public void setColumn19Count(Integer column19Count){
            this.column19Count.set(column19Count);
        }
        public Integer getColumn19Avg(){
            return column19Avg.get();
        }
        public void setColumn19Avg(Integer column19Avg){
            this.column19Avg.set(column19Avg);
        }
        
        public Integer getColumn21Total(){
            return column21Total.get();
        }
        public void setColumn21Total(Integer column21Total){
            this.column21Total.set(column21Total);
        }
        public Integer getColumn21Count(){
            return column21Count.get();
        }
        public void setColumn21Count(Integer column21Count){
            this.column21Count.set(column21Count);
        }
        public Integer getColumn21Avg(){
            return column21Avg.get();
        }
        public void setColumn21Avg(Integer column21Avg){
            this.column21Avg.set(column21Avg);
        }
        
        public Integer getColumn22Total(){
            return column22Total.get();
        }
        public void setColumn22Total(Integer column22Total){
            this.column22Total.set(column22Total);
        }
        public Integer getColumn22Count(){
            return column22Count.get();
        }
        public void setColumn22Count(Integer column22Count){
            this.column22Count.set(column22Count);
        }
        public Integer getColumn22Avg(){
            return column22Avg.get();
        }
        public void setColumn22Avg(Integer column22Avg){
            this.column22Avg.set(column22Avg);
        }
        
        public Integer getColumn23Total(){
            return column23Total.get();
        }
        public void setColumn23Total(Integer column23Total){
            this.column23Total.set(column23Total);
        }
        public Integer getColumn23Count(){
            return column23Count.get();
        }
        public void setColumn23Count(Integer column23Count){
            this.column23Count.set(column23Count);
        }
        public Integer getColumn23Avg(){
            return column23Avg.get();
        }
        public void setColumn23Avg(Integer column23Avg){
            this.column23Avg.set(column23Avg);
        }
    }
    
    public static class DataRow {
        private SimpleStringProperty column1;
        private SimpleStringProperty column2;
        private SimpleStringProperty column3;
        private SimpleStringProperty column4;
        private SimpleStringProperty column5;
        private SimpleStringProperty column6;
        private SimpleStringProperty column7;
        private SimpleStringProperty column8;
        private SimpleStringProperty column9;
        private SimpleStringProperty column10;
        private SimpleStringProperty column11;
        private SimpleStringProperty column12;
        private SimpleStringProperty column13;
        private SimpleStringProperty column14;
        private SimpleStringProperty column15;
        private SimpleStringProperty column16;
        private SimpleStringProperty column17;
        private SimpleStringProperty column18;
        private SimpleStringProperty column19;
        private SimpleStringProperty column20;
        private SimpleStringProperty column21;
        private SimpleStringProperty column22;
        private SimpleStringProperty column23;
        private SimpleStringProperty column24;
        
        private DataRow(){
            column1 = new SimpleStringProperty(" ");
            column2 = new SimpleStringProperty(" ");
            column3 = new SimpleStringProperty(" ");
            column4 = new SimpleStringProperty(" ");
            column5 = new SimpleStringProperty(" ");
            column6 = new SimpleStringProperty(" ");
            column7 = new SimpleStringProperty(" ");
            column8 = new SimpleStringProperty(" ");
            column9 = new SimpleStringProperty(" ");
            column10 = new SimpleStringProperty(" ");
            column11 = new SimpleStringProperty(" ");
            column12 = new SimpleStringProperty(" ");
            column13 = new SimpleStringProperty(" ");
            column14 = new SimpleStringProperty(" ");
            column15 = new SimpleStringProperty(" ");
            column16 = new SimpleStringProperty(" ");
            column17 = new SimpleStringProperty(" ");
            column18 = new SimpleStringProperty(" ");
            column19 = new SimpleStringProperty(" ");
            column20 = new SimpleStringProperty(" ");
            column21 = new SimpleStringProperty(" ");
            column22 = new SimpleStringProperty(" ");
            column23 = new SimpleStringProperty(" ");
            column24 = new SimpleStringProperty(" ");
        }
        
        public String getColumn1(){
            return column1.get();
        }
        public void setColumn1(String column1){
            this.column1.set(column1);
        }
        
        public String getColumn2(){
            return column2.get();
        }
        public void setColumn2(String column2){
            this.column2.set(column2);
        }
        
        public String getColumn3(){
            return column3.get();
        }
        public void setColumn3(String column3){
            this.column3.set(column3);
        }
        
        public String getColumn4(){
            return column4.get();
        }
        public void setColumn4(String column4){
            this.column4.set(column4);
        }
        
        public String getColumn5(){
            return column5.get();
        }
        public void setColumn5(String column5){
            this.column5.set(column5);
        }
        
        public String getColumn6(){
            return column6.get();
        }
        public void setColumn6(String column6){
            this.column6.set(column6);
        }
        
        public String getColumn7(){
            return column7.get();
        }
        public void setColumn7(String column7){
            this.column7.set(column7);
        }
        
        public String getColumn8(){
            return column8.get();
        }
        public void setColumn8(String column8){
            this.column8.set(column8);
        }
        
        public String getColumn9(){
            return column9.get();
        }
        public void setColumn9(String column9){
            this.column9.set(column9);
        }
        
        public String getColumn10(){
            return column10.get();
        }
        public void setColumn10(String column10){
            this.column10.set(column10);
        }
        
        public String getColumn11(){
            return column11.get();
        }
        public void setColumn11(String column11){
            this.column11.set(column11);
        }
        
        public String getColumn12(){
            return column12.get();
        }
        public void setColumn12(String column12){
            this.column12.set(column12);
        }
        
        public String getColumn13(){
            return column13.get();
        }
        public void setColumn13(String column13){
            this.column13.set(column13);
        }
        
        public String getColumn14(){
            return column14.get();
        }
        public void setColumn14(String column14){
            this.column14.set(column14);
        }
        
        public String getColumn15(){
            return column15.get();
        }
        public void setColumn15(String column15){
            this.column15.set(column15);
        }
        
        public String getColumn16(){
            return column16.get();
        }
        public void setColumn16(String column16){
            this.column16.set(column16);
        }
        
        public String getColumn17(){
            return column17.get();
        }
        public void setColumn17(String column17){
            this.column17.set(column17);
        }
        
        public String getColumn18(){
            return column18.get();
        }
        public void setColumn18(String column18){
            this.column18.set(column18);
        }
        
        public String getColumn19(){
            return column19.get();
        }
        public void setColumn19(String column19){
            this.column19.set(column19);
        }
        
        public String getColumn20(){
            return column20.get();
        }
        public void setColumn20(String column20){
            this.column20.set(column20);
        }
        
        public String getColumn21(){
            return column21.get();
        }
        public void setColumn21(String column21){
            this.column21.set(column21);
        }
        
        public String getColumn22(){
            return column22.get();
        }
        public void setColumn22(String column22){
            this.column22.set(column22);
        }
        
        public String getColumn23(){
            return column23.get();
        }
        public void setColumn23(String column23){
            this.column23.set(column23);
        }
        
        public String getColumn24(){
            return column24.get();
        }
        public void setColumn24(String column24){
            this.column24.set(column24);
        }
        
    }
}
