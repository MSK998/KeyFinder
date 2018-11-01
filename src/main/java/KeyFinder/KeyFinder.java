package KeyFinder;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * KeyFinder.KeyFinder Database Handler Module was created by Mark Scott-Kiddie and Lewis Ross on 04/10/2018.
 *
 * The Database Handler Module's purpose is to import xlsx file formats into the system as it will
 * have to work with the original users system if there are any problems with the new system.
 */
public class KeyFinder {

    private List<List<String>> spreadSheet;

    public KeyFinder(){
        spreadSheet = new ArrayList<List<String>>();
    }

    public void displaySpecific(int y, int x){
        System.out.println(spreadSheet.get(y).get(x));
    }

    public String displayArray(){

        for(int r = 0; r < spreadSheet.size(); r++){

            String fullRows = "";

            for(int c = 0; c < spreadSheet.get(r).size(); c++){

                fullRows += (spreadSheet.get(r).get(c));
                KeyFinderGUI.outputTextArea.append(spreadSheet.get(r).get(c));
            }
            return fullRows;
        }
        return "Finished";
    }

    public void addColumn(String fieldName, String defaultString, boolean defaultBoolean){
        spreadSheet.get(0).add(fieldName);

        /*
        This section of the method will be dictated by the checkboxes to decide wht kind of field it is.
        if(
        >Certain checkbox is ticked then do string type insert
        >else if boolean checkbox is ticked then do boolean insert
        >else (if nothing is checked) do nothing
        )
        */
        for(int x = 1; x < spreadSheet.size(); x++){
            spreadSheet.get(x).add(defaultString);
        }
    }

    public void loadKeyData() {
        long startTime = System.nanoTime();

        /*
         * To Create an option to add an extra field we will need to have a boolean
         * to add an extra field, if wanted we can add an extra variable for the user to decide the field.
         */
        try {
            File myFile = new File(new File("src/main/resources/Key Records Sample.xlsx").getAbsolutePath());

            FileInputStream fis = new FileInputStream(myFile);

            //Finds the workbook instance for XLSX file
            XSSFWorkbook myWorkBook = new XSSFWorkbook(fis);

            //Return first sheet from the XLSX workbook
            XSSFSheet sheet = myWorkBook.getSheetAt(1);

            //Get manually iterate through all rows in the sheet
            XSSFRow row;

            int totalNumCol = (sheet.getRow(0).getLastCellNum());
            System.out.println(totalNumCol);
            

            //Traverse over the row of the XLSX file
            int finalRowNum = sheet.getLastRowNum();

            for(int x = 0; x < finalRowNum; x++){
                row = sheet.getRow(x);
                spreadSheet.add(new ArrayList<String>());
                String fullRow = "";
                


                for(int y = 0; y < totalNumCol; y++) {
                    XSSFCell cell = row.getCell(y, MissingCellPolicy.RETURN_BLANK_AS_NULL);

                    if(cell == null){
                        spreadSheet.get(x).add("[BLANK]");
                        fullRow += "[BLANK]" + "\t\t\t";
                        
                    } else {
                        switch (cell.getCellType()) {
                            case STRING:
                                //System.out.println(cell.getStringCellValue() + "\t");
                                fullRow += cell.getStringCellValue() + "\t\t\t";
                                spreadSheet.get(x).add(cell.getStringCellValue());
                                KeyFinderGUI.outputTextArea.append("\n");
                                
                                
                                break;
                            case NUMERIC:
                                //System.out.println(cell.getNumericCellValue() + "\t");
                                fullRow += cell.getNumericCellValue() + "\t\t\t";
                                spreadSheet.get(x).add(String.valueOf(cell.getNumericCellValue()));
                               
                                
                                break;
                            case BOOLEAN:
                                //System.out.println(cell.getBooleanCellValue() + "\t");
                                fullRow += cell.getBooleanCellValue() + "\t\t\t";
                                spreadSheet.get(x).add(String.valueOf(cell.getBooleanCellValue()));
                                
                               
                                break;
                            case FORMULA:
                                break;
                            default:
                                fullRow += "[BLANK]" + "\t\t\t";
                                spreadSheet.get(x).add("[BLANK]");
                               
                        }
                    }
                }
                System.out.println(fullRow);
                KeyFinderGUI.outputTextArea.append(fullRow);
            }

        } catch (Exception e) {
            System.out.println("error" + e);
            System.exit(1);
        }

        System.out.println((System.nanoTime() - startTime) / 1000000);
    }
    
    public void loadFobs(){
         long startTime = System.nanoTime();

        /*
         * To Create an option to add an extra field we will need to have a boolean
         * to add an extra field, if wanted we can add an extra variable for the user to decide the field.
         */
        try {
            File myFile = new File(new File("src/main/resources/Key Records Sample.xlsx").getAbsolutePath());

            FileInputStream fis = new FileInputStream(myFile);

            //Finds the workbook instance for XLSX file
            XSSFWorkbook myWorkBook = new XSSFWorkbook(fis);

            //Return first sheet from the XLSX workbook
            XSSFSheet sheet = myWorkBook.getSheetAt(0);

            //Get manually iterate through all rows in the sheet
            XSSFRow row;

            int totalNumCol = (sheet.getRow(0).getLastCellNum());
            System.out.println(totalNumCol);
            

            //Traverse over the row of the XLSX file
            int finalRowNum = sheet.getLastRowNum();

            for(int x = 0; x < finalRowNum; x++){
                row = sheet.getRow(x);
                spreadSheet.add(new ArrayList<String>());
                String fullRow = "";
                


                for(int y = 0; y < totalNumCol; y++) {
                    XSSFCell cell = row.getCell(y, MissingCellPolicy.RETURN_BLANK_AS_NULL);

                    if(cell == null){
                        spreadSheet.get(x).add("[BLANK]");
                        fullRow += "[BLANK]" + "\t\t\t";
                        
                    } else {
                        switch (cell.getCellType()) {
                            case STRING:
                                //System.out.println(cell.getStringCellValue() + "\t");
                                fullRow += cell.getStringCellValue() + "\t\t\t";
                                spreadSheet.get(x).add(cell.getStringCellValue());
                                KeyFinderGUI.outputTextArea.append("\n");
                                
                                
                                break;
                            case NUMERIC:
                                //System.out.println(cell.getNumericCellValue() + "\t");
                                fullRow += cell.getNumericCellValue() + "\t\t\t";
                                spreadSheet.get(x).add(String.valueOf(cell.getNumericCellValue()));
                               
                                
                                break;
                            case BOOLEAN:
                                //System.out.println(cell.getBooleanCellValue() + "\t");
                                fullRow += cell.getBooleanCellValue() + "\t\t\t";
                                spreadSheet.get(x).add(String.valueOf(cell.getBooleanCellValue()));
                                
                               
                                break;
                            case FORMULA:
                                break;
                            default:
                                fullRow += "[BLANK]" + "\t\t\t";
                                spreadSheet.get(x).add("[BLANK]");
                               
                        }
                    }
                }
                System.out.println(fullRow);
                KeyFinderGUI.outputTextArea.append(fullRow);
            }

        } catch (Exception e) {
            System.out.println("error" + e);
            System.exit(1);
        }

        System.out.println((System.nanoTime() - startTime) / 1000000);
        
    }
    
    
    public void loadLost() {
         long startTime = System.nanoTime();

        /*
         * To Create an option to add an extra field we will need to have a boolean
         * to add an extra field, if wanted we can add an extra variable for the user to decide the field.
         */
        try {
            File myFile = new File(new File("src/main/resources/Key Records Sample.xlsx").getAbsolutePath());

            FileInputStream fis = new FileInputStream(myFile);

            //Finds the workbook instance for XLSX file
            XSSFWorkbook myWorkBook = new XSSFWorkbook(fis);

            //Return first sheet from the XLSX workbook
            XSSFSheet sheet = myWorkBook.getSheetAt(2);

            //Get manually iterate through all rows in the sheet
            XSSFRow row;

            int totalNumCol = (sheet.getRow(0).getLastCellNum());
            System.out.println(totalNumCol);
            

            //Traverse over the row of the XLSX file
            int finalRowNum = sheet.getLastRowNum();

            for(int x = 0; x < finalRowNum; x++){
                row = sheet.getRow(x);
                spreadSheet.add(new ArrayList<String>());
                String fullRow = "";
                


                for(int y = 0; y < totalNumCol; y++) {
                    XSSFCell cell = row.getCell(y, MissingCellPolicy.RETURN_BLANK_AS_NULL);

                    if(cell == null){
                        spreadSheet.get(x).add("[BLANK]");
                        fullRow += "[BLANK]" + "\t\t\t";
                        
                    } else {
                        switch (cell.getCellType()) {
                            case STRING:
                                //System.out.println(cell.getStringCellValue() + "\t");
                                fullRow += cell.getStringCellValue() + "\t\t\t";
                                spreadSheet.get(x).add(cell.getStringCellValue());
                                KeyFinderGUI.outputTextArea.append("\n");
                                
                                
                                break;
                            case NUMERIC:
                                //System.out.println(cell.getNumericCellValue() + "\t");
                                fullRow += cell.getNumericCellValue() + "\t\t\t";
                                spreadSheet.get(x).add(String.valueOf(cell.getNumericCellValue()));
                               
                                
                                break;
                            case BOOLEAN:
                                //System.out.println(cell.getBooleanCellValue() + "\t");
                                fullRow += cell.getBooleanCellValue() + "\t\t\t";
                                spreadSheet.get(x).add(String.valueOf(cell.getBooleanCellValue()));
                                
                               
                                break;
                            case FORMULA:
                                break;
                            default:
                                fullRow += "[BLANK]" + "\t\t\t";
                                spreadSheet.get(x).add("[BLANK]");
                               
                        }
                    }
                }
                System.out.println(fullRow);
                KeyFinderGUI.outputTextArea.append(fullRow);
            }

        } catch (Exception e) {
            System.out.println("error" + e);
            System.exit(1);
        }

        System.out.println((System.nanoTime() - startTime) / 1000000);
        
        
    }
    
  public void keywrite(){
        KeyFinderGUI key = new KeyFinderGUI();
          String excelFilePath = "src/main/resources/Key Records Sample.xlsx";

    try {
        //finds the file and sets as input stream
    FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
    //sets workbook and sheet
    Workbook workbook = WorkbookFactory.create(inputStream);
    
    Sheet sheet = workbook.getSheetAt(1);

    //test data
    Object[][] newKeys = {
        {key.jTextField1.getText(), key.jTextField2.getText(), key.jTextField3.getText(), key.jTextField4.getText() },
        
           
    };

    //finds the amount of rows in the sheet
    int rowCount = sheet.getLastRowNum();

    //for the object row
    for (Object[] aKey : newKeys) {
        Row row = sheet.createRow(++rowCount);
            
        
        //column count set to 0
       int columnCount = 0;

        for (Object field : aKey) {
            Cell cell = row.createCell(columnCount++);
            if (field instanceof String) {
                cell.setCellValue((String) field);
            } else if (field instanceof Integer) {
                cell.setCellValue((Integer) field);
                
            }
        }
                
    }
    
    inputStream.close();
    
    
    FileOutputStream outputStream = new FileOutputStream("src/main/resources/Key Records Sample.xlsx");
			workbook.write(outputStream);
			workbook.close();
			outputStream.close();
                       
			
		} catch (IOException |
                        EncryptedDocumentException  ex) {
			ex.printStackTrace();
		}
	}
 
    
public void fobWrite(){
    
           KeyFinderGUI fob = new KeyFinderGUI();
          String excelFilePath = "src/main/resources/Key Records Sample.xlsx";

    try {
        //finds the file and sets as input stream
    FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
    //sets workbook and sheet
    Workbook workbook = WorkbookFactory.create(inputStream);
    
    Sheet sheet = workbook.getSheetAt(0);

    //test data
    Object[][] newKeys = {
        {fob.jTextField1.getText(), fob.jTextField4.getText() },
        
           
    };

    //finds the amount of rows in the sheet
    int rowCount = sheet.getLastRowNum();

    //for the object row
    for (Object[] aKey : newKeys) {
        Row row = sheet.createRow(++rowCount);
            
        
        //column count set to 0
       int columnCount = 0;

        for (Object field : aKey) {
            Cell cell = row.createCell(columnCount++);
            if (field instanceof String) {
                cell.setCellValue((String) field);
            } else if (field instanceof Integer) {
                cell.setCellValue((Integer) field);
                
            }
        }
                
    }
    
    inputStream.close();
    
    
    FileOutputStream outputStream = new FileOutputStream("src/main/resources/Key Records Sample.xlsx");
			workbook.write(outputStream);
			workbook.close();
			outputStream.close();
                       
			
		} catch (IOException |
                        EncryptedDocumentException  ex) {
			ex.printStackTrace();
		}

}


public void lostKeyWrite(){
    
           KeyFinderGUI lost = new KeyFinderGUI();
          String excelFilePath = "src/main/resources/Key Records Sample.xlsx";

    try {
        //finds the file and sets as input stream
    FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
    //sets workbook and sheet
    Workbook workbook = WorkbookFactory.create(inputStream);
    
    Sheet sheet = workbook.getSheetAt(2);

    //test data
    Object[][] newKeys = {
        {lost.jTextField1.getText(), lost.jTextField2.getText(), lost.jTextField4.getText() },
        
           
    };

    //finds the amount of rows in the sheet
    int rowCount = sheet.getLastRowNum();

    //for the object row
    for (Object[] aKey : newKeys) {
        Row row = sheet.createRow(++rowCount);
            
        
        //column count set to 0
       int columnCount = 0;

        for (Object field : aKey) {
            Cell cell = row.createCell(columnCount++);
            if (field instanceof String) {
                cell.setCellValue((String) field);
            } else if (field instanceof Integer) {
                cell.setCellValue((Integer) field);
                
            }
        }
                
    }
    
    inputStream.close();
    
    
    FileOutputStream outputStream = new FileOutputStream("src/main/resources/Key Records Sample.xlsx");
			workbook.write(outputStream);
			workbook.close();
			outputStream.close();
                       
			
		} catch (IOException |
                        EncryptedDocumentException  ex) {
			ex.printStackTrace();
		}
    
}

}



