package KeyFinder;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import  org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

/**
 *
 * @author Lewis
 */
public class KeyWriter {
    
    
    public void updateSpecificRow() throws IOException {
        Scanner user_input = new Scanner(System.in);
        Scanner rowNumber = new Scanner(System.in);
        Scanner ColumnNumber = new Scanner(System.in);
            try {
        //Get the excel file.
        FileInputStream file = new FileInputStream(new File("/Users/Lewis/Desktop/University/CM3108/KeyRecordsSample.xlsx"));
 
        //Get workbook for XLS file.
        XSSFWorkbook yourworkbook = new XSSFWorkbook(file);
 
        //Get first sheet from the workbook.
        //If there have >1 sheet in your workbook, you can change it here IF you want to edit other sheets.
        XSSFSheet sheet1 = yourworkbook.getSheetAt(1);
 
        // Get the row of your desired cell.
        // desired cell is at row 3.
        System.out.println("Enter the row number of which you want to change");
        Row row = sheet1.getRow(rowNumber.nextInt());
       //Row row = sheet1.getRow(2);
        // Get the column of your desired cell in your selected row.
        // desired cell is at column 3.
        System.out.println("Enter the column number of which you want to change");
        
       Cell column = row.getCell(2);
        // If the cell is String type
        String updatename = column.getStringCellValue();
        //New content for desired cell
        System.out.println("Make the changes");
        updatename=user_input.next();
        //Print out the updated content.
        System.out.println(updatename);
        //Set the new content to your desired cell(column).
        column.setCellValue(updatename); 
        //Close the excel file.
        file.close();
        //Where you want to save the updated sheet.
        FileOutputStream out = 
            new FileOutputStream(new File("/Users/Lewis/Desktop/University/CM3108/KeyRecordsSample.xlsx"));
        yourworkbook.write(out);
        out.close();
 
    } catch (FileNotFoundException e) {
        e.printStackTrace();
    }
    }

    
  public static void main(String[] args) {
      //states the file location, will change to a relative path in the project.
    String excelFilePath = ("/Users/Lewis/Desktop/University/CM3108/KeyRecordsSample.xlsx");

    try {
        //finds the file and sets as input stream
    FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
    //sets workbook and sheet
    Workbook workbook = WorkbookFactory.create(inputStream);
    
    Sheet sheet = workbook.getSheetAt(1);

    //test data
    Object[][] newKeys = {
        {2456, "Tambour unit", "Lewis", "N533"},
        {2245, "Tambour unit", "Jim", "N524"},
        {1234, "Lab", "Mark", "N117"},
        {2459, "Door", "Ryan", "N128"},
            
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
    
    
    FileOutputStream outputStream = new FileOutputStream("/Users/Lewis/Desktop/University/CM3108/KeyRecordsSample.xlsx");
			workbook.write(outputStream);
			workbook.close();
			outputStream.close();
			
		} catch (IOException |
                        EncryptedDocumentException  ex) {
			ex.printStackTrace();
		}
	}
  
  

}
