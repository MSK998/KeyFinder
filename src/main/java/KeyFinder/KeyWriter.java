package KeyFinder;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

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
        
        Cell cell = row.createCell(columnCount);
        cell.setCellValue(rowCount);

        //for each column
        for (Object field : aKey) {
            cell = row.createCell(++columnCount);
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
