package KeyFinder;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.commons.compress.archivers.dump.InvalidFormatException;
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
    String excelFilePath = ("/Users/Lewis/Desktop/University/CM3108/KeyRecordsSample.xlsx");
    
    try {
    FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
    Workbook workbook = WorkbookFactory.create(inputStream);
    
    Sheet sheet = workbook.getSheetAt(0);
    
    Object[][] newKeys = {
        {2456, "Tambour unit", "Lewis", "N533"},
        {2245, "Tambour unit", "Jim", "N524"},
        {1234, "Lab", "Mark", "N117"},
        {2459, "Door", "Ryan", "N128"},
            
    };
    
    int rowCount = sheet.getLastRowNum();
    
    for (Object[] aKey : newKeys) {
        Row row = sheet.createRow(++rowCount);
        
        int columnCount = 0;
        
        Cell cell = row.createCell(columnCount);
        cell.setCellValue(rowCount);
        
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
