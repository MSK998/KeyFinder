import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.*;

/**
 * PACKAGE_NAME was created by Mark Scott-Kiddie on 04/10/2018.
 */
public class KeyFinderApp {
    public static void main(String[] args) {

        //This is to test if the method works as there are no objects to test this on yet
        try {
            File myFile = new File("C:\\Users\\mwsco\\IdeaProjects\\KeyFinder\\src\\main\\resources\\Key Records Sample.xlsx");

            FileInputStream fis = new FileInputStream(myFile);

            //Finds the workbook instance for XLSX file
            XSSFWorkbook myWorkBook = new XSSFWorkbook(fis);

            //Return first sheet from the XLSX workbook
            XSSFSheet sheetOne = myWorkBook.getSheetAt(0);

            //Get iterator to move through all rows in the sheet
            Iterator<Row> rowIterator = sheetOne.iterator();

            //Traverse over the row of the XLSX file
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();

                //For each row iterate through the columns
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {

                    Cell cell = cellIterator.next();

                    switch (cell.getCellType()) {
                        case STRING:
                            System.out.println(cell.getStringCellValue() + "\t");
                            break;
                        case NUMERIC:
                            System.out.println(cell.getNumericCellValue() + "\t");
                            break;
                        case BOOLEAN:
                            System.out.println(cell.getBooleanCellValue() + "\t");
                            break;
                        case _NONE:
                            System.out.println(cell.getStringCellValue() + "\t");
                            break;

                        default:
                    }

                }
                System.out.println("");
            }

        } catch (Exception e) {
            System.out.println(e);
            System.exit(1);
        }

    }
}
