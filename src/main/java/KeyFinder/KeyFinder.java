package KeyFinder;

import java.io.File;
import java.io.FileInputStream;
import java.util.*;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;

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

    public void displayArray(){
        for(int r = 0; r < spreadSheet.size(); r++){
            for(int c = 0; c < spreadSheet.get(r).size(); c++){
                System.out.println(spreadSheet.get(r).get(c));
              //  KeyFinderGUI.jTextArea1.append(spreadSheet.get(r).get(c));
                
            }
        }
    }



    public void loadData() {
        long startTime = System.nanoTime();

        /*
         * To Create an option to add an extra field we will need to have a boolean
         * to add an extra field, if wanted we can add an extra variable for the user to decide the field.
         */
        try {
            File myFile = new File(new File("src\\main\\resources\\Key Records Sample.xlsx").getAbsolutePath());

            FileInputStream fis = new FileInputStream(myFile);

            //Finds the workbook instance for XLSX file
            XSSFWorkbook myWorkBook = new XSSFWorkbook(fis);

            //Return first sheet from the XLSX workbook
            XSSFSheet sheetOne = myWorkBook.getSheetAt(1);

            //Get iterator to move through all rows in the sheet
            Iterator<Row> rowIterator = sheetOne.iterator();

            int numCol = (sheetOne.getRow(0).getLastCellNum());
            System.out.println(numCol);

            //Traverse over the row of the XLSX file
            int rowNum = 0;
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();

                spreadSheet.add(new ArrayList<String>());

                //For each row iterate through the columns
                Iterator<Cell> cellIterator = row.cellIterator();
                String fullRow = "";

                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    cell.setCellType(CellType.STRING);

                    if (org.apache.commons.lang3.StringUtils.isBlank(cell.getStringCellValue())) {
                        fullRow += "[BLANK]" + "\t\t\t";
                       
                        spreadSheet.get(rowNum).add("[BLANK]");

                    } else {

                        switch (cell.getCellType()) {
                            case STRING:
                                //System.out.println(cell.getStringCellValue() + "\t");
                                fullRow += cell.getStringCellValue() + "\t\t\t";
                                spreadSheet.get(rowNum).add(cell.getStringCellValue());
                                KeyFinderGUI.jTextArea1.append("\n");
                                break;
                            case NUMERIC:
                                //System.out.println(cell.getNumericCellValue() + "\t");
                                fullRow += cell.getNumericCellValue() + "\t\t\t";
                                spreadSheet.get(rowNum).add(String.valueOf(cell.getNumericCellValue()));
                                KeyFinderGUI.jTextArea1.append("\n");
                                break;
                            case BOOLEAN:
                                //System.out.println(cell.getBooleanCellValue() + "\t");
                                fullRow += cell.getBooleanCellValue() + "\t\t\t";
                                spreadSheet.get(rowNum).add(String.valueOf(cell.getBooleanCellValue()));
                                KeyFinderGUI.jTextArea1.append("\n");
                                break;
                            case FORMULA:
                                break;
//                        case BLANK:
//                            //System.out.println("[BLANK]");
//                            fullRow += cell.getStringCellValue() + "\t\t";
//                            spreadSheet.get(rowNum).add("[BLANK]");
//                            break;
//                        case _NONE:
//                            //System.out.println(cell.getStringCellValue() + "\t");
//                            fullRow += "[BLANK]" + "\t\t";
//                            spreadSheet.get(rowNum).add("[BLANK]");
//                            break;
//
//                        case ERROR:
//                            break;
                            default:
                                fullRow += "[BLANK]" + "\t\t\t";
                                spreadSheet.get(rowNum).add("[BLANK]");
                                KeyFinderGUI.jTextArea1.append("\n");
                        }
                        /*
                         * This is where we should put the "add extra field code"
                         */
                    }
                }

                System.out.println(fullRow);
                KeyFinderGUI.jTextArea1.append(fullRow);
                rowNum++;
            }

        } catch (Exception e) {
            System.out.println("error" + e);
            System.exit(1);
        }

        System.out.println((System.nanoTime() - startTime) / 1000000);
    }
}
