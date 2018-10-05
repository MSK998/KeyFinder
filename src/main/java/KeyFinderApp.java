import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.sql.SQLOutput;
import java.util.*;

/**
 * PACKAGE_NAME was created by Mark Scott-Kiddie on 04/10/2018.
 */
public class KeyFinderApp {
    public static void main(String[] args) {

        KeyFinder sheet1 = new KeyFinder();
        sheet1.loadData();

        System.out.println("Displaying Array Now");

        sheet1.displayArray();

        System.out.println("Outputting cell 2,1");

        sheet1.displaySpecific(2,0);
    }
}
