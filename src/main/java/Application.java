
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;


public class Application {
    /**
     * @author : Ahmed Zakaria Ali
     * @param args
     */
    public static void main(String[] args) {
        try {
            FileInputStream myFile2 = new FileInputStream(new File("E:\\Billing\\src\\main\\java\\B2BbillingServices.xlsx"));

            XSSFWorkbook myFile = new XSSFWorkbook(myFile2);
            int number_of_sheets = myFile.getNumberOfSheets();
            System.out.println(number_of_sheets);
            Sheet first_sheet = myFile.getSheetAt(0);
            System.out.println(first_sheet.getLastRowNum());
            Iterator<Row>myIterator=first_sheet.rowIterator();
            System.out.println("Happy Win");
            
            
        } catch (IOException e) {
            System.out.print("we have an Exception " + e.getMessage());
            for (StackTraceElement h : e.getStackTrace()) {

                System.out.println(h.getLineNumber());
            }
        }

    }
}
