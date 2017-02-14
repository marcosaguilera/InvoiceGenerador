import java.io.FileInputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;

import com.mysql.jdbc.Driver;
import java.util.*;
import java.sql.SQLException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import javax.mail.MessagingException;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Row;

public class AccountServices {
    
    public static void main(String[] args) throws SQLException, ParseException, MessagingException {     
        InvoiceActions actions = new InvoiceActions();
        SendMail sendMail = new SendMail();
        
        //1. Truncate table facturacion_sapiens
        actions.TruncateFactTable();
        System.out.println("All data in table facturacion_sapiens deleted.");
        
        //2. Truncate table importacion_cartera
        actions.TruncateCarteraTable();
        System.out.println("All data in table importacion_cartera deleted.");
        
        //3. Delete the existing file PSE_FILE_MAIL.txt
        actions.deleteFileTxt();
        System.out.println("File PSE_FILE_MAIL.txt deleted.");
        
        //4. Locate and import facturacion.xls file into table: facturacion_sapiens
        actions.ImportDataFact();
        System.out.println("All data in facturacion.xls imported.");
        
        //5. Locate and import cartera.xls file into table: importacion_cartera
        actions.ImportDataCartera();  
        System.out.println("All data in cartera.xls imported.");
        
        //6. Update records in mensualidades in table: importacion_cartera
        actions.ReplaceDateMonth();
        System.out.println("All months updated.");
        
        //7. Update records in table: importacion_cartera putting a month number depending of the MONTH NAME.
        actions.ReplaceMonthNumber();
        System.out.println("All months numbers updated.");
        
        //8. PSE File generator
        actions.FilePSEGenerator();
        System.out.println("File created at the default location.");
              
        //9. Sending email with attached file
        sendMail.SendMail();
        System.out.println("Correo enviado!.");
        
        System.out.println("Routine completed successfully!");
    }   
}
