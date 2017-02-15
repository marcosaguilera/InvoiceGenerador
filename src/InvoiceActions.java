import java.io.FileInputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;

import com.mysql.jdbc.Driver;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.sql.ResultSet;
//import com.mysql.jdbc.Statement;
//import java.sql.Date;
import java.util.*;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.Format;
import java.text.ParseException;
import java.text.SimpleDateFormat;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Row;

public class InvoiceActions {
    
    public void ImportDataFact() throws SQLException, ParseException {
        
        try {
                Connection con = (Connection) DriverManager.getConnection("jdbc:mysql://10.0.3.8:3306/aprendoz_desarrollo", "root", "irc4Quag");
            con.setAutoCommit(false);
            PreparedStatement pstm = null;
            FileInputStream input = new FileInputStream("/Volumes/sapiens/INTERFASE/facturacion.xls");
            POIFSFileSystem fs = new POIFSFileSystem(input);
            HSSFWorkbook wb = new HSSFWorkbook(fs);

            HSSFSheet sheet = wb.getSheetAt(0);
            
            Row row;
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                
                row = sheet.getRow(i);
                String dono     = row.getCell(0).getStringCellValue();
                String ticon    = row.getCell(1).getStringCellValue();
                Date fecha      = row.getCell(2).getDateCellValue();      // Esta variable Date es de tipo java.util.Date;
                String concepto = row.getCell(3).getStringCellValue();
                String nconcepto= row.getCell(4).getStringCellValue();
                Double valor    = row.getCell(5).getNumericCellValue();
                Double pordcto  = row.getCell(6).getNumericCellValue();
                Double valdcto  = row.getCell(7).getNumericCellValue();
                Double anticipo = row.getCell(8).getNumericCellValue();
                Double saldo    = row.getCell(9).getNumericCellValue();
                String nnombre  = row.getCell(10).getStringCellValue();
                String cursoact = row.getCell(11).getStringCellValue();
                Boolean alu     = row.getCell(12).getBooleanCellValue();
                Boolean otro    = row.getCell(13).getBooleanCellValue();
                
                /*Parsing some variables*/
                Date today = new Date();
                
                java.sql.Date date_sql = new java.sql.Date(fecha.getTime()); //Esta variable Date es de tipo java.sql.Date
                java.sql.Date today_sql = new java.sql.Date(today.getTime());
                
                String valor_2   = String.valueOf(valor);
                String pordcto_2 = String.valueOf(pordcto);
                String valdcto_2 = String.valueOf(valdcto);
                String anticipo_2= String.valueOf(anticipo);
                String saldo_2   = String.valueOf(saldo);
                String alu_2     = String.valueOf(alu);
                String otro_2    = String.valueOf(otro);
                
                String sql = "INSERT INTO facturacion_sapiens (dono, ticon, fecha, concepto, nconcepto, valor, pordcto, valdcto, anticipo, saldo, nnombre, cursoact, alu, otro, created_at) VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";
                pstm = (PreparedStatement) con.prepareStatement(sql);

                pstm.setString(1, dono);
                pstm.setString(2, ticon);
                pstm.setDate(3, date_sql); // aquí se envia la fecha en formato Date SQL
                pstm.setString(4, concepto);
                pstm.setString(5, nconcepto);
                pstm.setString(6, valor_2);
                pstm.setString(7, pordcto_2);
                pstm.setString(8, valdcto_2);
                pstm.setString(9, anticipo_2);
                pstm.setString(10, saldo_2);
                pstm.setString(11, nnombre);
                pstm.setString(12, cursoact);
                pstm.setString(13, alu_2);
                pstm.setString(14, otro_2);
                pstm.setDate(15, today_sql);
                
                pstm.execute();
                System.out.println("Facturacion -> Import rows " + i + " with cell date: "+ fecha);
            }
            con.commit();
            pstm.close();
            con.close();
            input.close();
            System.out.println("Success import excel to mysql table");
        } catch (IOException e) {
            e.printStackTrace();
        }     
    }    
    
    public void ImportDataCartera() throws SQLException{
        try {
            Connection con = (Connection) DriverManager.getConnection("jdbc:mysql://10.0.3.8:3306/aprendoz_desarrollo", "root", "irc4Quag");
            con.setAutoCommit(false);
            PreparedStatement pstm = null;
            FileInputStream input = new FileInputStream("/Volumes/sapiens/INTERFASE/cartera.xls");
            POIFSFileSystem fs = new POIFSFileSystem(input);
            HSSFWorkbook wb = new HSSFWorkbook(fs);

            HSSFSheet sheet = wb.getSheetAt(0);
            
            Row row;
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                
                row = sheet.getRow(i);
                String ncodigo  = row.getCell(1).getStringCellValue();
                Date fecha_a    = row.getCell(2).getDateCellValue();            // Esta variable Date es de tipo java.util.Date;
                String nconcepto= row.getCell(16).getStringCellValue();      
                Double saldo    = row.getCell(14).getNumericCellValue();
                String nnombre  = row.getCell(64).getStringCellValue();
                
                /*Parsing some variables*/
                Date today = new Date();
                
                java.sql.Date date_sql = new java.sql.Date(fecha_a.getTime());    //Esta variable Date es de tipo java.sql.Date
                java.sql.Date today_sql = new java.sql.Date(today.getTime());
                
                String sql = "INSERT INTO importacion_cartera (codigo, mensualidad, concepto, valor, alumno) VALUES(?, ?, ?, ?, ?)";
                pstm = (PreparedStatement) con.prepareStatement(sql);

                pstm.setString(1, ncodigo);
                pstm.setDate(2, date_sql);   // aquí se envia la fecha en formato Date SQL
                pstm.setString(3, nconcepto); 
                pstm.setDouble(4, saldo);
                pstm.setString(5, nnombre);
                
                pstm.execute();
                System.out.println("Cartera -> Import rows " + i + " with cell date: "+ fecha_a);
            }
            con.commit();
            pstm.close();
            con.close();
            input.close();
            System.out.println("Success import excel to mysql table");
        } catch (IOException e) {
            e.printStackTrace();
        }        
    }    

    public void TruncateFactTable() throws SQLException{
        Connection con = (Connection) DriverManager.getConnection("jdbc:mysql://10.0.3.8:3306/aprendoz_desarrollo", "root", "irc4Quag");
        
        try{
            Statement stmt = null;
            stmt = con.createStatement();
            String sql = "DELETE FROM facturacion_sapiens";
            stmt.executeUpdate(sql);
            System.out.println("Table  deleted in given database...");
        }catch(SQLException se){
            //Handle errors for JDBC
            se.printStackTrace();
         }catch(Exception e){
            //Handle errors for Class.forName
            e.printStackTrace();
         }finally{
            con.close();
        }   
        System.out.println("Success: TRUNCATE table ***facturacion_sapiens***"); 
    }
    
    public void TruncateCarteraTable() throws SQLException{
        Connection con = (Connection) DriverManager.getConnection("jdbc:mysql://10.0.3.8:3306/aprendoz_desarrollo", "root", "irc4Quag");
        
        try{
            Statement stmt = null;
            stmt = con.createStatement();
            String sql = "DELETE FROM importacion_cartera";
            stmt.executeUpdate(sql);
            System.out.println("Table  deleted in given database...");
        }catch(SQLException se){
            //Handle errors for JDBC
            se.printStackTrace();
         }catch(Exception e){
            //Handle errors for Class.forName
            e.printStackTrace();
         }finally{
            con.close();
        }   
        System.out.println("Success: TRUNCATE table ***importacion_cartera***"); 
    }  
    
    public void ReplaceDateMonth() throws SQLException{
        Connection con = (Connection) DriverManager.getConnection("jdbc:mysql://10.0.3.8:3306/aprendoz_desarrollo", "root", "irc4Quag");
        
        try{
            Statement stmt = null;
            stmt = con.createStatement();
                String sql_update_aug  = "UPDATE importacion_cartera i set i.mensualidad = \"AGOSTO\" WHERE i.mensualidad=\"2016-08-01\"";
                String sql_update_sep  = "UPDATE importacion_cartera i set i.mensualidad = \"SEPTIEMBRE\" WHERE i.mensualidad=\"2016-09-01\"";
                String sql_update_oct  = "UPDATE importacion_cartera i set i.mensualidad = \"OCTUBRE\" WHERE i.mensualidad=\"2016-10-01\"";
                String sql_update_nov  = "UPDATE importacion_cartera i set i.mensualidad = \"NOVIEMBRE\" WHERE i.mensualidad=\"2016-11-01\"";
                String sql_update_dec  = "UPDATE importacion_cartera i set i.mensualidad = \"DICIEMBRE\" WHERE i.mensualidad=\"2016-12-01\"";
                String sql_update_jan  = "UPDATE importacion_cartera i set i.mensualidad = \"ENERO\" WHERE i.mensualidad=\"2017-01-01\"";
                String sql_update_feb  = "UPDATE importacion_cartera i set i.mensualidad = \"FEBRERO\" WHERE i.mensualidad=\"2017-02-01\"";
                String sql_update_mar  = "UPDATE importacion_cartera i set i.mensualidad = \"MARZO\" WHERE i.mensualidad=\"2017-03-01\"";
                String sql_update_apr  = "UPDATE importacion_cartera i set i.mensualidad = \"ABRIL\" WHERE i.mensualidad=\"2017-04-01\"";
                String sql_update_may  = "UPDATE importacion_cartera i set i.mensualidad = \"MAYO\" WHERE i.mensualidad=\"2017-05-01\"";
                String sql_update_jun  = "UPDATE importacion_cartera i set i.mensualidad = \"JUNIO\" WHERE i.mensualidad=\"2017-06-01\"";
               
                stmt.executeUpdate(sql_update_aug);
                stmt.executeUpdate(sql_update_sep);
                stmt.executeUpdate(sql_update_oct);
                stmt.executeUpdate(sql_update_nov);
                stmt.executeUpdate(sql_update_dec);
                stmt.executeUpdate(sql_update_jan);
                stmt.executeUpdate(sql_update_feb);
                stmt.executeUpdate(sql_update_mar);
                stmt.executeUpdate(sql_update_apr);
                stmt.executeUpdate(sql_update_may);
                stmt.executeUpdate(sql_update_jun);
                
            System.out.println("Table updated in given database...");
            
        }catch(SQLException se){
            //Handle errors for JDBC
            se.printStackTrace();
         }catch(Exception e){
            //Handle errors for Class.forName
            e.printStackTrace();
         }finally{
            con.close();
        }   
        System.out.println("Success: UPDATE mensualidad in table ***importacion_cartera***"); 
    }
    
    public void ReplaceMonthNumber() throws SQLException{
        Connection con = (Connection) DriverManager.getConnection("jdbc:mysql://10.0.3.8:3306/aprendoz_desarrollo", "root", "irc4Quag");
        
        try{
            Statement stmt = null;
            stmt = con.createStatement();
                String sql_update_jul_num  = "UPDATE importacion_cartera i set i.numero_mes = 0 where i.mensualidad = \"JULIO\"";
                String sql_update_aug_num  = "UPDATE importacion_cartera i set i.numero_mes = 1 where i.mensualidad = \"AGOSTO\"";
                String sql_update_sep_num  = "UPDATE importacion_cartera i set i.numero_mes = 2 where i.mensualidad = \"SEPTIEMBRE\"";
                String sql_update_oct_num  = "UPDATE importacion_cartera i set i.numero_mes = 3 where i.mensualidad = \"OCTUBRE\"";
                String sql_update_nov_num  = "UPDATE importacion_cartera i set i.numero_mes = 4 where i.mensualidad = \"NOVIEMBRE\"";
                String sql_update_dec_num  = "UPDATE importacion_cartera i set i.numero_mes = 5 where i.mensualidad = \"DICIEMBRE\"";
                String sql_update_jan_num  = "UPDATE importacion_cartera i set i.numero_mes = 6 where i.mensualidad = \"ENERO\"";
                String sql_update_feb_num  = "UPDATE importacion_cartera i set i.numero_mes = 7 where i.mensualidad = \"FEBRERO\"";
                String sql_update_mar_num  = "UPDATE importacion_cartera i set i.numero_mes = 8 where i.mensualidad = \"MARZO\"";
                String sql_update_apr_num  = "UPDATE importacion_cartera i set i.numero_mes = 9 where i.mensualidad = \"ABRIL\"";
                String sql_update_may_num  = "UPDATE importacion_cartera i set i.numero_mes = 10 where i.mensualidad = \"MAYO\"";
                String sql_update_jun_num  = "UPDATE importacion_cartera i set i.numero_mes = 11 where i.mensualidad = \"JUNIO\"";
               
                stmt.executeUpdate(sql_update_jul_num);
                stmt.executeUpdate(sql_update_aug_num);
                stmt.executeUpdate(sql_update_sep_num);
                stmt.executeUpdate(sql_update_oct_num);
                stmt.executeUpdate(sql_update_nov_num);
                stmt.executeUpdate(sql_update_dec_num);
                stmt.executeUpdate(sql_update_jan_num);
                stmt.executeUpdate(sql_update_feb_num);
                stmt.executeUpdate(sql_update_mar_num);
                stmt.executeUpdate(sql_update_apr_num);
                stmt.executeUpdate(sql_update_may_num);
                stmt.executeUpdate(sql_update_jun_num);
                
            System.out.println("Table updated in given database...");
            
        }catch(SQLException se){
            //Handle errors for JDBC
            se.printStackTrace();
         }catch(Exception e){
            //Handle errors for Class.forName
            e.printStackTrace();
         }finally{
            con.close();
        }   
        System.out.println("Success: UPDATE mensualidad in table ***importacion_cartera***"); 
    }
    
    public void FilePSEGenerator() throws SQLException{
        Connection con = (Connection) DriverManager.getConnection("jdbc:mysql://10.0.3.8:3306/aprendoz_desarrollo", "root", "irc4Quag");
        List data = new ArrayList();
        List logdata = new ArrayList();
        
        try{
            Statement stmt = null;
            stmt = con.createStatement();
            ResultSet rs = stmt.executeQuery("SELECT \n" +
                                        "	CONCAT(MONTH(CURDATE()),'17',meses.mesnumero,importacion_cartera.codigo,'|') as nodoc,\n" +
                                        "	PERSONA.Codigo as cod,\n" +
                                        "	'|' as pipe1,\n" +
                                        "	/*modificacion temporal*/\n" +
                                        "	/*(cartera_codigo_mes(importacion_cartera.codigo,importacion_cartera.mensualidad)) as total_deuda_temporal,*/\n" +
                                        "	\n" +
                                        "	ROUND(sum((cartera_codigo_mes(importacion_cartera.codigo,importacion_cartera.mensualidad) \n" +
                                        "	* ( 1 + calcula_valor_recargo_mensualidades(importacion_cartera.numero_mes))) / numerodeconceptosmensuales(importacion_cartera.codigo,importacion_cartera.numero_mes))) AS totaldeuda,\n" +
                                        "	'|0|' as pipe2,\n" +
                                        "	CONCAT(meses.mes,'-',importacion_cartera.concepto,'|') as concat1,\n" +
                                        "	DATE_FORMAT(CURDATE(),'%d/%m/%Y') as date1,\n" +
                                        "	'|||' as pipe3,\n" +
                                        "	CASE persona.Nombre1 WHEN NULL THEN ' ' ELSE persona.Nombre1 end as primernombre,\n" +
                                        "	CASE persona.Nombre2 WHEN NULL THEN ' ' ELSE persona.Nombre2 end as segundonombre,\n" +
                                        "	'|' as pipe4,\n" +
                                        "	persona.Apellido1 as apellido1,\n" +
                                        "	' ' as space1,\n" +
                                        "	persona.Apellido2 as apellido2,\n" +
                                        "	'|' as pipe5,\n" +
                                        "	persona.Telefono as telefono,\n" +
                                        "	'||||' as pipe6,\n" +
                                        "	DATE_FORMAT(CURDATE(),'%d/%m/%Y') as date2,\n" +
                                        "	'|' as pipe7,\n" +
                                        "	/*(cartera_codigo_mes(importacion_cartera.codigo,importacion_cartera.mensualidad)) as total_deuda_temporal2,*/\n" +
                                        "	ROUND(sum(cartera_codigo_mes(importacion_cartera.codigo,importacion_cartera.mensualidad) \n" +
                                        "	* ( 1 + calcula_valor_recargo_mensualidades(importacion_cartera.numero_mes))) / numerodeconceptosmensuales(importacion_cartera.codigo,importacion_cartera.numero_mes)) AS totaldeuda2,\n" +
                                        "	'|0|PAGOMENSUALIDADES'\n" +
                                        "FROM importacion_cartera inner join PERSONA on persona.Codigo = importacion_cartera.codigo\n" +
                                        "	inner join meses on importacion_cartera.mensualidad = meses.mes\n" +
                                        "WHERE persona.Tipo_Persona_Id_Tipo_Persona = 1 /*AND PERSONA.Codigo=12150*/\n" +
                                        "group by importacion_cartera.codigo, meses.mesnumero\n" +
                                        "order by importacion_cartera.codigo, importacion_cartera.numero_mes");
                
            while (rs.next()) {
                            
                            String logdate = dateFileName();
                            
                            int rows            = rs.getRow();
                            String nodoc        = rs.getString("nodoc");
                            String cod          = rs.getString("cod");
                            String pipe1        = rs.getString("pipe1");
                            String totaldeuda   = rs.getString("totaldeuda");
                            String pipe2        = rs.getString("pipe2");
                            String concat1      = rs.getString("concat1");
                            String date1        = rs.getString("date1");
                            String pipe3        = rs.getString("pipe3");
                            String primernombre = rs.getString("primernombre");
                            String segundonombre= rs.getString("segundonombre");
                            String pipe4        = rs.getString("pipe4");
                            String apellido1    = rs.getString("apellido1");
                            String space1       = rs.getString("space1");
                            String apellido2    = rs.getString("apellido2");
                            String pipe5        = rs.getString("pipe5");
                            String telefono     = rs.getString("telefono");
                            String pipe6        = rs.getString("pipe6");
                            String date2        = rs.getString("date2");
                            String pipe7        = rs.getString("pipe7");
                            String totaldeuda2  = rs.getString("totaldeuda2");
                            String pipe8        = rs.getString("|0|PAGOMENSUALIDADES");
                            
                            System.out.print(rows+"->");
                            System.out.println(nodoc+""+cod+""+pipe1+""+totaldeuda+""+pipe2+""+concat1+""+date1+""+pipe3+""+primernombre+""+segundonombre+""+pipe4+""+apellido1+""+space1+""+apellido2+""+pipe5+""+telefono+""+pipe6+""+date2+""+pipe7+""+totaldeuda2+""+pipe8);
                            data.add(nodoc+""+cod+""+pipe1+""+totaldeuda+""+pipe2+""+concat1+""+date1+""+pipe3+""+primernombre+""+segundonombre+""+pipe4+""+apellido1+""+space1+""+apellido2+""+pipe5+""+telefono+""+pipe6+""+date2+""+pipe7+""+totaldeuda2+""+pipe8);
                            logdata.add(logdate+" -> "+rows+" -> "+nodoc+""+cod+""+pipe1+""+totaldeuda+""+pipe2+""+concat1+""+date1+""+pipe3+""+primernombre+""+segundonombre+""+pipe4+""+apellido1+""+space1+""+apellido2+""+pipe5+""+telefono+""+pipe6+""+date2+""+pipe7+""+totaldeuda2+""+pipe8);
            }
            
            String file_date = dateFileName();
            writeToFile(data, "/Volumes/sapiens/INTERFASE/PSE_FILES/PSE_FILE_"+file_date+".txt");
            mailFile(data, "/Volumes/sapiens/INTERFASE/PSE_FILES/PSE_FILE_MAIL.txt");
            logFile(logdata, "/Volumes/sapiens/INTERFASE/PSE_FILES/logs/pse_log_"+file_date+".txt");
                
            System.out.println("Table updated in given database...");
            
        }catch(SQLException se){
            //Handle errors for JDBC
            se.printStackTrace();
         }catch(Exception e){
            //Handle errors for Class.forName
            e.printStackTrace();
         }finally{
            con.close();
        }   
        System.out.println("Success: UPDATE mensualidad in table ***importacion_cartera***"); 
    }
    
    private static void writeToFile(java.util.List list, String path) {
            BufferedWriter out = null;
            try {
                    File file = new File(path);
                    out = new BufferedWriter(new FileWriter(file, true));
                    for (Object s : list) {
                            out.write((String) s);
                            out.newLine();
                    }
                    out.close();
            } catch (IOException e) {
            }
    } 
    
    private static void mailFile(java.util.List list, String path){
            BufferedWriter out = null;
            try {
                    File file = new File(path);
                    out = new BufferedWriter(new FileWriter(file, true));
                    for (Object s : list) {
                            out.write((String) s);
                            out.newLine();
                    }
                    out.close();
            } catch (IOException e) {
            }
    }
    
    private static void logFile(java.util.List list, String path){
            BufferedWriter out = null;
            try {
                    File file = new File(path);
                    out = new BufferedWriter(new FileWriter(file, true));
                    for (Object s : list) {
                            out.write((String) s);
                            out.newLine();
                    }
                    out.close();
            } catch (IOException e) {
            }
    }
    
    private static String dateFileName(){
        Date today = new Date();
        Format formatter = new SimpleDateFormat("yyyyMMdd_HHmmss");
        String s = formatter.format(today);
        
        return s;
    } 
    
    public void deleteFileTxt(){
        
        try{
    		File file = new File("/Volumes/sapiens/INTERFASE/PSE_FILES/PSE_FILE_MAIL.txt");
                
    		if(file.delete()){
    			System.out.println(file.getName() + " is deleted!");
    		}else{
    			System.out.println("Delete operation is failed.");
    		}
    	}catch(Exception e){
    		e.printStackTrace();
    	}
    }
}