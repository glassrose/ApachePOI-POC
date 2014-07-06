/*
 * The MIT License
 *
 * Copyright 2014 Admin.
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 */

package testpoi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.Statement;


/**
 *
 * @author chandni
 */
public class POI_MySQL_Test {
    public static void main (String args[])
    {
        try {
 
            FileInputStream file = new FileInputStream(new File("C:\\Documents and Settings\\Admin\\My Documents\\NetBeansProjects\\TestPOI\\Docs\\OPD_NEW_2.xlsx"));

            //Get the workbook instance for XLS file 
            XSSFWorkbook workbook = new XSSFWorkbook (file);

            //Get first sheet from the workbook
            XSSFSheet sheet = workbook.getSheetAt(0);
            
            Connection conn = connectToDatabase ();
            assert (conn != null);

            //Get iterator to all the rows in current sheet
            Iterator<Row> rowIterator = sheet.iterator();
            
            //Skip the 1st row
            rowIterator.next();
            
            while(rowIterator.hasNext()) {
                Row row = rowIterator.next();

                //For each row, get values of each column
                Iterator<Cell> cellIterator = row.cellIterator();
                Cell cell = cellIterator.next();
                int crNo;
                if(cell.getCellType()==0)
                     crNo = (int)(cell.getNumericCellValue());
                else
                {
                    System.out.println ("crNo cell value: "+cell.getStringCellValue());
                    crNo = (int)Integer.parseInt(cell.getStringCellValue().trim());
                }
                
                cell = cellIterator.next();
                String dept;
                if(cell.getCellType()==1)
                    dept = cell.getStringCellValue();
                else
                {
                    int no = (int)cell.getNumericCellValue();
                    dept = no+"";
                    System.out.println(dept);
                }
                cell = cellIterator.next();
                String name = cell.getStringCellValue();
                cell = cellIterator.next();
                String guardian = cell.getStringCellValue();
                cell = cellIterator.next();
                String rel = cell.getStringCellValue();
                cell = cellIterator.next();
                System.out.println("\n cell.getCellType() :"+ cell.getCellType());
                int ageYrs=0;
                if(cell.getCellType()==0)
                     ageYrs = (int)(cell.getNumericCellValue());
                else
                {
                    System.out.println ("age cell value: "+cell.getStringCellValue());
                    ageYrs = (int)Integer.parseInt(cell.getStringCellValue().trim());
                }
                
                cell = cellIterator.next();
                String gender = cell.getStringCellValue();
                cell = cellIterator.next();                
                String add = cell.getStringCellValue();
                cell = cellIterator.next();                
                String city = cell.getStringCellValue();
                cell = cellIterator.next();                
                String state = cell.getStringCellValue();
                
                
                try {
                    Statement st = conn.createStatement();
                    String insertSql = "INSERT INTO OPDData VALUES(" + crNo + ",'"+
                            dept+"','"+name+"','" +guardian+"','"+rel+"',"+
                            ageYrs+",'"+gender+"','"+add+"','"+city+"','"+state+"')";
                    System.out.println (insertSql);
                    int val = st.executeUpdate(insertSql);
                    System.out.println("One row get affected...");

                }
                catch (SQLException ex) {
                    System.out.println("Cannot connect to database server...!!");
                    ex.printStackTrace();
                }                
            }
            file.close();
            
            if(conn != null){
                try {
                    // close() releases this Connection object's database and JDBC resources immediately instead of waiting for them to be automatically released.
                    conn.close();
                    System.out.println ("Database connection terminated...!!!");
                } catch (SQLException ex) {
                    ex.printStackTrace();
                }
            }

        }
        catch (FileNotFoundException e)
        {
            e.printStackTrace();
        }
        catch (IOException e)
        {
            e.printStackTrace();
        }
    }

    private static Connection connectToDatabase() {
        Connection conn = null;

//        String dbUserName = "root"; // MySQL database username
//        String dbPassword = "21120509"; // MySQL database password
        String dbUserName = "sa"; // MySQL database username
        String dbPassword = "mcsgoc123"; // MySQL database password
        // Actually dbUrl variable can be divided into three categories for understanding purpose
        // "jdbc:mysql://" is a required syntax to create the connection with MySQL
        // "localhost/" this part is the "datadir" attribute that we have defined within my.ini file
        // I am using WAMP server 2.1 and you can see the above attribute in line 39 within my.ini file
        // "HMS" this is the name of my database. You can define this as a different variable too.
//        String dbUrl = "jdbc:mysql://localhost/HMS";
        String dbUrl = "jdbc:sqlserver://localhost;databaseName=HMS;";

        try {
            // forName() method is static.
            // It returns the Class object associated with the class or interface with the given string name.
            //Class forNam = Class.forName("com.mysql.jdbc.Driver");
            Class forNam = Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");

            try {
                // newInstance() creates a new instance of the class represented by this Class object
                forNam.newInstance();
            } catch (InstantiationException ex) {
                ex.printStackTrace();
            } catch (IllegalAccessException ex) {
                ex.printStackTrace();
            }
        } catch (ClassNotFoundException ex) {
            ex.printStackTrace();
        }

        try {
            // DriverManager is the basic service for managing a set of JDBC drivers.
            // getConnection() attempts to establish a connection to the given database URL
            conn = DriverManager.getConnection(dbUrl, dbUserName, dbPassword);
            System.out.println("Database connection establish...!");

            
        } catch (SQLException ex) {
            System.out.println("Cannot connect to database server...!!");
            ex.printStackTrace();
        }
        
        return conn;
    }
}
