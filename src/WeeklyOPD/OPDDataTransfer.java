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

package WeeklyOPD;

import testpoi_.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.Connection;
import java.sql.Date;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.Statement;
import java.sql.Time;
import java.sql.Timestamp;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Admin
 */
public class OPDDataTransfer {
    
    /*************************** TO UPDATE ON EVERY RUN ******************************/
    static final Date date = Date.valueOf("2014-02-04");
    static Time time;
    static final String OPDDATE = "04022014";
    static String dateFolder = "4.2.14";
    /********************************************************************************/
    
    static int entryNumber;
//    static int entryCrNo;
    static String[] loginUsers = {"Anand",
                                "Hemant",
                                "Priyanka",
                                "Abhishek",
                                "Admin",
                                "Anuj",
                                "Arti",
                                "Arvind",
                                "Beauty",
                                "Gangaram",
                                "GangaSagar",
                                "LKPandey",
                                "mohit",
                                "Mohitbajaj",
                                "Neha",
                                "Pankaj",
                                "PankajG",
                                "Rajan",
                                "Renu",
                                "Rishabh",
                                "Ruchi",
                                "Shivraj",
                                "ShiyaRam"
                                };
    
    public static void main (String args[])
    {
        try {
 
            FileInputStream file = new FileInputStream(new File("C:\\Documents and Settings\\Admin\\My Documents\\NetBeansProjects\\TestPOI\\Docs\\"+
                    dateFolder+"\\"+dateFolder+".xlsx"));

            //Get the workbook instance for XLS file 
            XSSFWorkbook workbook = new XSSFWorkbook (file);

            //Get first sheet from the workbook
            XSSFSheet sheet = workbook.getSheetAt(0);
            
            Connection conn = connectToDatabaseHMS();
            assert (conn != null);
            Connection conn2 = connectToDatabaseHMSSecurity();
            assert (conn2 != null);
            
            try
            {
                conn.setAutoCommit(false);
                conn2.setAutoCommit(false);
            }
            catch (SQLException sqle)
            {
                System.err.println ("Could not set autocommit to false");
                sqle.printStackTrace();
            }

            //Get iterator to all the rows in current sheet
            Iterator<Row> rowIterator = sheet.iterator();
            
            //Skip the 1st row
            rowIterator.next();
            
            /*************************** TO UPDATE ON ERROR ******************************/
            //set time
            time = Time.valueOf("09:00:00");
            //set entry number default to 1.
            entryNumber = 1;
            
            /*****************************************************************************/
           
//            //set entry crNo
//            entryCrNo = 1;
            Timestamp timestamp = new Timestamp (date.getTime()+time.getTime() + 19800000/*for IST*/);
            System.out.println(timestamp.toString());
            
            while(rowIterator.hasNext()) {
                Row row = rowIterator.next();

                //For each row, get values of each column
                Iterator<Cell> cellIterator = row.cellIterator();
                
                Cell cell = cellIterator.next();
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
                String type = cell.getStringCellValue();
                cell = cellIterator.next();
                
                int crNo;
                if(cell.getCellType()==0)
                     crNo = (int)(cell.getNumericCellValue());
                else
                {
                    System.out.println ("crNo cell value: "+cell.getStringCellValue());
                    crNo = (int)Integer.parseInt(cell.getStringCellValue().trim());
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
                
                int deptID = getDeptID (dept);
                    assert (deptID != 0);
                int stateID = getStateID (state);
                    assert (stateID != 0);
                int drID = getDrID (deptID);
                    assert (drID != 0); // if drID = 0 that means a dept. has been entered which doesn't have a doctor
                String loginUserName = getLoginUserName ();
                String userID = getLoginUserID (loginUserName);
                    
                long OPDNo = Long.parseLong (OPDDATE+entryNumber);
                
                boolean queryExecuted = true;
                if (type.equals("New"))  // As only New patients must be registered and have a CR generated
                {
                    try {
                        String insertSql = "INSERT INTO Reg "
                                + "(Regno, Name, Fname, Relation, AgeY, Sex, Address1, City, State, Department, Date)"
                                + "VALUES(" + crNo + ",'"+
                                name+"','" +guardian+"','"+rel+"',"+
                                ageYrs+",'"+gender+"','"+add+"','"+city+"',"+stateID+","+deptID+",'"+date+"')";
                        System.out.println (insertSql);
                        Statement st = conn.createStatement();
                        int val = st.executeUpdate(insertSql);
                        System.out.println("One row in Reg gets affected...");

                    }
                    catch (SQLException ex) {
                        queryExecuted = false;
                        System.out.println("Cannot insert row into Reg...!!");
                        ex.printStackTrace();
                    }
                }
                try {
                    String insertSql = "INSERT INTO OPD "
                            + "(OPDNo, CrNo, PatientType, DepartmentId, DrId, Date, Time, LoginUserName, IsActive)"
                            + "VALUES("+ OPDNo+ "," + crNo + ",'"+type+"',"+deptID+ "," +drID+ ",'"
                            + date +"','"+time+"','"+loginUserName+"','"+true+"')";
                    System.out.println (insertSql);
                    Statement st = conn.createStatement();
                    int val = st.executeUpdate(insertSql);
                    System.out.println("One row in OPD gets affected...");
                }
                catch (SQLException ex) {
                    queryExecuted = false;
                    System.out.println("Cannot insert row into OPD...!!");
                    ex.printStackTrace();
                }
                
                try {
                    String updateSql = "UPDATE aspnet_Users "
                            + "SET LastActivityDate='"+timestamp+"'"
                            + "WHERE UserId='"+userID+"'";
                    System.out.println (updateSql);
                    Statement st = conn2.createStatement();
                    int val = st.executeUpdate(updateSql);
                    System.out.println("One row in aspnet_Users gets affected...");
                }
                catch (SQLException ex) {
                    queryExecuted = false;
                    System.out.println("Cannot update timestamp in aspnet_Users...!!");
                    ex.printStackTrace();
                }
                
                if (!queryExecuted)
                    //if insertion to any table fails, rollback.
                    try
                    {
                        conn.rollback();
                        conn2.rollback();
                        
                        break; // and run program again at any error
                    }
                    catch (SQLException ex)
                    {
                        ex.printStackTrace();
                    }
                else
                    try
                    {
                        conn.commit();
                        conn2.commit();
                    
                        entryNumber++;
//                        entryCrNo++;
                        time = new Time(time.getTime()+35000);//add 35 seconds
                        timestamp = new Timestamp (timestamp.getTime()+35000);
                    }
                    catch (SQLException ex)
                    {
                        ex.printStackTrace();
                    }
                
            }
            file.close();
            
            if(conn != null){
                try {
                    // close() releases this Connection object's database and JDBC resources immediately instead of waiting for them to be automatically released.
                    conn.close();
                    System.out.println ("HMS Database connection terminated...!!!");
                } catch (SQLException ex) {
                    ex.printStackTrace();
                }
            }
            if(conn2 != null){
                try {
                    // close() releases this Connection object's database and JDBC resources immediately instead of waiting for them to be automatically released.
                    conn2.close();
                    System.out.println ("HMS_Security Database connection terminated...!!!");
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

    private static Connection connectToDatabaseHMS() {
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
        String dbUrl = "jdbc:sqlserver://mcs_pc;databaseName=HMS;";

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
            System.out.println("Database connection established with HMS...!");

            
        } catch (SQLException ex) {
            System.out.println("Cannot connect to database HMS...!!");
            ex.printStackTrace();
        }
        
        return conn;
    }
    
    
    private static Connection connectToDatabaseHMSSecurity() {
        Connection conn = null;

//        String dbUserName = "root"; // MySQL database username
//        String dbPassword = "21120509"; // MySQL database password
        String dbUserName = "sa"; // MySQL database username
        String dbPassword = "mcsgoc123"; // MySQL database password
        // Actually dbUrl variable can be divided into three categories for understanding purpose
        // "jdbc:mysql://" is a required syntax to create the connection with MySQL
        // "localhost/" this part is the "datadir" attribute that we have defined within my.ini file
        // I am using WAMP server 2.1 and you can see the above attribute in line 39 within my.ini file
        // "HMS_Security" this is the name of my database. You can define this as a different variable too.
//        String dbUrl = "jdbc:mysql://localhost/HMS";
//        String dbUrl = "jdbc:sqlserver://localhost;databaseName=HMS_Security;";
        String dbUrl = "jdbc:sqlserver://mcs_pc;databaseName=HMS_Security;";

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
            System.out.println("Database connection established with HMS_Security...!");

            
        } catch (SQLException ex) {
            System.out.println("Cannot connect to database HMS_Security...!!");
            ex.printStackTrace();
        }
        
        return conn;
    }

    private static int getDeptID(String dept) {
        dept = dept.trim().toLowerCase();
        switch (dept)
        {
            case "medicine": return 1;
            case "obs & gynae": return 2;
            case "surgery": return 3;
            case "paediatrics": return 4;
            case "pediatrics": return 4;
            case "ent": return 5;
            case "ophthalmology": return 6;
            case "opthalmology": return 6;
            case "orthopaedics": return 7;
            case "orthopedics": return 7;
            case "radiology": return 8;
            case "pathology": return 9;
            case "blood bank": return 10;
            case "anaesthesiology": return 11;
            case "forensic medicine": return 12;
            case "dental": return 16;
            case "casualty": return 17;             
            default:  
                System.err.println ("Not a department");
                return 0;
        }
    }
    
    private static int getStateID(String state) {
        state = state.trim().toLowerCase();
        switch (state)
        {
            case "andhra pradesh": return 1;
            case "arunachal pradesh": return 2;
            case "assam": return 3;
            case "bihar": return 4;
            case "chhattisgarh": return 5;
            case "goa": return 6;
            case "gujarat": return 7;
            case "haryana": return 8;
            case "himachal pradesh": return 9;
            case "jammu and kashmir": return 10;
            case "jharkhand": return 11;
            case "karnataka": return 12;
            case "kerala": return 13;
            case "madhya pradesh": return 14;
            case "maharashtra": return 15;
            case "manipur": return 16;
            case "meghalaya": return 17;
            case "mizoram": return 18;
            case "nagaland": return 19;
            case "odisha": return 20;
            case "punjab": return 21;
            case "rajasthan": return 22;
            case "sikkim": return 23;
            case "tamil nadu": return 24;
            case "tripura": return 25;
            case "uttar pradesh": return 26;
            case "uttarakhand": return 27;
            case "west bengal": return 28;
            default: 
                System.err.println ("Not a state");
                return 0;
        }
    }

    private static int getDrID(int deptID) {
        int drID = 0;
        
        int medicineDrs[] = {1, 8};
        int ogDrs[] = {9, 14};
        int surgeryDrs[] = {4};
        int paediaDrs[] = {2, 3};
        int entDrs[] = {5};
        int opthaDrs[] = {6};
        int orthoDrs[] = {12};
        int radiologyDrs[] = {};
        int pathologyDrs[] = {};
        int bbDrs[] = {};
        int anaesthesiologyDrs[] = {};
        int fmDrs[] = {};
        int dentalDrs[] = {7};
        int casualtyDrs[] = {1, 2, 3, 4, 5, 6, 7, 8, 9, 12, 13, 14};
        
        double prn = Math.random();
        
        switch (deptID)
        {
            case 1://medicine
                drID =  medicineDrs[(int)(prn * medicineDrs.length)];
                break;
            case 2:
                drID =  ogDrs[(int)(prn * ogDrs.length)];
                break;
            case 3:
                drID =  surgeryDrs[(int)(prn * surgeryDrs.length)];
                break;
            case 4:
                drID =  paediaDrs[(int)(prn * paediaDrs.length)];
                break;
            case 5:
                drID =  entDrs[(int)(prn * entDrs.length)];
                break;
            case 6:
                drID =  opthaDrs[(int)(prn * opthaDrs.length)];
                break;
            case 7:
                drID =  orthoDrs[(int)(prn * orthoDrs.length)];
                break;
//            case 8:
//                drID =  radiologyDrs[(int)(prn * radiologyDrs.length)];
//                break;
//            case 9:
//                drID =  pathologyDrs[(int)(prn * pathologyDrs.length)];
//                break;
//            case 10:
//                drID =  bbDrs[(int)(prn * bbDrs.length)];
//                break;
//            case 11:
//                drID =  anaesthesiologyDrs[(int)(prn * anaesthesiologyDrs.length)];
//                break;
//            case 12:
//                drID =  fmDrs[(int)(prn * fmDrs.length)];
//                break;
            case 16:
                drID =  dentalDrs[(int)(prn * dentalDrs.length)];
                break;
            case 17:
                drID =  casualtyDrs[(int)(prn * casualtyDrs.length)];
                break;
            default:
                System.err.println ("Department not found");
        }
        return drID;
    }

    private static String getLoginUserName() {

        return loginUsers[(int)(Math.random()*loginUsers.length)];
    }

    private static String getLoginUserID(String loginUserName) {
        String userID = null;
        switch (loginUserName)
        {
            case    "Anand":
                userID = "702d4807-57cd-407e-816c-6f9103d05d66";
                break;
            case    "Hemant":
                userID = "856302af-fc5f-44f6-bd32-a051af7e8732";
                break;
            case    "Priyanka":
                userID = "6bfdc703-9aa1-45d0-85a8-0dd92216d000";
                break;
            case    "Abhishek":
                userID = "79481970-e6c4-438b-a32b-bc9a6f1ddf1e";
                break;
            case    "Admin":
                userID = "eae2c5e0-6eac-4f8d-bd18-da045f0d4738";
                break;
            case    "Anuj":
                userID = "d4f2f2f5-8809-4eb8-bc0b-b3df0771c351";
                break;
            case    "Arti":
                userID = "c516e50a-3a0d-4114-8efb-0590027db190";
                break;
            case    "Arvind":
                userID = "883feeb3-8bfb-43ee-addb-ae85e1962ba6";
                break;
            case    "Beauty":
                userID = "6bffca41-6a70-43a5-8f7f-05621414a865";
                break;
            case    "Gangaram":
                userID = "f6da1a84-8d35-48e0-8446-1c14451266a3";
                break;
            case    "GangaSagar":
                userID = "c5ffd090-1bbc-4cb2-bffd-61c0f15e282d";
                break;
            case    "LKPandey":
                userID = "665bf2c7-59c3-4f6f-bd66-28d0c3ae7a4c";
                break;
            case    "mohit":
                userID = "c906f5b5-6edf-49de-9617-5e64aa3c5754";
                break;
            case    "Mohitbajaj":
                userID = "45de58bb-75e3-4db1-b0fb-a30561bc5e01";
                break;
            case    "Neha":
                userID = "580f2b67-1dce-40ea-b052-7d1e7361f6e2";
                break;
            case    "Pankaj":
                userID = "1da823b6-fe30-4e07-8fd4-e6bde479fe9d";
                break;
            case    "PankajG":
                userID = "0f25b500-6109-41d0-adef-9af45bb2f426";
                break;
            case    "Rajan":
                userID = "444401e0-5371-44e9-85b2-6083591a132d";
                break;
            case    "Renu":
                userID = "48190cd6-0a64-48a5-9c5c-963f3b0f4aa8";
                break;
            case    "Rishabh":
                userID = "7e009588-d812-409d-890a-e2d9d55a6578";
                break;
            case    "Ruchi":
                userID = "411533b4-e5e0-4d11-b04f-b60c44eca7cc";
                break;
            case    "Shivraj":
                userID = "4c0081b6-4af8-4531-a4e9-08e6364980b4";
                break;
            case    "ShiyaRam":
                userID = "07f7355a-0401-485f-8e25-f8cb9a24f886";
                break;
                
            default:
                System.err.println ("LoginUserName and thus corresponding LoginUserID not found.");
                
        }
        
        return userID;
    }
}
