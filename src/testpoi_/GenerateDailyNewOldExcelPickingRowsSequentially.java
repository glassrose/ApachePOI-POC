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

package testpoi_;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Admin
 */

class Department {
    String name;
    int newCnt;
    boolean isNew;
    
    public Department (String name, int newCnt, boolean isNew)
    {
        this.name = name;
        this.newCnt = newCnt;
        this.isNew = isNew;
    }
}

class OldDepttSheet
{
    XSSFSheet sheet;
    int rowCnt;
    
    public OldDepttSheet(XSSFSheet sheet)
    {
        this.sheet = sheet;
        rowCnt = 1;
    }
}

public class GenerateDailyNewOldExcelPickingRowsSequentially {
    /*************************** TO UPDATE ON EVERY RUN ******************************/
    
    final static String date = "17.1.13";
    
    /*********************************************************************************/
    
    final static String path = "C:\\Documents and Settings\\Admin\\My Documents\\NetBeansProjects\\TestPOI\\Docs\\"+date+"\\";
//    final static String path = "/home/chandni/NetBeansProjects/POI_POC/docs/"+date+"/";
    static XSSFWorkbook workbookOld;
    static XSSFSheet sheetAll;
    static XSSFSheet sheetFemale;
    static XSSFSheet sheetNew;
    static XSSFSheet sheetChildren;
    static ArrayList<Department> deptts;
    static HashMap<String, OldDepttSheet> depttToOldSheetsMap;
    static int rowCnt;
    static int femaleRowNum;
    static int childRowNum;
    static int allRowNum;
    static int crNo;
   
    public static void main (String args[])
    {
        //For Reading
        FileInputStream file1 = null,file2 = null, fileOldIn = null;
        try {
 
            file1 = new FileInputStream(new File(path+"new.xlsx"));

            XSSFWorkbook workbook1 = new XSSFWorkbook (file1);

            //Get first sheet from the workbook1
            sheetAll = workbook1.getSheetAt(0);
            //Get second sheet from the workbook1
            sheetFemale = workbook1.getSheetAt(1);
            
            
            file2 = new FileInputStream(new File(path+"children.xlsx"));

            XSSFWorkbook workbook2 = new XSSFWorkbook (file2);

            //Get first sheet from the workbook2
            sheetChildren = workbook2.getSheetAt(0);
            
            fileOldIn = new FileInputStream(new File(path+"old.xlsx"));
            workbookOld = new XSSFWorkbook (fileOldIn);
             
        }
        catch (Exception e)
        {
            System.err.println ("Error opening files for reading.");
            e.printStackTrace();
        }
        
        //For writing
        XSSFWorkbook workbook = new XSSFWorkbook();
        sheetNew = workbook.createSheet("Generated File - Do not edit");
        //Create a new row in current sheet for heading.
        Row row = sheetNew.createRow(0);
        //Create a new cell in current row
        Cell cell = row.createCell(0);
        //Set value to new value
        cell.setCellValue("Department");
        cell = row.createCell(1);
        cell.setCellValue("Patient Type");
        cell = row.createCell(2);
        cell.setCellValue("CR No.");
        cell = row.createCell(3);
        cell.setCellValue("Name");
        cell = row.createCell(4);
        cell.setCellValue("Guardian's Name");
        cell = row.createCell(5);
        cell.setCellValue("Relation");
        cell = row.createCell(6);
        cell.setCellValue("AgeYrs");
        cell = row.createCell(7);
        cell.setCellValue("Gender");
        cell = row.createCell(8);
        cell.setCellValue("Address");
        cell = row.createCell(9);
        cell.setCellValue("City");
        cell = row.createCell(10);
        cell.setCellValue("State");
        
        rowCnt=1;
        femaleRowNum=1;
        childRowNum=1;
        allRowNum=1;
        
        /************************ TO SET AT EVERY RUN **************************/
        crNo = 1050;
        
        
        deptts = new ArrayList<>();
        /* New */
        deptts.add(new Department("Medicine", 118, true));
        deptts.add(new Department("Surgery", 89, true));
        deptts.add(new Department("Obs & Gynae", 67, true));
        deptts.add(new Department("Paediatrics", 20, true));
        deptts.add(new Department("Orthopaedics", 54, true));
        deptts.add(new Department("Ophthalmology", 33, true));
        deptts.add(new Department("ENT", 28, true));
        deptts.add(new Department("Dental", 27, true));
        deptts.add(new Department("Casualty", 42, true));
        /* Old */
        deptts.add(new Department("Medicine", 15, false));
        deptts.add(new Department("Surgery", 13, false));
        deptts.add(new Department("Obs & Gynae", 12, false));
        deptts.add(new Department("Paediatrics", 9, false));
        deptts.add(new Department("Orthopaedics", 11, false));
        deptts.add(new Department("Ophthalmology", 16, false));
        deptts.add(new Department("ENT", 6, false));
        deptts.add(new Department("Dental", 8, false));
        
        
        
//        Casualty is only new
        
        
        /***********************************************************************/
        
        //Fill depttToOldSheetsMap
        Iterator<XSSFSheet> oldSheetsIter = workbookOld.iterator();
        //Skip 1st sheet which contains all old patients
        oldSheetsIter.next();
        depttToOldSheetsMap = new HashMap<>();
        while (oldSheetsIter.hasNext())
        {
            XSSFSheet oldSheet = oldSheetsIter.next();
            depttToOldSheetsMap.put(oldSheet.getSheetName(), new OldDepttSheet (oldSheet));
        }
        
        try {
            generateRows ();
        }
        catch (IllegalArgumentException e)
        {
            System.err.println(e.getMessage());
            e.printStackTrace();
        }
        
        try {
            FileOutputStream out = new FileOutputStream(new File(path+date+".xlsx"));
            workbook.write(out);
            out.close();
            if (file1!=null)
                file1.close();
            if (file2!=null)
                file2.close();
            System.out.println("Excel written successfully..");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void generateRows() throws IllegalArgumentException{
        int totalRows = 0;
        for (int i=0; i<deptts.size(); i++)
            totalRows+=(deptts.get(i).newCnt);
        
        for (int i=0; i<totalRows /*&& deptts.size()>0*/; i++)
        {
            double random = Math.random();
            int depttNo = (int)(random*deptts.size());
            
            if (deptts.get(depttNo).newCnt == 0)
            {
                deptts.remove(depttNo);
                i--;
                continue;
            }
            
            makeEntry (deptts.get(depttNo));
            deptts.get(depttNo).newCnt--;
        }
    }

    private static void makeEntry(Department deptt) {
        //create new row in xlsx to be generated
        Row newRow = sheetNew.createRow(rowCnt++);
        //Create a new cell in current row
        Cell newCell = newRow.createCell(0);
        //Set value to the department's name
        newCell.setCellValue(deptt.name);
        newCell = newRow.createCell(1);
        newCell.setCellValue(deptt.isNew?"New":"Old");
        
        if (deptt.isNew)
        {
            newCell = newRow.createCell(2);
            newCell.setCellValue(crNo++);

            Row row = null;
            if (deptt.name.equals("Obs & Gynae"))
            {
    //            //Pick a row from female sheet randomly (Female sheet should have all reproducible ages)
    //            int rowNum = (int)(random*sheetFemale.getPhysicalNumberOfRows());
                
                if (femaleRowNum<sheetFemale.getPhysicalNumberOfRows())
                {
                    row = sheetFemale.getRow(femaleRowNum++);
                    System.out.println("Sheet:Female, row: "+row.getRowNum());
                }
                else
                {
                    System.err.println ("Female entries exhausted!");
                }
            }
            else if (deptt.name.equals("Paediatrics"))
            {
                if (childRowNum<sheetChildren.getPhysicalNumberOfRows())
                {
                    row = sheetChildren.getRow(childRowNum++);
                    System.out.println("Sheet:Children, row: "+row.getRowNum());
                }
                else
                {
                    System.err.println ("Child entries exhausted!");
                }
            }
            else
            {
                if (allRowNum<sheetAll.getPhysicalNumberOfRows())
                {
                    row = sheetAll.getRow(allRowNum++);
                    System.out.println("Sheet:All, row: "+row.getRowNum());
                }
                else
                {
                    System.err.println ("All(General New) entries exhausted!");
                }
            }
            if (row==null)
            {
                throw new IllegalArgumentException("New input Rows Exhausted");
            }
            assert row!= null;

            //read and write fetched row
            Iterator<Cell> cellIterator = row.cellIterator();
            int newCellCnt=3;
            while (cellIterator.hasNext())
            {
                //May we write all cells as strings?
                Cell cell = cellIterator.next();
                String cellValue=null;
                try {
                    if (cell.getCellType()==Cell.CELL_TYPE_NUMERIC)
                        cellValue = (int)(cell.getNumericCellValue())+"";
                    else
                        cellValue =  cell.getStringCellValue();

                    newCell = newRow.createCell(newCellCnt++);
                    newCell.setCellValue(cellValue);
                }
                catch (Exception e)
                {
                    System.out.println ("Could not write from cell (value:"+cellValue+
    //                        ", column:"+cell.getSheet().getWorkbook().+
                            ", sheet:"+cell.getSheet().getSheetName()+
                            ", row:"+cell.getRowIndex()+
                            ", column:"+cell.getColumnIndex()+")");
                    e.printStackTrace();
                }
            }
        }
        else //deptt is old
        {
            OldDepttSheet oldDepttSheetToUse = depttToOldSheetsMap.get(deptt.name);
            
            Row row = oldDepttSheetToUse.sheet.getRow(oldDepttSheetToUse.rowCnt++);
            
            if (row==null)
            {
                throw new IllegalArgumentException("Old Input Rows Exhausted in department "+deptt.name);
            }
            
            System.out.println("Sheet:"+deptt.name+", row: "+row.getRowNum());
            
            //Copy row from old sheet to newRow
            int newCellCnt=2;
            Iterator<Cell> cellIterator = row.cellIterator();
            //Skip columns Department and Patient Type
            cellIterator.next();
            cellIterator.next();
            
            while (cellIterator.hasNext())
            {
                Cell cell = cellIterator.next();
                String cellValue=null;
                try {
                    if (cell.getCellType()==Cell.CELL_TYPE_NUMERIC)
                        cellValue = (int)(cell.getNumericCellValue())+"";
                    else
                        cellValue =  cell.getStringCellValue();

                    newCell = newRow.createCell(newCellCnt++);
                    newCell.setCellValue(cellValue);
                }
                catch (Exception e)
                {
                    System.out.println ("Could not write from old sheet cell (value:"+cellValue+
                            ", sheet:"+cell.getSheet().getSheetName()+
                            ", row:"+cell.getRowIndex()+
                            ", column:"+cell.getColumnIndex()+")");
                    e.printStackTrace();
                }
                
            }
        }
    }
}
