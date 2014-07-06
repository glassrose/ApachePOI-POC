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
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Admin
 */
public class SplitOldDepartmentwise {
    
    /*************************** TO UPDATE ON EVERY RUN ******************************/
    
    final static String date = "1.2.14";
    
    /*********************************************************************************/
    
    final static String path = "C:\\Documents and Settings\\Admin\\My Documents\\NetBeansProjects\\TestPOI\\Docs\\"+date+"\\";
//    final static String path = "/home/chandni/NetBeansProjects/POI_POC/docs/"+date+"/";
    static XSSFWorkbook workbookOld;
    static XSSFSheet sheetAllOld;
    static XSSFSheet sheets[];
    public static void main (String args[])
    {
        FileInputStream fileOldIn;
        FileOutputStream fileOldOut;
        try
        {
            fileOldIn = new FileInputStream(new File(path+"old.xlsx"));
            workbookOld = new XSSFWorkbook (fileOldIn);
            sheetAllOld = workbookOld.getSheetAt(0);
            sheets = new XSSFSheet[9];
            fileOldOut = new FileOutputStream(new File(path+"old.xlsx"));
            createDepartmentwiseSheets ();
            workbookOld.write(fileOldOut);
            fileOldIn.close();
            fileOldOut.close();
            System.out.println("old.xlsx split successfully..");
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }

    private static void createDepartmentwiseSheets() {
        HashMap<String, Integer> depttMap = new HashMap<>();
        depttMap.put("Medicine", 1);
        depttMap.put("Surgery", 2);
        depttMap.put("Obs & Gynae", 3);
        depttMap.put("Paediatrics", 4);
        depttMap.put("Orthopaedics", 5);
        depttMap.put("Ophthalmology", 6);
        depttMap.put("ENT", 7);
        depttMap.put("Dental", 8);
        depttMap.put("Casualty", 9);
        
        int depttSheetCreateFlag = 0;
        
        Iterator<Row> rowIterator = sheetAllOld.rowIterator();
        //Store the first row to be printed as it is.
        ArrayList<String> heading = new ArrayList<>();
        Row row = rowIterator.next();
        Iterator<Cell> cellIterator = row.cellIterator();
        while (cellIterator.hasNext())
            heading.add(cellIterator.next().getStringCellValue());
        
        int rowNums[] = {1,1,1,1,1,1,1,1,1};
        while (rowIterator.hasNext())
        {
            row = rowIterator.next();
            XSSFSheet sheetToWrite = null;
            
            Cell cell = row.getCell(0);
            if ((depttSheetCreateFlag & 1<<(depttMap.get(cell.getStringCellValue()))) == 0)
            {
                //that means this deptt came in this sheet for the first time in this row.
                XSSFSheet sheet = sheets[depttMap.get(cell.getStringCellValue())-1] =
                        workbookOld.createSheet(cell.getStringCellValue());
                //create heading row in this sheet
                Row headingRow = sheet.createRow(0);
                for (int i=0; i<heading.size(); i++)
                {
                    String cellString = heading.get(i);
                    Cell headingCell = headingRow.createCell(i);
                    headingCell.setCellValue(cellString);//sets cell type to string too
                }
                //mark this deptt. as seen
                depttSheetCreateFlag |= (1<<(depttMap.get(cell.getStringCellValue())));
            }

            int sheetNum = depttMap.get(cell.getStringCellValue())-1;
            sheetToWrite = sheets[sheetNum];
            assert (sheetToWrite!=null);
            
            //write row to sheetToWrite
            Row rowNew = sheetToWrite.createRow(rowNums[sheetNum]++);
            
            cellIterator = row.cellIterator();
            int cellNum = 0;
            while (cellIterator.hasNext())
            {
                cell = cellIterator.next();
                
                //write cell
                Cell cellNew = rowNew.createCell(cellNum++);
                String cellValue;
                if (cell.getCellType()==Cell.CELL_TYPE_NUMERIC)
                    cellValue = (int)(cell.getNumericCellValue())+"";
                else
                    cellValue = cell.getStringCellValue();
                cellNew.setCellValue(cellValue);
            }
        }
    }
}
