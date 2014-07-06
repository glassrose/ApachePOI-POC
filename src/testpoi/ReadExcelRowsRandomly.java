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
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/* Read rows randomly from an excel file containing strings and numeric values */

/**
 *
 * @author Admin
 */
public class ReadExcelRowsRandomly {
    public static void main (String args[])
    {
        try {
 
            FileInputStream file = new FileInputStream(new File("C:\\Documents and Settings\\Admin\\My Documents\\NetBeansProjects\\TestPOI\\Docs\\1.xlsx"));

            //Get the workbook instance for XLS file 
            XSSFWorkbook workbook = new XSSFWorkbook (file);

            //Get first sheet from the workbook
            XSSFSheet sheet = workbook.getSheetAt(0);
            
            double random = Math.random ();
            int rowNum = (int)(random*sheet.getPhysicalNumberOfRows());
            
            Row row = sheet.getRow(rowNum);
            
            Iterator<Cell> cellIterator = row.cellIterator();
            
            while (cellIterator.hasNext())
            {
                Cell cell = cellIterator.next();
                if (cell.getCellType()== Cell.CELL_TYPE_NUMERIC)
                {
                    double cellValue = cell.getNumericCellValue();
                    System.out.print(cellValue+"\t");
                }
                else if (cell.getCellType()== Cell.CELL_TYPE_STRING)
                {
                    String cellValue = cell.getStringCellValue();
                    System.out.print(cellValue+"\t");
                }
            }
            System.out.println ();
                       
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }
}
