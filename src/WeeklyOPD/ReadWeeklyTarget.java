/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package WeeklyOPD;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Chandni
 */



public class ReadWeeklyTarget {
//    static String path = "C:\\Users\\Chandni\\Documents\\NetBeansProjects\\TestPOI\\Docs\\March2014\\";
    static String path = "/home/chandni/NetBeansProjects/TestPOI/Docs/Weekly Data/May2014/";
    static String targetFile;
    static int weekNumber;
    
    static FileInputStream file1, file2, fileOldIn;
    static XSSFWorkbook workbookOld;
    static XSSFWorkbook workbook;//for writing
    static XSSFSheet sheetAll;
    static XSSFSheet sheetFemale;
    static XSSFSheet sheetNew;
    static XSSFSheet sheetChildren;
    static HashMap<String, OldDepttSheet> depttToOldSheetsMap;
    static int rowCnt;
    static int femaleRowNum;
    static int childRowNum;
    static int allRowNum;
    static int crNo;
    
    static FileOutputStream out;//for writing
    
    public static void main (String args[])
    {
        /******************************* TO UPDATE ON EACH RUN ***************************/
        targetFile = "May 2014.xlsx";
        weekNumber = 1;
        femaleRowNum=1212;
        childRowNum=749;
        allRowNum=7414;
        crNo = 187394;//crNo to begin with
        

        /*********************************************************************************/
        
        rowCnt=1;
        GenerateDailyNewOldExcelPickingRowsSequentially.mainCreateExcelAndInitialize();
        
        try
        {
            FileInputStream targetFileIn = new FileInputStream(path+targetFile);
            XSSFWorkbook workbookTarget = new XSSFWorkbook (targetFileIn);
            XSSFSheet sheetTarget = workbookTarget.getSheetAt(0);
            
            out = new FileOutputStream(new File(path+"Week"+weekNumber+targetFile));
                
            
            int startAtRow = weekNumber==1?4:weekNumber==2?11:weekNumber==3?18:weekNumber==4?25:32;
            int rowsToRead = 7;
            
            System.out.println("Physical No. Of Rows: "+sheetTarget.getPhysicalNumberOfRows());
            for (int rowNum = startAtRow; rowNum<=startAtRow+rowsToRead-1 && rowNum<sheetTarget.getPhysicalNumberOfRows(); rowNum++)
            {
                System.out.println ("RowNum: "+rowNum);
                Row row = sheetTarget.getRow(rowNum);
                
//                int date = (int) row.getCell(0).getNumericCellValue();
////                String[] splitDate = date.split("/");
//                System.out.println (date);
////                System.out.println (splitDate.length);
//                int day = Integer.parseInt(splitDate[0]);
//                int month = Integer.parseInt(splitDate[1]);
//                int year = Integer.parseInt(splitDate[2]);
                
                int medicineNew = (int) row.getCell(1).getNumericCellValue();
                int surgeryNew = (int) row.getCell(4).getNumericCellValue();
                int ophthalmologyNew = (int) row.getCell(7).getNumericCellValue();
                int entNew = (int) row.getCell(10).getNumericCellValue();
                int paediatricsNew = (int) row.getCell(13).getNumericCellValue();
                int ogNew = (int) row.getCell(16).getNumericCellValue();
                int orthopaedicsNew = (int) row.getCell(19).getNumericCellValue();
                int dentalNew = (int) row.getCell(22).getNumericCellValue();
                int casualtyNew = (int) row.getCell(25).getNumericCellValue();
                
                int medicineOld = (int) row.getCell(2).getNumericCellValue();
                int surgeryOld = (int) row.getCell(5).getNumericCellValue();
                int ophthalmologyOld = (int) row.getCell(8).getNumericCellValue();
                int entOld = (int) row.getCell(11).getNumericCellValue();
                int paediatricsOld = (int) row.getCell(14).getNumericCellValue();
                int ogOld = (int) row.getCell(17).getNumericCellValue();
                int orthopaedicsOld = (int) row.getCell(20).getNumericCellValue();
                int dentalOld = (int) row.getCell(23).getNumericCellValue();
                
                System.out.println(medicineNew+"\t"+medicineOld+
                        "\t"+surgeryOld+"\t"+surgeryNew+
                        "\t"+surgeryOld+"\t"+surgeryNew+
                        "\t"+ophthalmologyOld+"\t"+ophthalmologyNew+
                        "\t"+entOld+"\t"+entNew+
                        "\t"+paediatricsOld+"\t"+paediatricsNew+
                        "\t"+ogOld+"\t"+ogNew+
                        "\t"+orthopaedicsOld+"\t"+orthopaedicsNew+
                        "\t"+dentalOld+"\t"+dentalNew+
                        "\t"+casualtyNew
                        );
                
                
                GenerateDailyNewOldExcelPickingRowsSequentially.mainGenerateExcel(medicineNew, medicineOld, surgeryNew, surgeryOld, ophthalmologyNew, ophthalmologyOld, entNew, entOld, paediatricsNew, paediatricsOld, ogNew, ogOld, orthopaedicsNew, orthopaedicsOld, dentalNew, dentalOld, casualtyNew);
                
                
            }
            workbook.write(out);
                
            GenerateDailyNewOldExcelPickingRowsSequentially.writeExcelAndcloseFiles();

            System.out.println ("Week "+weekNumber+" Excel generated successfully!");
            System.out.println ("New Row Numbers to start with:");
            System.out.println ("All: "+allRowNum+"\tFemale: "+femaleRowNum+"\tChild: "+childRowNum);
            System.out.println ("New CrNo. to start with: "+ crNo);
            
            
        }
        catch (Exception e)
        {
            System.err.println ("Error reading target!");
            e.printStackTrace();
        }
        
        
    }
}
