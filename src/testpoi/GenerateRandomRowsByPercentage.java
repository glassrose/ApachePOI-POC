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

/* Program to generate rows randomly based on composition percentages of rowtypes */
package testpoi;

import java.util.ArrayList;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
/**
 *
 * @author Admin
 */

class RowType
{
    String str;
    int cnt;
    
    public RowType (String str, int cnt)
    {
        this.str = str;
        this.cnt = cnt;
    }
}

public class GenerateRandomRowsByPercentage {
    public static void main (String args[])
    {
        int totalRows = 13;
        ArrayList<RowType> rowTypes = new ArrayList <>();
        
        int composition[] = {10, 20, 40, 30};
        String[] rowValues = {"Chandni", "Milan", "Rashmi", "Tauseef"};
        int len = composition.length;
        
        for (int i=0;i<len;i++)
        {
            rowTypes.add (new RowType(rowValues[i], Math.round(composition[i]*totalRows/100)));
        }
        
        int i;
        for (i=0; i<totalRows && rowTypes.size()>0; i++)
        {
            double r = Math.random();
            int rowTypeIndexSelected = (int)(r*rowTypes.size());
            
            if (rowTypes.get(rowTypeIndexSelected).cnt == 0)
            {
                rowTypes.remove(rowTypeIndexSelected);
                i--;
                continue;
            }
            
            System.out.println (rowTypes.get(rowTypeIndexSelected).str);
            rowTypes.get(rowTypeIndexSelected).cnt--;
        }
        
        //To deal with fractional values, generate a single value for all remaining rows
        for (;i<totalRows;i++)
        {
            System.out.println ("Chandni");
        }
    }
}
