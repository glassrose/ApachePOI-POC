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

/* Program to recusively read all files in a folder and print their data to console */
package testpoi;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.io.IOException;

/**
 *
 * @author Admin
 */
public class FlatFilesInFolderToconsole {

   static File folder = new File("D:\\Folder");
//   static File excelFile;
    
    public static void main (String args[])
    {
        
        processAllFiles (folder);
    }
    
    private static void processAllFiles(File folder) {
        for (File file: folder.listFiles())
        {
            if (file.isDirectory())
                processAllFiles (file);
            else
            {
                //this is a file. process it.
                try
                {
                    BufferedReader br = new BufferedReader (new FileReader (file));
                    String contents = "", line = null;
                    while ((line = br.readLine()) != null)
                    {
                        contents += line;
                    }
                    System.out.println(contents);
//                    writeToExcel (contents, excelFile);
                }
                catch (Exception e)
                {
                    e.printStackTrace();
                }
            }
        }
    }
    
}
