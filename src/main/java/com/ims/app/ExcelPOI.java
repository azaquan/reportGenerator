package com.ims.app;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

import java.sql.*;

import javax.naming.Context;
import javax.naming.InitialContext;
import javax.sql.DataSource;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.model.Sheet;
import org.apache.poi.hssf.usermodel.*;

/**
* @author Praveen John
*
*/

public class ExcelPOI{
   private ResultSet result = null;

   public void ExcelPOI(){
   }

   public void createReport(ResultSet r) throws IOException {
      result=r;
      int j=0;
      try{
         //File Operation - Open the excel template
         File file = new File("sampleExcel.xls");
         InputStream fileinput= new FileInputStream(file);

         HSSFWorkbook workbook = (HSSFWorkbook) WorkbookFactory.create(fileinput);
         HSSFSheet sheet = workbook.getSheetAt(0); //Get 1st Sheet

         //Iterate through the ResultSet and update the Excel report

         while(result.next())
         {
            HSSFRow row = sheet.createRow(j+1); //first row is header
            for(int i=1;i<=6;i++) // Update 6 columns mentioned in screen shot
           {
             HSSFCell cell = row.createCell(0);
             cell.setCellValue(result.getString(i));//set value for each cell
           }
           j=j+1;
          }
         //Save data
         FileOutputStream fileOut = new FileOutputStream("incident_result.xls");
         workbook.write(fileOut);
         fileOut.close();
      }
      catch(Exception ex)
      {
          //Print exception to standard output
          System.out.println(ex);
      }
   }
}