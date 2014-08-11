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

public class ExcelPOI{
   private ResultSet result = null;

   public void ExcelPOI(){
   }

   public void createReport(ResultSet r) throws IOException {
      result=r;
      int j=0;
      File file = null;
      InputStream fileinput = null;
      HSSFWorkbook workbook = null;
      HSSFSheet sheet = null;
      try{
         String path = System.getProperty("java.class.path"); //TODO
         file = new File("target/templates/sampleExcel.xls");
         fileinput= new FileInputStream(file);
      }catch(IOException e){
          System.out.println(e);
      }

      if(fileinput!=null){
         try{
            workbook = (HSSFWorkbook) WorkbookFactory.create(fileinput);
            sheet = workbook.getSheetAt(0);
         }catch(Exception e){
            System.out.println(e);
         }

         try{
            while(result.next()){
               HSSFRow row = sheet.createRow(j+1);
               for(int i=0;i<=result.getMetaData().getColumnCount()-1;i++){
                  HSSFCell cell = row.createCell(i);
                  cell.setCellValue(result.getString(i+1));
               }
               j=j+1;
             }
         }catch(SQLException e){
          System.out.println(e);
         }
         try{
            FileOutputStream fileOut = new FileOutputStream("target/templates/incident_result.xls");
            workbook.write(fileOut);
            fileOut.close();
         }catch(IOException e){
             System.out.println(e);
         }
      }
   }
}