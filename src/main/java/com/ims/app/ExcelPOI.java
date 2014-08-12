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
   public String path;


   public ExcelPOI(String p){
      path=p;
   }

   public void createReport(ResultSet r) throws IOException {
      result=r;
      int j=0;
      File file = null;
      InputStream fileinput = null;
      HSSFWorkbook workbook = null;
      HSSFSheet sheet = null;
      try{
         file = new File(path+"templates/sampleExcel.xls");
         fileinput= new FileInputStream(file);

         if(fileinput!=null){
            workbook = (HSSFWorkbook) WorkbookFactory.create(fileinput);
            sheet = workbook.getSheetAt(0);
            while(result.next()){
               HSSFRow row = sheet.createRow(j+1);
               for(int i=0;i<=result.getMetaData().getColumnCount()-1;i++){
                  HSSFCell cell = row.createCell(i);
                  cell.setCellValue(result.getString(i+1));
               }
               j=j+1;
            }
            FileOutputStream fileOut = new FileOutputStream(path+"templates/incident_result.xls");
            workbook.write(fileOut);
            fileOut.close();
         }
      }catch(Exception e){
          System.out.println(e);
      }
   }
}