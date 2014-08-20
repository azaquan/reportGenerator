package com.ims.app;

import java.util.Properties;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Map;
import java.util.HashMap;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.model.Sheet;
import org.apache.poi.hssf.usermodel.*;

public class ImsUtils{
   public static Properties getProperties(String path){
      Properties props = new Properties();
      FileInputStream input = null;
      String fileName = "reportGenerator.properties";
      try{
         input = new FileInputStream(path+fileName);
         if(input==null){
            System.out.println("Sorry, unable to find "+fileName);
            return props;
         }
         props.load(input);
      }catch(IOException e){
         System.out.println("No properties could be loaded:");
         e.printStackTrace();
      }finally{
         if (input != null) {
            try {
               input.close();
            } catch (IOException e) {
               e.printStackTrace();
            }
         }
      }
      return props;
   }

   public static Map<Integer,HSSFCellStyle> getStyleMap(HSSFSheet sheet, int cols){
      HSSFRow row = sheet.getRow(1);
      Map<Integer,HSSFCellStyle> map = new HashMap<Integer,HSSFCellStyle>();
      if(cols>row.getPhysicalNumberOfCells()){
         cols=row.getPhysicalNumberOfCells();
      }
      for(int i=0;i<=cols-1;i++){
         map.put(i,row.getCell(i).getCellStyle());
      }
      return map;
   }

   public static HSSFRow getNewRow(HSSFSheet sheet, Map<Integer,HSSFCellStyle> map, int rowNumber){
      HSSFRow row = sheet.createRow(rowNumber);
      for(int i=0;i<=map.size()-1;i++){
         HSSFCell cell = row.createCell(i);
         cell.setCellStyle(map.get(i));
      }
      return row;
   }
}