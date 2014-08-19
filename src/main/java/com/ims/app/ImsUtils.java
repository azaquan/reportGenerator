package com.ims.app;

import java.util.Properties;
import java.io.FileInputStream;
import java.io.IOException;


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
}