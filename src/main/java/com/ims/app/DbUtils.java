package com.ims.app;

import java.lang.Exception;
import java.sql.*;
import java.util.Properties;
import java.io.FileInputStream;
import java.io.IOException;

public class DbUtils{
   public static Connection getConnection(String path){
      Properties props = new Properties();
      Connection conn = null;
      try{
         FileInputStream fis = new FileInputStream(path+"jdbc.properties");
         props.load(fis);
         if(props==null){
            System.out.println("No properties found");
         }else{
            conn = DriverManager.getConnection(
               "jdbc:sqlserver:"+props.getProperty("jdbc.server")+";databaseName="+props.getProperty("jdbc.database"),
               props.getProperty("jdbc.user"),
               props.getProperty("jdbc.password")
            );
         }
      }catch(IOException | SQLException e){
         System.out.println("from the server: "+e.getMessage());
      }
      return conn;
   }

   public static ResultSet getResultSet(Connection conn, String query){
      ResultSet result = null;
      try{
         Statement stmt = conn.createStatement();
         result = stmt.executeQuery(query);
      }catch(SQLException e){
         System.out.println("@getResultSet: "+e.getMessage());
      }
      return result;
   }
}