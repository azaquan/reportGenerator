package com.ims.app;

import java.lang.Exception;
import java.sql.*;

public class DbUtils{
   public static Connection getConnection(String connectionText, String user, String password){
      Connection conn = null;
      try{
         conn = DriverManager.getConnection(connectionText, user, password);
      }catch(SQLException e){
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
      }
      return result;
   }
}