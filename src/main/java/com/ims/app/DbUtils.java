package com.ims.app;

import java.lang.Exception;
import java.sql.*;
import java.sql.CallableStatement;
import java.util.Properties;
import java.util.Map;
import java.io.FileInputStream;
import java.io.IOException;


public class DbUtils{
   public static Connection getConnection(String path, Properties props){
      DbUtils dbUtils = new DbUtils();
      Connection conn = null;
      try{
         if(props==null){
            System.out.println("No properties found");
         }else{
            conn = DriverManager.getConnection(
               "jdbc:sqlserver:"+props.getProperty("jdbc.server")+";databaseName="+props.getProperty("jdbc.database"),
               props.getProperty("jdbc.user"),
               props.getProperty("jdbc.password")
            );
         }
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
         System.out.println("@getResultSet: "+e.getMessage());
      }
      return result;
   }


   public static boolean sendReport(Connection dbConnection, Map<Integer,String> param){
      boolean sent = false;
      String insertStoreProc = "{call InsertEmailFax(?,?,?,?,?)}";
      CallableStatement callableStatement = null;
 		try {
         callableStatement = dbConnection.prepareCall(insertStoreProc);
         for(int i=1;i<param.size()+1;i++){
            callableStatement.setString(i, param.get(i));
         }
         callableStatement.executeUpdate();
         sent = true;
		} catch (SQLException e) {
			System.out.println(e.getMessage());
		} finally {
		   try{
		      callableStatement.close();
         } catch (SQLException e) {
            System.out.println(e.getMessage());
         }
		}
		return sent;
   }

   public static boolean isValidToGenerate(ResultSet record){
      boolean isValid=false;
      return isValid;
   }
}