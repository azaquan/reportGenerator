package com.ims.app;

import java.sql.*;
import java.io.IOException;

public class ReportGenerator{
   private Connection conn = null;
   private String path;

	public ReportGenerator(){
	   System.out.println("constructor 0");
	   path = System.getProperty("java.class.path");
	   int lastPoint = path.indexOf("/")+path.indexOf("\\");
	   path = path.substring(0,lastPoint+2);
	   System.out.println("constructor 1");
	}

	public String getPath(){
	      return path;
	}

	public static void main(String []args){
		System.out.println("begining");
		ReportGenerator report = new ReportGenerator();
		System.out.println("@1 path:"+report.getPath());
		report.reviewReportList();
		System.out.println("end");
	}

	public void reviewReportList(){
	   conn = DbUtils.getConnection(path);
	   if(conn==null){
	      System.out.println("Wrong connection with the DB");
	   }else{
         String query = "select * from reportGenerator";
         String reportQuery = "";
         String reportName = "";
         ResultSet result = DbUtils.getResultSet(conn,query);
         if(result==null){
            System.out.println("Nothing was retrieved from the table");
         }else{
            try{
               while(result.next()){
                  reportName=result.getString("reportName");
                  reportQuery=result.getString("query");
                  if(!reportQuery.isEmpty()){
                     generateReport(reportQuery);
                  }
                  break;
               }
            }catch(SQLException e){
               System.out.println("Error when iterating with the table");
               System.out.println("from server: "+e.getMessage());
            }
         }
      }
	}

	public void generateReport(String query){
	   System.out.println("@2 path:"+path);
	   ResultSet result = DbUtils.getResultSet(conn,query);
      if(result==null){
         System.out.println("Nothing got from: "+query);
      }else{
         ExcelPOI excel = new ExcelPOI(path);
         try{
            excel.createReport(result);
            while(result.next()){

            }
         }catch(IOException | SQLException e){
            System.out.println("generateReport error: "+e.getMessage());
         }
      }
	}
}