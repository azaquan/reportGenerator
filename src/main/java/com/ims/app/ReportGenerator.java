package com.ims.app;

import java.sql.*;
import java.io.IOException;

public class ReportGenerator{
   private Connection conn = null;
	public void ReportGenerator(){
	}

	public static void main(String []args){
		System.out.println("begining");
		ReportGenerator report = new ReportGenerator();
		report.reviewReportList();
		System.out.println("end");
	}


	public void reviewReportList(){
	   System.out.println("reviewReportList method");
	   conn = DbUtils.getConnection("jdbc:sqlserver://192.168.1.11:1433;databaseName=exxonTest;", "sa", "juancarl0s");
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
	   ResultSet result = DbUtils.getResultSet(conn,query);
      if(result==null){
         System.out.println("Nothing got from: "+query);
      }else{
         ExcelPOI excel = new ExcelPOI();
         try{
            excel.createReport(result);
            while(result.next()){

            }
         }catch(SQLException e){
            System.out.println("generateReport error: "+e.getMessage());
         }catch(IOException e){
            System.out.println("generateReport error: "+e.getMessage());
         }
      }
	}
}