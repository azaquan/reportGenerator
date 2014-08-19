package com.ims.app;

import java.sql.*;
import java.io.IOException;
import java.util.Map;
import java.util.HashMap;
import java.util.Properties;
import java.util.Date;

public class ReportGenerator{
   private Connection conn = null;
   private String path;
   private Properties props = null;

	public ReportGenerator(){
	   path = System.getProperty("java.class.path");
	   int lastPoint = path.indexOf("/")+path.indexOf("\\");
	   path = path.substring(0,lastPoint+2);
	   props = ImsUtils.getProperties(path);
	   conn = DbUtils.getConnection(path, props);
	}

	public String getPath(){
	      return path;
	}

	public static void main(String []args){
		System.out.println("begining");
		ReportGenerator report = new ReportGenerator();
		report.reviewReportList();
		System.out.println("end");
	}

	public void reviewReportList(){
	   if(conn==null){
	      System.out.println("Wrong connection with the DB");
	   }else{
         String query = "select * from reportGenerator";
         String reportQuery = "";
         ResultSet result = DbUtils.getResultSet(conn,query);
         if(result==null){
            System.out.println("Nothing was retrieved from the table");
         }else{
            try{
               while(result.next()){
                  if(result.getBoolean("active")){
                     reportQuery=result.getString("query");
                     if(!reportQuery.isEmpty()){
                        generateReport(result);
                     }
                  }
               }
            }catch(SQLException e){
               System.out.println("Error when iterating with the table");
               System.out.println("from server: "+e.getMessage());
            }
         }
      }
	}

	public void generateReport(ResultSet repoRef){
	   try{
         String reportQuery=repoRef.getString("query");
         String reportName=repoRef.getString("reportName");
         String reportNamespace=repoRef.getString("namespace");
         String reportTemplate=repoRef.getString("template");
         String reportUserId=repoRef.getString("userId");
         String reportRecipients=repoRef.getString("receipients");
         String reportDescription=repoRef.getString("description");

         ResultSet result = DbUtils.getResultSet(conn,reportQuery);
         if(result==null){
            System.out.println("Nothing got from: "+reportQuery);
         }else{
            String sufixDate = String.format("-%1$tY_%1$tm_%1$te-%1$tH_%1$tM", new Date());
            String reportFileName = reportNamespace+"-"+reportName+sufixDate+".xls";
            ExcelPOI excel = new ExcelPOI(path, reportFileName, reportTemplate, props);
            if(excel.createReport(result)){

               String outboxPath = props.getProperty("path.outbox");

               String emailSubject= reportName+" Report ";
               String emailBody="This is a testing email for the automation of the reports";  //It should be taken from properties
               String emailAttachmentFile=outboxPath+reportFileName;
               String emailRecipientStr=reportRecipients;
               String emailCreaUser=reportUserId;

               Map<Integer, String> map = new HashMap<Integer, String>();
               map.put(1, emailSubject);
               map.put(2, emailBody);
               map.put(3, emailAttachmentFile);
               map.put(4, emailRecipientStr);
               map.put(5, emailCreaUser);
               if(DbUtils.sendReport(conn,map)){
                  System.out.println("The report has been routed to be sent succesfully");
               }else{
                  System.out.println("Something went wrong.  The report was not routed to be sent");
               }
            }else{
               System.out.println("Something went wrong.  The "+reportFileName+" could not be generated");
            }
         }
      }catch(IOException | SQLException e){
         System.out.println("generateReport error: "+e.getMessage());
      }
	}
}