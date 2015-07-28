package com.ims.app;

import java.sql.*;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Map;
import java.util.HashMap;
import java.util.Properties;
import java.util.Date;
import java.util.Calendar;
import org.apache.commons.lang3.time.DateUtils;

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
         String reportBodyText1=repoRef.getString("bodyText1");

         int fromDayRef=repoRef.getInt("fromDayRef");
         int toDayRef=repoRef.getInt("toDayRef");
         int fromMonthsRef=repoRef.getInt("fromMonthRef");
         int toMonthsRef=repoRef.getInt("toMonthRef");
         boolean withFromDate=false;
         if (fromDayRef>0){
         	withFromDate=true;
         }

         if(isValidToGenerate(repoRef)){
            System.out.println("fromDayRef->"+fromDayRef);
            Date date=DateUtils.addDays(new Date(), fromDayRef);
            String fromDate=String.format("%1$tY%1$tm%1$te", date);
            fromDate = fromDate + " 00:00:00:001";
            System.out.println("fromDate->"+fromDate);

            System.out.println("toDayRef->"+toDayRef);
            date=DateUtils.addDays(new Date(), toDayRef);
            String toDate=String.format("%1$tY%1$tm%1$te", date);
            toDate = toDate+" 23:59:59:999";

            String newLine = System.getProperty("line.separator");
            reportBodyText1 = reportDescription+newLine+"This is a test."+newLine;
            if (withFromDate){
            	reportBodyText1=reportBodyText1+newLine+"From date: "+fromDate+newLine+newLine+"To date: "+toDate;
            }else{
            	reportBodyText1=reportBodyText1+newLine+"To date: "+toDate;
            }
            System.out.println("toDate->"+toDate);

            reportQuery = reportQuery.replaceFirst("\\?", fromDate);
            reportQuery = reportQuery.replaceFirst("\\?", toDate);
            ResultSet result = DbUtils.getResultSet(conn,reportQuery);

            System.out.println("reportQuery: "+reportQuery);
            if(result==null){
               System.out.println("Nothing got from: "+reportQuery);
            }else{
               String sufixDate = String.format("-%1$tY_%1$tm_%1$te-%1$tH_%1$tM", new Date());
               String reportFileName = reportNamespace+"-"+reportName+sufixDate+".xls";
               ExcelPOI excel = new ExcelPOI(path, reportFileName, reportTemplate, props);
               if(excel.createReport(result)){

                  String outboxPath = props.getProperty("path.outbox");

                  String emailSubject= reportName+" Report ";
                  String emailBody=reportBodyText1;
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
         }
      }catch(IOException | SQLException e){
         System.out.println("generateReport error: "+e.getMessage());
      }
	}

   public static boolean isValidToGenerate(ResultSet record){
      boolean isValid=false;
	   try{
	      String frequency=record.getString("frequency");
	      String scheduledMonthDayRef=record.getString("scheduledMonthDayRef");
	      if(frequency!=null){
            switch(frequency){
               case "daily":
                  isValid=true;
                  break;
               case "monthly":                                   
               	System.out.println("@@ monthly->"+scheduledMonthDayRef);
               	int reportDay=ImsUtils.stringToInt(scheduledMonthDayRef,1);
               	if (reportDay==Calendar.getInstance().getTime().getDay()){
               			isValid=true;
               	}else{
               		System.out.println("Report active but scheduled for day number: "+reportDay);
               	}
                  break; 	
               default:
            }
         }
      }catch(SQLException e){
         System.out.println("generateReport error: "+e.getMessage());
      }
      return isValid;
   }
}