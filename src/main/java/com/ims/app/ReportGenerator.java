package com.ims.app;

import java.sql.*;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Map;
import java.util.HashMap;
import java.util.Properties;
import java.lang.Integer;
import org.joda.time.DateTime;
import org.joda.time.format.DateTimeFormatter;
import org.joda.time.format.DateTimeFormat;
import org.joda.time.LocalDate;

import org.apache.log4j.Logger;
import org.apache.log4j.PropertyConfigurator;

public class ReportGenerator{
private Connection conn = null;
private String path;
private Properties props = null;
private static final Logger log = Logger.getLogger(ReportGenerator.class);
 
String name="";
String lang="";
String fromDate="";
String toDate="";
int fromDayRef;
int toDayRef;
int fromMonthsRef;
int toMonthsRef;
String namespace="";
String company="";
String location="";
String logical="";
String stock="";
String po="";
String transac="";
	
	public ReportGenerator(String []args){
		for(int i=0;i<args.length;i++){
			try{
				switch(args[i]){
					case "-name":
						name=args[i+1];
						break;
					case "-lang":
						lang=args[i+1];
						break;
					case "-fromDayRef":
						fromDayRef=Integer.parseInt(args[i+1]);
						break;
					case "-toDayRef":
						toDayRef=Integer.parseInt(args[i+1]);
						break;
					case "-fromMonthsRef":
						fromMonthsRef=Integer.parseInt(args[i+1]);
						break;
					case "-toMonthsRef":
						toMonthsRef=Integer.parseInt(args[i+1]);
						break;
					case "-namespace":
						namespace=args[i+1];
						break;
					case "-company":
						company=args[i+1];
						break;
					case "-location":
						location=args[i+1];
						break;
					case "-logical":
						logical=args[i+1];
						break;
					case "-stock":
						stock=args[i+1];
						break;
					case "-po":
						po=args[i+1];
						break;
					case "-transac":
						transac=args[i+1];
						break;
				}
			}catch(Exception e){
				log.debug("Problem while reading arguments: "+e);
			}
		}
	   path = System.getProperty("java.class.path");
	   int lastPoint = path.indexOf("/")+path.indexOf("\\");
	   path = path.substring(0,lastPoint+2);
	   props = ImsUtils.getProperties(path);
	   conn = DbUtils.getConnection(path, props);
		String log4jConfPath = "sources/log4j.properties";
		PropertyConfigurator.configure(log4jConfPath);
	   if(log.isDebugEnabled()){
	   	log.debug("getLocalCurrentDate() has been executed!");
	   }
	}

	public String getPath(){
	      return path;
	}

	public static void main(String []args){
		System.out.println("- -");
		System.out.println("@@@@@ args-> "+args.length);
		ReportGenerator report = new ReportGenerator(args);
		if(args.length>0){
			report.doOnDemand(args);
		}else{
			report.reviewReportList();
		}
	   log.info("ReportGenerator END");
	}
	
	public void doOnDemand(String []args){
		log.info("ReportGenerator ON DEMAN START - - - - - - - - - - - - - - - - - - - - - ");
	   if(conn==null){
	      log.error("Wrong connection with the DB");
	   }else{
         String query = "select * from reportGenerator where reportName='"+name+"' ";
         String reportQuery = "";
         ResultSet result = DbUtils.getResultSet(conn,query);
         if(result==null){
         	log.debug("reviewReportList() has been executed!");                                      
         }else{
            try{
               while(result.next()){
                  if(result.getBoolean("active")){
                     reportQuery=result.getString("query");
                     if(!reportQuery.isEmpty()){
								reportQuery.replace("-namespace",namespace);
								reportQuery.replace("-company",company);
								reportQuery.replace("-location",location);
								reportQuery.replace("-logical",logical);
								reportQuery.replace("-stock",stock);
								reportQuery.replace("-po",po);
								reportQuery.replace("-transac",transac);
                     	log.info("Report: "+name);  
                        generateReport(result,false);
                     }
                  }
               }
            }catch(SQLException e){
            	log.error("Error when iterating with the table");
            	log.error("from server: "+e.getMessage());
            }
         }
      }
	   if(log.isDebugEnabled()){
	   	log.debug("reviewReportList() has been executed!");
	   }
	}

	public void reviewReportList(){
		log.info("ReportGenerator START - - - - - - - - - - - - - - - - - - - - - ");
	   if(conn==null){
	      log.error("Wrong connection with the DB");
	   }else{
         String query = "select * from reportGenerator";
         String reportQuery = "";
         ResultSet result = DbUtils.getResultSet(conn,query);
         if(result==null){
         	log.debug("reviewReportList() has been executed!");                                      
         }else{
            try{
               while(result.next()){
                  if(result.getBoolean("active")){
                     reportQuery=result.getString("query");
                     if(!reportQuery.isEmpty()){
                     	String name=result.getString("reportName");
                     	log.info("Report: "+name);  
                        generateReport(result,true);
                     }
                  }
               }
            }catch(SQLException e){
            	log.error("Error when iterating with the table");
            	log.error("from server: "+e.getMessage());
            }
         }
      }
	   if(log.isDebugEnabled()){
	   	log.debug("reviewReportList() has been executed!");
	   }
	}

	public void generateReport(
			ResultSet repoRef,
			boolean isAutomatic){
	   try{
         String reportQuery=repoRef.getString("query");
         String namespace=repoRef.getString("namespace");
         String reportTemplate=repoRef.getString("template");
         String reportUserId=repoRef.getString("userId");
         String reportRecipients=repoRef.getString("receipients");
         String reportDescription=repoRef.getString("description");
         String reportBodyText1=repoRef.getString("bodyText1");

         if(fromDayRef==0){
         	fromDayRef=repoRef.getInt("fromDayRef");
         }
         if(toDayRef==0){
         	toDayRef=repoRef.getInt("toDayRef");
         }
         if(fromMonthsRef==0){
         	fromMonthsRef=repoRef.getInt("fromMonthRef");
         }
         if(toMonthsRef==0){
         	toMonthsRef=repoRef.getInt("toMonthRef");
         }
         boolean withFromDate=false;
         if (fromDayRef>0){
         	withFromDate=true;
         }
         DateTimeFormatter fmt=DateTimeFormat.forPattern("yyyy-MM-dd");
         boolean goAhead=false;
         if(isAutomatic){
         	goAhead = isValidToGenerate(repoRef);
         }else{
         	goAhead=true;
         }
         if(goAhead){
         	if(withFromDate){
					DateTime startDate=new DateTime().now();
					startDate=startDate.plusMonths(fromMonthsRef);
					startDate=startDate.dayOfMonth().setCopy(fromDayRef);
					DateTime endDate=new DateTime().now();
					endDate=endDate.plusMonths(toMonthsRef);
					if(toDayRef==99){
						endDate=endDate.dayOfMonth().setCopy(endDate.dayOfMonth().getMaximumValue());
					}else{
						endDate=endDate.dayOfMonth().setCopy(toDayRef);
					}
					fromDate = fmt.print(startDate);
					log.debug("- fromDate->"+startDate+" / "+fromDate);
					toDate = fmt.print(endDate);
					log.debug("- fromDate->"+endDate+" / "+toDate);
	
					String newLine = System.getProperty("line.separator");
					reportBodyText1 = reportBodyText1+newLine+newLine+"Description: "+reportDescription;
					reportBodyText1=reportBodyText1+newLine+"From date: "+fromDate+newLine+newLine+"To date: "+toDate;
					reportQuery = reportQuery.replaceAll("\\?1", "'"+fromDate+" 00:00:00.000'");
					reportQuery = reportQuery.replaceAll("\\?2", "'"+toDate+" 23:59:59.999'");
					log.debug("- query->"+reportQuery);
            }
            ResultSet result = DbUtils.getResultSet(conn,reportQuery);
            log.debug("query to process: "+reportQuery);
            if(result==null){
               log.debug("query returns null");
            }else{
               String sufixDate = fmt.print(new DateTime().now());
               String reportFileName = namespace+"-"+name+sufixDate+".xls";
               ExcelPOI excel = new ExcelPOI(path, reportFileName, reportTemplate, props);
               if(excel.createReport(result)){
               	String outboxPath = props.getProperty("path.outbox");
                  String emailSubject= name+" Report ";
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
                     log.info("The report has been routed to be sent succesfully");
                  }else{
                  	log.info("Something went wrong.  The report was not routed to be sent");
                  }
               }else{
                  log.info("Something went wrong.  The "+reportFileName+" could not be generated");
               }
            }
         }
      }catch(IOException | SQLException e){
         log.error("generateReport error: "+e.getMessage());
      }
	   if(log.isDebugEnabled()){
	   	log.debug("generateReport() has been executed!");
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
               	DateTime now = new DateTime();
               	log.debug("@@ today->"+now.dayOfMonth().get());
               	int reportDay=ImsUtils.stringToInt(scheduledMonthDayRef,1);
               	if (reportDay==now.dayOfMonth().get()){
               			isValid=true;
               	}else{
               		log.info("Report active but scheduled for day number: "+reportDay);
               	}
                  break; 	
               default:
            }
         }
      }catch(SQLException e){
         log.error("generateReport error: "+e.getMessage());
      }
	   if(log.isDebugEnabled()){
	   	log.debug("isValidToGenerate() has been executed!");
	   }
      return isValid;
   }
}