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
 
String nullString = null;
String name="";
String lang="";
String fromDate="";
String toDate="";
int fromDayRef;
int toDayRef;
int fromMonthsRef;
int toMonthsRef;
String matrix;
int rowFrom;
int rowTo;
String xuser="";
String namespace="";
String company="";
String location="";
String logical="";
String stock="";
String po="";
String transac="";
String currency="";
boolean sendEmail=true;
	
	public ReportGenerator(String []args){
		for(int i=0;i<args.length;i++){
			try{
				log.debug("--------"+args[i]); 
				switch(args[i]){
					case "-name":
						name=args[i+1];
					case "-sendEmail":
						String sendEmailAnswer =args[i+1];
						if(sendEmailAnswer.toLowerCase().equals("no")){
							sendEmail=false;
						}
						break;
					case "-lang":
						lang=args[i+1];
						break;
					case "-fromDate":
						fromDate=args[i+1];
						break;
					case "-toDate":
						toDate=args[i+1];
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
					case "-xuser":
						xuser=args[i+1];
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
					case "-currency":
						currency=args[i+1];
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
		log.debug("@@@@@ args-> "+args.length);
		ReportGenerator report = new ReportGenerator(args);
		if(args.length>0){
			report.doOnDemand(args);
		}else{
			report.reviewReportList();
		}
		if(log.isInfoEnabled()){
			log.info("ReportGenerator END");
		}
	}
	
	public void doOnDemand(String []args){
		if(log.isInfoEnabled()){
			log.info("ReportGenerator ON DEMAND START - - - - - - - - - - - - - - - - - - - - - ");
		}
	   if(conn==null){
	      log.error("Wrong connection with the DB");
	   }else{
         String query = "select * from reportGenerator where reportName='"+name+"' ";
         String reportQuery = "";
         ResultSet result = DbUtils.getResultSet(conn,query);
         if(result==null){
         	log.debug("doOnDemand() has been executed with no results from reportGenerator table!");                                      
         }else{
            try{
               while(result.next()){
                  if(result.getBoolean("active")){
                     reportQuery=result.getString("query");
							if(log.isDebugEnabled()){ 
								log.info("-Raw reportQuery->"+reportQuery); 
							}
                     if(!reportQuery.isEmpty()){
								reportQuery=reportQuery.replace("-namespace",namespace);
								reportQuery=reportQuery.replace("-company",company);
								reportQuery=reportQuery.replace("-location",location);
								reportQuery=reportQuery.replace("-logical",logical);
								reportQuery=reportQuery.replace("-stock",stock);
								reportQuery=reportQuery.replace("-po",po);
								reportQuery=reportQuery.replace("-transac",transac);
								reportQuery=reportQuery.replace("-currency",currency);
								reportQuery=reportQuery.replace("-fromDate",fromDate);
								reportQuery=reportQuery.replace("-toDate",toDate);
								if(log.isInfoEnabled()){
									log.info("-Prepared reportQuery->"+reportQuery);
								}
                     	String email = DbUtils.getUserEmail(conn, namespace, xuser);
                        generateReport(result,false,reportQuery, email);
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
		if(log.isInfoEnabled()){
			log.info("ReportGenerator START - - - - - - - - - - - - - - - - - - - - - ");
		}
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
                     	name=result.getString("reportName");
                     	if(log.isInfoEnabled()){
                     		log.info("Report: "+name);  
                     	}
                        generateReport(result,true,"","");
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
			boolean isAutomatic,
			String reportQuery,
			String email){
	   try{
	   	if(reportQuery.equals("")){
	   		System.out.println("@@@ after reportQuery->NULL");
	   		reportQuery=repoRef.getString("query");
	   	}
	   	System.out.println("@@@ reportQuery->"+reportQuery);
         String reportTemplate=repoRef.getString("template");
         String reportUserId=repoRef.getString("userId");
         String reportRecipients=repoRef.getString("receipients");
         if (!email.equals("")){
         		reportRecipients=email;
         }
         String reportDescription=repoRef.getString("description");
         String reportTitle=repoRef.getString("title");
         String reportBodyText1=repoRef.getString("bodyText1");
			matrix = repoRef.getString("matrix");
			if(repoRef.wasNull()){
				matrix = "";
			}
			rowFrom = repoRef.getInt("rowFrom");
			rowTo = repoRef.getInt("rowTo");
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
					reportQuery = reportQuery=reportQuery.replaceAll("\\?1", "'"+fromDate+" 00:00:00.000'");
					reportQuery = reportQuery=reportQuery.replaceAll("\\?2", "'"+toDate+" 23:59:59.999'");
					log.debug("- query->"+reportQuery);
            }
            ResultSet result = DbUtils.getResultSet(conn,reportQuery);
            log.debug(".  .  .  .  .  .  .  .  .  .  .query to process -> "+reportQuery);
            if(result==null){
               log.debug("query returns null");
            }else{
            	DateTimeFormatter fmtTime=DateTimeFormat.forPattern("yyyy-MM-dd_HH-mm-ss");
               String sufixDate = fmtTime.print(new DateTime().now());
               String sql = "select npce_name from namespace where npce_code='"+namespace+"'";
               String namespaceName=name;
               ResultSet fromNamespace = DbUtils.getResultSet(conn,sql);
               if(fromNamespace==null){
               	log.debug("namespace query returns null");
               }else{
               	try{
               		while(fromNamespace.next()){
               			namespaceName=fromNamespace.getString("npce_name");
               		}
						}catch(SQLException e){
							log.error("from server: "+e.getMessage());
						}
						namespaceName=namespaceName.replace(" ","_");
						String reportFileName =
							name+"-"+
							(namespaceName.equals("")?"":namespaceName+"-")+
							(company.equals("")?"":company+"-")+
							(location.equals("")?"":location+"-")+sufixDate+".xls";
						String title =
							(reportTitle.equals("")?"":reportTitle+" ")+
							(namespaceName.equals("")?"":namespaceName+" ");
						String period=
							(fromDate.equals("")?"":fromDate+" to ")+
							(toDate.equals("")?"":toDate);
						ExcelPOI excel = new ExcelPOI(path, reportFileName, reportTemplate, props, title, period, matrix, rowFrom, rowTo);
						if(excel.createReport(result)){
							if(sendEmail){
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
									if(log.isInfoEnabled()){
										log.info("The report has been routed to be sent succesfully");
									}
								}else{
									if(log.isInfoEnabled()){
										log.info("Something went wrong.  The report was not routed to be sent");
									}
								}
							}
						}else{
							if(log.isInfoEnabled()){
								log.info("Something went wrong.  The "+reportFileName+" could not be generated");
							}
						}
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
               		if(log.isInfoEnabled()){
               			log.info("Report active but scheduled for day number: "+reportDay);
               		}
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