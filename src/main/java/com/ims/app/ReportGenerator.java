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
    String fromDateQuery="";
    String toDate="";
    String toDateQuery="";
    DateTime repoDate=new DateTime().now();  
    int fromDayRef;
    int toDayRef;
    int fromMonthsRef;
    int toMonthRef;
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
    String tempEmail="";
    boolean sendEmail=true;
    String reportQuery;
	
	public ReportGenerator(String []args){
		for(int i=0;i<args.length;i++){
			try{
				log.debug("- - arguments --------"+args[i]); 
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
						fromDateQuery = fromDate;
						break;
					case "-toDate":
						toDate=args[i+1];
						toDateQuery = toDate;
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
					case "-toMonthRef":
						toMonthRef=Integer.parseInt(args[i+1]);
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
					case "-tempEmail":
						tempEmail=args[i+1];
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
             reportQuery = "";
             ResultSet result = DbUtils.getResultSet(conn,query);
             if(result==null){
                log.debug("doOnDemand() has been executed with no results from reportGenerator table!");                                      
             }else{
                 try{
                     while(result.next()){
                         if(result.getBoolean("active")){
                             reportQuery=result.getString("query");
                             if(log.isDebugEnabled()){
                                 log.debug("-Raw reportQuery->"+reportQuery); 
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
                                 if (!fromDate.equals("")){
                                     reportQuery=reportQuery.replace("-fromDate",fromDate+" 00:00:00.000");
                                     reportQuery=reportQuery.replace("-toDate",toDate+" 23:59:59.999");
								 }
								 String email="";
								 if (tempEmail.equals("")){
								     email = DbUtils.getUserEmail(conn, namespace, xuser);
								 }else{
								     email=tempEmail;
								 }
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
                            generateReport(result,true,reportQuery,"");
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
        boolean goAhead=false;
        if(isAutomatic){
         	goAhead = isValidToGenerate(repoRef);
        }else{
         	goAhead=true;
        }
        if(goAhead){
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
                log.debug("--- - fromDayRef->"+fromDayRef);
            }
            if(toDayRef==0){
                toDayRef=repoRef.getInt("toDayRef");
                log.debug("--- - toDayRef->"+toDayRef);
            }
            if(fromMonthsRef==0){
                fromMonthsRef=repoRef.getInt("fromMonthRef");
                log.debug("--- - fromMonthsRef->"+fromMonthsRef);
            }
            if(toMonthRef==0){
                toMonthRef=repoRef.getInt("toMonthRef");
                log.debug("--- - toMonthRef->"+toMonthRef);
            }
            boolean withFromDate=false;
            if (fromDayRef>0){
                withFromDate=true;
            }
            DateTimeFormatter fmt=DateTimeFormat.forPattern("yyyy-MM-dd");
            if(withFromDate){
                DateTime startDate=new DateTime().now();
                log.info("--- - startDate (now)->"+startDate);
                startDate=startDate.plusMonths(fromMonthsRef);
                log.info("--- - startDate (month)->"+startDate);
                startDate=startDate.dayOfMonth().setCopy(fromDayRef);
                log.info("--- - startDate- (day)>"+startDate);
                DateTime endDate=new DateTime().now();  
                log.info("--- - endDate (now)>"+endDate);
                endDate=endDate.plusMonths(toMonthRef);
                log.info("--- - endDate (month)>"+endDate);
                if(toDayRef==99){
                    endDate=endDate.dayOfMonth().setCopy(endDate.dayOfMonth().getMaximumValue());
                }else{
                    endDate=endDate.dayOfMonth().setCopy(toDayRef);
                }
                log.info("--- - endDate (day)>"+endDate);
                fromDate = fmt.print(startDate);
                fromDate = fromDate+" 00:00:00.000";
                log.info("- fromDate->"+fromDate);
                toDate = fmt.print(endDate);
                toDate = toDate+" 23:59:59.999";
                log.info("- toDate->"+toDate);

                log.info("- before all->"+reportQuery);
                reportQuery = reportQuery.replace("?1", fromDate);
                reportQuery = reportQuery.replace("-fromDate",fromDate);
                reportQuery = reportQuery.replace("?2", toDate);
                reportQuery = reportQuery.replace("-toDate",toDate);
                log.info("- after all->"+reportQuery);
                String newLine = System.getProperty("line.separator");
                reportBodyText1 = reportBodyText1+newLine+newLine+"Description: "+reportDescription;
                if(!fromDateQuery.isEmpty()){
                    reportBodyText1=reportBodyText1+newLine+"From date: "+fromDate+newLine+"To date: "+toDate;
                }
            }
            ResultSet result = DbUtils.getResultSet(conn,reportQuery);
            log.info(".  .  .  .  .  .  .  .  .  .  . final query to get -> "+reportQuery);
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
                    String period="";
                        if(!fromDateQuery.isEmpty()){
                                period = (fromDate.equals("")?"":fromDate+" to ");
                        }
                        if(toDateQuery.isEmpty()){
                            toDate = fmt.print(repoDate);
                        }
                        period = period + (toDate.equals("")?"":toDate);
                    ExcelPOI excel = new ExcelPOI(path, reportFileName, reportTemplate, props, title, period, matrix, rowFrom, rowTo);
                    if(excel.createReport(result)){
                        try{
                            Thread.sleep(3000);
                        }catch(InterruptedException ex){
                            Thread.currentThread().interrupt();
                        }
                        if(sendEmail){
                            String outboxPath = props.getProperty("path.outbox");
                            String emailSubject= name+" Report ";
                            String emailBody=reportBodyText1;
                            String emailAttachmentFile=outboxPath+reportFileName;
                            String emailRecipientStr=reportRecipients;
                            String emailCreaUser=reportUserId;
    
                            Map<String, String> map = new HashMap<String, String>();
                            map.put("emailSubject", emailSubject + "("+reportFileName+")");
                            map.put("emailBody", emailBody);
                            map.put("emailAttachmentFile", emailAttachmentFile);
                            map.put("emailRecipientStr", emailRecipientStr);
                            map.put("emailCreaUser", emailCreaUser);
                            if(ImsUtils.sendEmail(conn, props, map)){
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
            boolean isMonthly = false;
            boolean isWeekly = false;
            boolean isDaily = false;
            String scheduledMonthDayRef=record.getString("scheduledMonthDayRef");
            String scheduledWeekDayRef =record.getString("scheduledWeekDayRef");
			DateTime now = new DateTime();
			log.info("@@ today->"+now.dayOfMonth().get());
			if(frequency!=null){
				if (frequency.indexOf("daily")>=0){
					isDaily = true;
					isValid=true;
				}
				if (frequency.indexOf("weekly")>=0){
					isWeekly = true;
					int weekReportDay=ImsUtils.stringToInt(scheduledWeekDayRef,1);
					if (weekReportDay==now.dayOfWeek().get()){
						isValid=true;
					}else{
						if(log.isInfoEnabled()){
							log.info("Report weekly active but scheduled for day number: "+weekReportDay);
						}
					}
				}       
				if (frequency.indexOf("monthly")>=0){
					isMonthly = true;
					int monthReportDay=ImsUtils.stringToInt(scheduledMonthDayRef,1);
					if (monthReportDay==now.dayOfMonth().get()){
							isValid=true;
					}else{
						if(log.isInfoEnabled()){
							log.info("Report monthly active but scheduled for day number: "+monthReportDay);
						}
					}
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