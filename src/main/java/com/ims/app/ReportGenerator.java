package com.ims.app;

import java.sql.*;
import java.io.IOException;
import java.io.FileInputStream;
import java.text.SimpleDateFormat;
import java.util.Map;
import java.util.HashMap;
import java.util.Properties;
import java.lang.Integer;
import org.joda.time.DateTime;
import org.joda.time.format.DateTimeFormatter;
import org.joda.time.format.DateTimeFormat;
import org.joda.time.LocalDate;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

public class ReportGenerator{
    private Connection conn = null;
    private String path;
    private Properties props = null;
    private static Logger LOGGER = null;
    
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
    int weekReportDay;
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
	boolean isMonthly = false;
	boolean isWeekly = false;
	boolean isDaily = false;
	DateTime now = new DateTime();
	DateTimeFormatter fmtTime=DateTimeFormat.forPattern("yyyy-MM-dd_HH-mm-ss");
	DateTimeFormatter fmt=DateTimeFormat.forPattern("yyyy-MM-dd");	
	
	public ReportGenerator(String []args){
        LOGGER.info("Into constructor");	
        path = System.getProperty("java.class.path");
        int lastPoint = path.indexOf("/")+path.indexOf("\\");
        LOGGER.info("_____________00 path:{}",path);	
        path = path.substring(0,lastPoint+2);
        props = ImsUtils.getProperties(path);
        conn = DbUtils.getConnection(path, props);
        
		for(int i=0;i<args.length;i++){
			try{
				if(LOGGER.isDebugEnabled()){
					LOGGER.debug("- - arguments --------{}"+args[i]); 
				}
				switch(args[i]){
					case "-name":
						name=args[i+1];
						LOGGER.debug("@@@@@ args-> {}",args.length);
					case "-sendEmail":
						String sendEmailAnswer =args[i+1];
						if(sendEmailAnswer.toLowerCase().equals("no")){
						    sendEmail=false;
						}
						LOGGER.debug("@@@@@     ->sendEmail {}",sendEmailAnswer);
						break;
					case "-lang":
						lang=args[i+1];
						LOGGER.debug("@@@@@     ->lang {}",lang);
						break;
					case "-fromDate":
						fromDate=args[i+1];
						fromDateQuery = fromDate;
						LOGGER.debug("@@@@@     ->fromDate {}",fromDate);
						break;
					case "-toDate":
						toDate=args[i+1];
						toDateQuery = toDate;
						LOGGER.debug("@@@@@     ->toDate {}",toDate);
						break;
					case "-fromDayRef":
						fromDayRef=Integer.parseInt(args[i+1]);
						LOGGER.debug("@@@@@     ->fromDayRef {}",fromDayRef);
						break;
					case "-toDayRef":
						toDayRef=Integer.parseInt(args[i+1]);
						LOGGER.debug("@@@@@     ->toDayRef {}",toDayRef);
						break;
					case "-fromMonthsRef":
						fromMonthsRef=Integer.parseInt(args[i+1]);
						LOGGER.debug("@@@@@     ->fromMonthsRef {}",fromMonthsRef);
						break;
					case "-toMonthRef":
						toMonthRef=Integer.parseInt(args[i+1]);
						LOGGER.debug("@@@@@     ->toMonthRef {}",toMonthRef);
						break;
					case "-xuser":
						xuser=args[i+1];
						LOGGER.debug("@@@@@     ->xuser {}",xuser);
						break;
					case "-namespace":
						namespace=args[i+1];
						LOGGER.debug("@@@@@     ->namespace {}",namespace);
						break;
					case "-company":
						LOGGER.debug("@@@@@     ->company {}",company);
						company=args[i+1];
						break;
					case "-location":
						LOGGER.debug("@@@@@     ->location {}",location);
						location=args[i+1];
						break;
					case "-logical":
						LOGGER.debug("@@@@@     ->logical {}",logical);
						logical=args[i+1];
						break;
					case "-stock":
						LOGGER.debug("@@@@@     ->stock {}",stock);
						stock=args[i+1];
						break;
					case "-po":
						LOGGER.debug("@@@@@     ->po {}",po);
						po=args[i+1];
						break;
					case "-transac":
						LOGGER.debug("@@@@@     ->transac {}",transac);
						transac=args[i+1];
						break;
					case "-currency":
						LOGGER.debug("@@@@@     ->currency {}",currency);
						currency=args[i+1];
						break;
					case "-tempEmail":
						LOGGER.debug("@@@@@     ->tempEmail {}",tempEmail);
						tempEmail=args[i+1];
						break;
				}
			}catch(Exception e){
				LOGGER.error("Problem while reading arguments: {}", e.getMessage());
			}
		}
        if(LOGGER.isDebugEnabled()){
            LOGGER.debug("getLocalCurrentDate() has been executed!");
        }
	}

	public String getPath(){
	    return path;
	}

	public static void main(String []args){
		System.setProperty("log4j.configurationFile",  "resources/log4j2.xml");
		LOGGER = LogManager.getLogger(ReportGenerator.class);
		LOGGER.info("-------------------------------------");
		LOGGER.info("Starting the application");
		if(LOGGER.isDebugEnabled()){
			LOGGER.debug("@@@@@ args-> {}",args.length);
		}
		ReportGenerator report = new ReportGenerator(args);
		if(args.length>0){
			report.doOnDemand(args);
		}else{
			report.reviewReportList();
		}
		LOGGER.info("ReportGenerator END");
	}
	
	public void doOnDemand(String []args){
        LOGGER.info("ReportGenerator ON DEMAND START - - - - - - - - - - - - - - - - - - - - - ");
        if(conn==null){
            LOGGER.error("Wrong connection with the DB");
        }else{
             String query = "select * from reportGenerator where reportName='"+name+"' ";
             reportQuery = "";
             ResultSet result = DbUtils.getResultSet(conn,query);
             if(result==null){
                LOGGER.debug("doOnDemand() has been executed with no results from reportGenerator table!");     
             }else{
                 try{
                     while(result.next()){
                         if(result.getBoolean("active")){
                             reportQuery=result.getString("query");
                             if(LOGGER.isDebugEnabled()){
                                 LOGGER.debug("-Raw reportQuery->"+reportQuery); 
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
				     LOGGER.error("Error when iterating with the table");
				     LOGGER.error("from server: {}", e.getMessage());
				 }
			 }
		}
		LOGGER.debug("doOnDemand() has been executed!");

	}

	public void reviewReportList(){
		LOGGER.info("ReportGenerator START - - - - - - - - - - - - - - - - - - - - - ");
        if(conn==null){
            LOGGER.error("Wrong connection with the DB");
        }else{
        String query = "select * from reportGenerator";
        ResultSet result = DbUtils.getResultSet(conn,query);
        if(result!=null){
            try{
                while(result.next()){
                    if(result.getBoolean("active")){
                        reportQuery=result.getString("query");
                        if(!reportQuery.isEmpty()){
                            name=result.getString("reportName");
                            LOGGER.info(".");
                            LOGGER.info("Report:{}", name);  
                            generateReport(result,true,reportQuery,"");
                        }
                    }
                }
            }catch(SQLException e){
            	LOGGER.error("Error when iterating with the table");
            	LOGGER.error("from server: ", e.getMessage());
            }
        }
    }
        LOGGER.debug("reviewReportList() has been executed!");
	}

	public void generateReport(
        ResultSet repoRef,
        boolean isAutomatic,
        String reportQuery,
        String email){
		LOGGER.debug("@@ today->{}",now.dayOfMonth().get());
        try{
            if(reportQuery.equals("")){
                reportQuery=repoRef.getString("query");
            }
        boolean goAhead=true;
        boolean isValid=isValidToGenerate(repoRef);
        if(isAutomatic){
         	goAhead = isValid;
        }
        if(goAhead){
        	LOGGER.debug("goAhead=true");
            String reportTemplate=repoRef.getString("template");
            LOGGER.info("--- - template={}", reportTemplate);
            String reportUserId=repoRef.getString("userId");
            String reportRecipients=repoRef.getString("receipients");
            if (!email.equals("")){
                reportRecipients=email;
            }
            String reportDescription=repoRef.getString("description");
            String reportTitle=repoRef.getString("title");
            LOGGER.info("--- - title={}", reportTitle);
            String reportBodyText1=repoRef.getString("bodyText1");
            matrix = repoRef.getString("matrix");
            if(repoRef.wasNull()){
                matrix = "";
            }
            LOGGER.info("--- - matrix={}", matrix);            
			rowFrom = repoRef.getInt("rowFrom");
			LOGGER.debug("--- - rowFrom->{}", rowFrom);
			rowTo = repoRef.getInt("rowTo");
			LOGGER.debug("--- - rowTo->{}", rowTo);
			if(isAutomatic){
				if(fromDayRef==0){
					fromDayRef=repoRef.getInt("fromDayRef");
					if(LOGGER.isDebugEnabled()){
						LOGGER.debug("--- - fromDayRef->", Integer.toString(fromDayRef));
					}
				}
				if(toDayRef==0){
					toDayRef=repoRef.getInt("toDayRef");
					if(LOGGER.isDebugEnabled()){
						LOGGER.debug("--- - toDayRef->", Integer.toString(toDayRef));
					}
				}
				if(fromMonthsRef==0){
					fromMonthsRef=repoRef.getInt("fromMonthRef");
					if(LOGGER.isDebugEnabled()){
						LOGGER.debug("--- - fromMonthsRef->", Integer.toString(fromMonthsRef));
					}
				}
				if(toMonthRef==0){
					toMonthRef=repoRef.getInt("toMonthRef");
					if(LOGGER.isDebugEnabled()){
						LOGGER.debug("--- - toMonthRef->", Integer.toString(toMonthRef));
					}
				}
				boolean withFromDate=false;
				if (fromDayRef>0){
					withFromDate=true;
					if(LOGGER.isDebugEnabled()){
						LOGGER.debug("--- - withFromDate->true");
					}
				}
				
				if(withFromDate){
					DateTime nowDate=new DateTime().now();
					DateTime startDate=new DateTime().now();
					if(LOGGER.isDebugEnabled()){
						LOGGER.debug("--- - nowDate (now)->{}",nowDate);
						LOGGER.debug("--- - isWeekly->{}",isWeekly);
						LOGGER.debug("--- - isMonthly-{}>",isMonthly);
					}
					if(isWeekly){
						startDate=nowDate.minusDays(7);
						if(LOGGER.isDebugEnabled()){
							LOGGER.debug("--- - startDate- (weekly)>{}", startDate);
						}
					}
					if(isMonthly){
						LOGGER.debug("--- - startDate- pre-calculation: [}", startDate);
						startDate=nowDate.plusMonths(fromMonthsRef);
						LOGGER.debug("--- - startDate- after-adding fromMonthsRef: {}", startDate);
						startDate=startDate.dayOfMonth().setCopy(fromDayRef);
						LOGGER.debug("--- - startDate- after-calculations>{}: {}", startDate);
					}
					
					DateTime endDate=new DateTime().now();  
					if(LOGGER.isDebugEnabled()){
						LOGGER.debug("--- - endDate (now)>{}", endDate);
					}
					endDate=endDate.plusMonths(toMonthRef);
					if(LOGGER.isDebugEnabled()){
						LOGGER.debug("--- - endDate (month)>{}", endDate);
					}                
					if(toDayRef==99){
						endDate=endDate.dayOfMonth().setCopy(endDate.dayOfMonth().getMaximumValue());
					}else{
						endDate=endDate.dayOfMonth().setCopy(toDayRef);
					}
					if(LOGGER.isDebugEnabled()){
						LOGGER.debug("--- - endDate (day) ->{}", endDate);
					}                  
	
					fromDate = fmt.print(startDate);
					fromDateQuery = fromDate;
					fromDate = fromDate+" 00:00:00.000";
					
					if(LOGGER.isDebugEnabled()){
						LOGGER.debug("--- - fromDate ->{}", fromDate);
					}                     
					toDate = fmt.print(endDate);
					toDate = toDate+" 23:59:59.999";
					if(LOGGER.isDebugEnabled()){
						LOGGER.debug("--- - toDate ->{}", toDate);
						LOGGER.debug("--- - before all ->{}", reportQuery);
					}                 
	
					reportQuery = reportQuery.replace("?1", fromDate);
					reportQuery = reportQuery.replace("-fromDate",fromDate);
					reportQuery = reportQuery.replace("?2", toDate);
					reportQuery = reportQuery.replace("-toDate",toDate);
					if(LOGGER.isDebugEnabled()){
						LOGGER.debug("________________________");
						LOGGER.debug("--- - after all ->{}", reportQuery);
					}                  
				}
			}
			String newLine = System.getProperty("line.separator");
			reportBodyText1 = reportBodyText1+newLine+newLine+"Description: "+reportDescription;
			if(!fromDateQuery.isEmpty()){
				reportBodyText1=reportBodyText1+newLine+"From date: "+fromDate+newLine+"To date: "+toDate;
			}			
            ResultSet result = DbUtils.getResultSet(conn,reportQuery);
			if(LOGGER.isDebugEnabled()){
				LOGGER.debug("________________________");
				LOGGER.debug("--- - final query to retrieve ->{}", reportQuery);
			}              
            if(result==null){
               LOGGER.info("query returns null");
            }else{                
                String sufixDate = fmtTime.print(new DateTime().now());
                String sql = "select npce_name from namespace where npce_code='"+namespace+"'";
                String namespaceName=name;
                ResultSet fromNamespace = DbUtils.getResultSet(conn,sql);
                if(fromNamespace==null){
                    LOGGER.info("namespace query returns null");
                }else{
                    try{
                        while(fromNamespace.next()){
                            namespaceName=fromNamespace.getString("npce_name");
                        }
                    }catch(SQLException e){
                        LOGGER.error("from server: {}", e.getMessage());
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
                    LOGGER.info("...............title:{}",title);
                    String period="";
                        if(fromDateQuery.isEmpty()){
                        	period = "xx-xx-xx to";
                        }else{
                            period = (fromDate.equals("")?"":fromDate+" to ");
                        }
                        if(toDateQuery.isEmpty()){
                            toDate = fmt.print(repoDate);
                        }
                        period = period + (toDate.equals("")?"":toDate);
                    LOGGER.info("...............period:{}",period);
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
                            	LOGGER.info("The report has been routed to be sent succesfully");
                            }else{
                                LOGGER.info("Something went wrong.  The report was not routed to be sent");
                            }
                        }
                    }else{
                    	LOGGER.info("Something went wrong.  The {} could not be generated", reportFileName);
                    }
                }
            }
        }
        }catch(IOException | SQLException e){
            LOGGER.error("generateReport error: {}",e.getMessage());
        }
        LOGGER.info("generateReport() has been completed!");
	}

	public boolean isValidToGenerate(ResultSet record){
	    boolean isValid=false;
	    try{
            String frequency=record.getString("frequency");
            LOGGER.debug("...frequency: {}",frequency);
            String scheduledMonthDayRef=record.getString("scheduledMonthDayRef");
            String scheduledWeekDayRef =record.getString("scheduledWeekDayRef");
			isMonthly = false;
			isWeekly = false;
			isDaily = false;
			if(frequency!=null){
				if (frequency.contains("daily")){
					isDaily = true;
					isValid=true;
				}
				if (frequency.contains("weekly")){
					isWeekly = true;
					weekReportDay=ImsUtils.stringToInt(scheduledWeekDayRef,1);
					if (weekReportDay==now.dayOfWeek().get()){
						isValid=true;
					}else{
						LOGGER.debug("Report weekly active but scheduled for day number: {}",weekReportDay);
					}
				}       
				if (frequency.contains("monthly")){
					isMonthly = true;
					int monthReportDay=ImsUtils.stringToInt(scheduledMonthDayRef,1);
					if (monthReportDay==now.dayOfMonth().get()){
							isValid=true;
					}else{
						LOGGER.debug("Report monthly active but scheduled for day number: {}", monthReportDay);
					}
				}
			}
			LOGGER.debug("@@ frequency->{}",frequency);
			LOGGER.debug("@@ isWeekly->{}",isWeekly);
        }catch(SQLException e){
            LOGGER.error("generateReport error: {}",e.getMessage());
        }
            LOGGER.debug("isValidToGenerate() has been executed!");
        return isValid;
    }
}