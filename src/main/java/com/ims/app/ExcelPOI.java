package com.ims.app;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

import java.sql.*;
import java.util.Properties;
import java.util.Map;
import java.util.HashMap;

import javax.naming.Context;
import javax.naming.InitialContext;
import javax.sql.DataSource;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.model.Sheet;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;

import java.math.BigDecimal;
import org.apache.log4j.Logger;

public class ExcelPOI{
   private String path;
   private String reportName;
   private String template;
   private Properties props;
   private static final Logger logger = Logger.getLogger(ReportGenerator.class);
   private String nullString = null;

   public ExcelPOI(String rPath, String rName, String rTemplate, Properties rProperties){
      path=rPath;
      reportName=rName;
      template=rTemplate;
      props = rProperties;
   }

   public boolean createReport(ResultSet result) throws IOException {
      boolean created = false;
      boolean[] totalCells = new boolean[1];
      int cols=0;
      int r=1;
      File file = null;
      InputStream fileinput = null;
      HSSFWorkbook workbook = null;
      HSSFSheet sheet = null;
      String templatesPath = props.getProperty("path.templates");
      String outboxPath = props.getProperty("path.outbox");
      HSSFCell cell = null;
      HSSFRow row = null;
      try{
         file = new File(path+templatesPath+template);
			if(logger.isDebugEnabled()){
				logger.info("@@ path->"+path+templatesPath+template);
			}
         fileinput= new FileInputStream(file);
         if(fileinput!=null){
            workbook = (HSSFWorkbook) WorkbookFactory.create(fileinput);
            sheet = workbook.getSheetAt(0);
            cols=result.getMetaData().getColumnCount();
            totalCells=new boolean[cols];
				if(logger.isDebugEnabled()){
					//logger.debug("@@ - cols="+cols);
				}
				int n=0;
				for(Row rr : sheet) {     
					if(n==1){
					  for(Cell c : rr) {              
							if(c==null){
								totalCells[c.getColumnIndex()]=false;
							}else{
								if(c.getStringCellValue().equals("sum")){
									totalCells[c.getColumnIndex()]=true;
								}
							}
					  }
					  break;
					}
					n++;
				} 
            while(result.next()){
					if(logger.isDebugEnabled()){
						logger.debug("@@ - row="+r);
					}
               row = sheet.createRow(r);
               for(int i=0;i<=cols-1;i++){
                  cell = row.createCell(i);
                  String dataType = result.getMetaData().getColumnTypeName(i+1);
						if(logger.isDebugEnabled()){
							//logger.debug("@@ dataType->"+dataType);
						}
                  switch(dataType){
                     case "bit":
                     	cell.setCellValue(result.getBoolean(i+1));
                        break;
                     case "tinyint":
                     case "int":
                     case "bigint":
                     case "integer":
                     case "smallint":
                        if(result.getInt(i+1)==0){
                           cell.setCellValue(0);
                        }else{
                           cell.setCellValue(result.getInt(i+1));
                        }
                        break;
                     case "float":
                     case "real":
                     case "numeric":
                     case "decimal":
                     case "double":
                        if(result.getDouble(i+1)==0){
                           cell.setCellValue(0);
                        }else{
                           cell.setCellValue(result.getDouble(i+1));
                        }
                        break;
                     case "money":
                     	BigDecimal bigResult = result.getBigDecimal(i+1);
                     	if(bigResult==null){
                     		cell.setCellValue(0);
                     	}else{
									if(bigResult.equals(BigDecimal.ZERO)){
										cell.setCellValue(0);
									}else{
										Double doubleResult = bigResult.doubleValue();
										cell.setCellValue(doubleResult);
									}
								}
                        break;
                     case "date":
                     case "datetime":
                     case "time":
                     case "timestamp":
                     	if(result.getDate(i+1)!=null){
                     		cell.setCellValue(result.getDate(i+1));
                     	}
                        break;
                     case "varchar":
                     case "char":
                     case "text":
                     	if(result.getString(i+1)!=null){
                        	String value=result.getString(i+1);
                        	if(value!=null){
										if(logger.isDebugEnabled()){
											//logger.debug("string->"+value);
										}
                        		cell.setCellValue(value.toString());
                        	}
                        }
                        break;
                     default:
                        if(result.getObject(i+1)!=null){
                        	Object o = result.getObject(i+1);
                           cell.setCellValue(o.toString());
                        }
                  }
               }
               r++;
            }
            row = sheet.createRow(r);
            HSSFCellStyle cellStyle = workbook.createCellStyle();
				cellStyle.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
				cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
				HSSFDataFormat dataFormat = workbook.createDataFormat();
				
            for(int i=0;i<cols;i++){
            	cell = row.createCell(i);                                                                                     
            	cellStyle.setDataFormat(dataFormat.getFormat("_($*#,##0.00_);_($*(#,##0.00);_($*\"-\"??_);_(@_)"));
					cell.setCellStyle(cellStyle);
            	if(totalCells[i]){
					cell.setCellType(HSSFCell.CELL_TYPE_FORMULA);
            		char l=(char)(65+i);
            		String letter=String.valueOf(l);
            		cell.setCellFormula("SUM("+letter+"1:"+letter+String.valueOf(r)+")");
            	}else{
            		
            	}
            }
            FileOutputStream fileOut = new FileOutputStream(outboxPath+reportName);
            workbook.write(fileOut);
            fileOut.close();
            created=true;
         }
      }catch(Exception e){
          logger.error("ExcelPoi.createReport: "+e);
      }
	   if(logger.isDebugEnabled()){
	   	logger.debug("createReport() has been executed!");
	   }
      return created;
   }
}