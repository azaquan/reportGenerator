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
import org.apache.logging.log4j.Logger;
import org.apache.logging.log4j.LogManager;

public class ExcelPOI{
   private String path;
   private String reportName;
   private String template;
   private Properties props;
   private String title;
   private String period;
   private static final Logger LOGGER = LogManager.getLogger(ExcelPOI.class);
   private String nullString = null;
   private String matrix="";
   private int rowFrom;
   private int rowTo;
   ResultSet result;
	HSSFWorkbook workbook = null;
	HSSFSheet sheet = null;
	HSSFCell cell = null;
	HSSFRow row = null;
	int cols=0;
	int r=3;

   public ExcelPOI(
   		String rPath, 
   		String rName, 
   		String rTemplate, 
   		Properties rProperties, 
   		String reportTitle, 
   		String reportPeriod,
   		String reportMatrix,
   		int reportrowFrom,
   		int reportrowTo){
      path=rPath;
      reportName=rName;
      template=rTemplate;
      props = rProperties;
      if(reportTitle.equals(nullString)){
      	title=reportName;
      }else{
      	title = reportTitle;
      }
      if(reportPeriod.equals(nullString)){
      	period="";
      }else{
      	period = reportPeriod;
      }
      matrix=reportMatrix;
      LOGGER.debug("@@@@@@ - matrix={}",reportMatrix);
      rowFrom=reportrowFrom;
      rowTo=reportrowTo;
   }

   public boolean createReport(ResultSet queryResult) throws IOException {
   	result=queryResult;
      boolean created = false;
      boolean[] totalCells = new boolean[1];
      File file = null;
      InputStream fileinput = null;
      String templatesPath = props.getProperty("path.templates");
      String outboxPath = props.getProperty("path.outbox");
      Double valueD;
      try{
         file = new File(path+templatesPath+template);
         fileinput= new FileInputStream(file);
         if(fileinput!=null){
            workbook = (HSSFWorkbook) WorkbookFactory.create(fileinput);
            sheet = workbook.getSheetAt(0);
            cols=result.getMetaData().getColumnCount();
            totalCells=new boolean[cols];
				if(LOGGER.isDebugEnabled()){
					LOGGER.debug("@@ - cols=",cols);
				}

				CellStyle cellStyle = workbook.createCellStyle();
				org.apache.poi.ss.usermodel.Font font = workbook.createFont();
				font.setFontHeightInPoints((short)12);
				font.setFontName("Arial Black");
				cellStyle.setFont(font);
				Cell cc = null;
				int n=0;
				for(Row rr : sheet) {    
					switch (n){
						case 0:
							cc = rr.getCell(0);
							if (cc==null){
								cc=rr.createCell(0);
								cc.setCellType(Cell.CELL_TYPE_STRING);
								cc.setCellStyle(cellStyle);
							}
							cc.setCellValue(title);
							break;
						case 1:
							cc = rr.getCell(0);
							if (cc==null){
								cc=rr.createCell(0);
								cc.setCellType(Cell.CELL_TYPE_STRING);
								cc.setCellStyle(cellStyle);
							}
							cc.setCellValue(period);
							break;
						case 2:
							break;
						case 3:
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
						default:
							break;
					}
					n++;
				} 
				boolean done = false;
				boolean doTotals = false;
				if(!matrix.equals("") && rowFrom>0 && rowTo>=rowFrom){
					LOGGER.debug("@@ - feeding matrix mode ------------------------------@@");
					if (getDataIntoMatrixSheet()) done=true;
				}else{
					LOGGER.debug("@@ - feeding regular mode ------------------------------@@");
					if (getDataIntoSheet()){
						done=true;
						doTotals=true;
					}
				}
				LOGGER.debug("-eof-");
				if(done){
					if(doTotals){
						row = sheet.createRow(r);
						cellStyle = workbook.createCellStyle();
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
							}
						}
					}
					Header header = sheet.getHeader();
					header.setCenter(title);
					if(LOGGER.isDebugEnabled()){
						LOGGER.debug("- title - {}",title);
					}
					workbook.setSheetName(0, title);   
					FileOutputStream fileOut = new FileOutputStream(outboxPath+reportName);
					if(LOGGER.isDebugEnabled()){
						LOGGER.debug("- fileOut - {}",fileOut);
					}
					workbook.write(fileOut);
					fileOut.close();
					created=true;
				}
         }
      }catch(Exception e){
          LOGGER.error("ExcelPoi.createReport: {}",e);
      }
	  LOGGER.debug("createReport() has been executed!");
      return created;
   }
   
   private boolean getDataIntoMatrixSheet(){
   boolean passed = false;
   int rowNumber=0;
   int matrixCursor=0;
  	try{
			String[] guide = matrix.split(",");
			if(LOGGER.isDebugEnabled()){
				LOGGER.debug("- - - - - - - - -      ------------ matrix->{}",matrix);
			}
			int refCol = 0;
			int targetCol = 0;
			int sourceColumn = 1;
			for(int i=0;i<guide.length-1;i++){
				if(guide[i].equals(String.valueOf(sourceColumn))){
					refCol=i;
					if(LOGGER.isDebugEnabled()){
						LOGGER.debug("- - - - - - - - -      ------------ refCol->{}",refCol);
					}
				}
			}
			while(result.next()){
				for(int i=0;i<=guide.length-1;i++){ //iterates horizontally within table columns
					//To get into the existing sheet to first check within the refCol
					//and validate if any of the cells matches with the source value and then update it
					//linen=row is to scan vertically the sheet
					if(i==sourceColumn){
						for(int line=0;line<=rowTo;line++){ //iterates only within the area delimited by rowFrom - rowTo
							Cell cc = null;
							for(Row rr : sheet) { //iterates within the entire sheet to find the matching row or line
								if(rr.getRowNum()==line){ //it filters only the areato evaluate
									cc = rr.getCell(refCol);
									if (cc!=null){
										if (cc.getCellType()==Cell.CELL_TYPE_STRING){
											String cellRef=cc.getStringCellValue();
											cellRef=cellRef.toUpperCase().trim();
											if(LOGGER.isDebugEnabled()){
												LOGGER.debug("- - - - - - - - - --sourceColumn->{}",sourceColumn);
											}
											String dataRef=result.getString(sourceColumn);
											dataRef=dataRef.toUpperCase().trim();
											if (cellRef.equals(dataRef)){ //when both references match, the data can be written down
												if(LOGGER.isDebugEnabled()){
													LOGGER.debug("- - - - - - - - - --cellRef/dataRef-->"+cellRef+"/"+dataRef+" "+ cellRef.equals(dataRef));
												}
												int dataCol;
												for(int m=1;m<guide.length ;m++){ //iterates within the matrix array 
													try{
														Cell ccc = rr.getCell(m);
														dataCol=Integer.parseInt(guide[m]);
														if(dataCol>0){ // it does not take zero because it is the ref source col
															if(LOGGER.isDebugEnabled()){
																LOGGER.debug("- - - - - - - - - ---dataCol:{}",dataCol);
															}
															stampValue(ccc, dataCol);
															passed = true;
														}
													}catch(Exception e){
														dataCol=0;
													}
												}
											}
										}
									}
								}
							}
						}
					}
				}
				r++;
			}
			HSSFFormulaEvaluator.evaluateAllFormulaCells(workbook);
		}catch(Exception e){
          LOGGER.error("ExcelPoi.createReport: {}",e);
          passed = false;
      }
      return passed;
   }
   
   private boolean getDataIntoSheet(){
   	boolean passed = false;
   	try{
			while(result.next()){
				if(LOGGER.isDebugEnabled()){
					LOGGER.debug("-- row={}",r);
				}
				row = sheet.createRow(r);
				for(int i=0;i<=cols-1;i++){
					cell = row.createCell(i);
					String dataType = result.getMetaData().getColumnTypeName(i+1);
					if(LOGGER.isDebugEnabled()){
						LOGGER.debug( "column->"+result.getMetaData().getColumnName(i+1)+"/" +dataType);
					}
					switch(dataType){
						case "bit":
							cell.setCellValue(result.getBoolean(i+1));
							if(!result.wasNull()){
								if(LOGGER.isDebugEnabled()){
									LOGGER.debug("bit->"+result.getBoolean(i+1));
								}
							}
							break;
						case "tinyint":
						case "int":
						case "bigint":
						case "integer":
						case "smallint":
							cell.setCellValue(result.getInt(i+1));
							if(!result.wasNull()){
								if(LOGGER.isDebugEnabled()){
									LOGGER.debug("int family->"+result.getInt(i+1));
								}
							}
							break;
						case "float": 
						case "real":
						case "numeric":
						case "decimal":
						case "double":
							cell.setCellValue(result.getDouble(i+1));
							if(!result.wasNull()){
								if(LOGGER.isDebugEnabled()){
									LOGGER.debug("double family->"+result.getDouble(i+1));
								}
							}
							break;
						case "money":
							BigDecimal bigResult = result.getBigDecimal(i+1);
							if(result.wasNull()){
								cell.setCellValue(0);
							}else{
								if(bigResult.equals(BigDecimal.ZERO)){
									cell.setCellValue(0);
								}else{
									Double doubleResult = bigResult.doubleValue();
									cell.setCellValue(doubleResult);
									if(LOGGER.isDebugEnabled()){
										LOGGER.debug("money->"+doubleResult);
									}
								}
							}
							break;
						case "date":
						case "datetime":
						case "time":
						case "timestamp":
							Timestamp timestamp = result.getTimestamp(i+1);
							if(result.wasNull()){
							}else{
								if (timestamp != null){
									if(LOGGER.isDebugEnabled()){
										LOGGER.debug("date family->"+result.getDate(i+1));
									}
									cell.setCellValue(result.getDate(i+1));
								}
							}
							break;
						case "text":
						case "char":
						case "varchar":
							String textValue=result.getString(i+1);
							if(LOGGER.isDebugEnabled()){
								LOGGER.debug("string----------------->"+textValue);
							}
							if(result.wasNull()){
							}else{
								if(textValue.equals(nullString)){
								}else{
									cell.setCellValue(textValue);
								}
							}
							break;
						default:
							Object o = (result.getObject(i+1));
							if(!result.wasNull()){
								cell.setCellValue("Process Error");
								if(LOGGER.isDebugEnabled()){
									LOGGER.debug("object->"+result.getObject(i+1));
								}
							}
					}
				}
				r++;
			}
			passed = true;
			LOGGER.debug("all data got into the sheet? {}",passed);
		}catch(Exception e){
         	LOGGER.error("ExcelPoi.createReport: {}",e);
      }
      return passed;
   }
   
   private void stampValue(Cell cell, int colPos){
   	   LOGGER.info("- - - - - - - - - STAMP");
			try{
				String dataType = result.getMetaData().getColumnTypeName(colPos);
				if(LOGGER.isDebugEnabled()){
					LOGGER.debug( "column->"+result.getMetaData().getColumnName(colPos)+"/" +dataType);
				}
				switch(dataType){
					case "bit":
						cell.setCellValue(result.getBoolean(colPos));
						if(!result.wasNull()){
							if(LOGGER.isDebugEnabled()){
								LOGGER.debug("bit->"+result.getBoolean(colPos));
							}
						}
						break;
					case "tinyint":
					case "int":
					case "bigint":
					case "integer":
					case "smallint":
						cell.setCellValue(result.getInt(colPos));
						if(!result.wasNull()){
							if(LOGGER.isDebugEnabled()){
								LOGGER.debug("int family->"+result.getInt(colPos));
							}
						}
						break;
					case "float": 
					case "real":
					case "numeric":
					case "decimal":
					case "double":
						cell.setCellValue(result.getDouble(colPos));
						if(!result.wasNull()){
							if(LOGGER.isDebugEnabled()){
								LOGGER.debug("double family->"+result.getDouble(colPos));
							}
						}
						break;
					case "money":
						BigDecimal bigResult = result.getBigDecimal(colPos);
						if(result.wasNull()){
							cell.setCellValue(0);
						}else{
							if(bigResult.equals(BigDecimal.ZERO)){
								cell.setCellValue(0);
							}else{
								Double doubleResult = bigResult.doubleValue();
								cell.setCellValue(doubleResult);
								if(LOGGER.isDebugEnabled()){
									LOGGER.debug("money->"+doubleResult);
								}
							}
						}
						break;
					case "date":
					case "datetime":
					case "time":
					case "timestamp":
						Timestamp timestamp = result.getTimestamp(colPos);
						if(result.wasNull()){
						}else{
							if (timestamp != null){
								if(LOGGER.isDebugEnabled()){
									LOGGER.debug("date family->"+result.getDate(colPos));
								}
								cell.setCellValue(result.getDate(colPos));
							}
						}
						break;
					case "text":
					case "char":
					case "varchar":
						String textValue=result.getString(colPos);
						if(LOGGER.isDebugEnabled()){
							LOGGER.debug("string----------------->{}",textValue);
						}
						if(result.wasNull()){
						}else{
							if(textValue.equals(nullString)){
							}else{
								cell.setCellValue(textValue);
							}
						}
						break;
					default:
						Object o = (result.getObject(colPos));
						if(!result.wasNull()){
							cell.setCellValue("Process Error");
							if(LOGGER.isDebugEnabled()){
								LOGGER.debug("object->"+result.getObject(colPos));
							}
						}
				}
		}catch(Exception e){
			LOGGER.error("ExcelPoi.stampValue: {}",e);
      }
   }
}