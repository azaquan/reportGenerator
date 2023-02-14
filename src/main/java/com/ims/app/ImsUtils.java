package com.ims.app;

import java.util.Properties;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Map;
import java.util.HashMap;
import java.io.File;
import java.sql.Connection;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.model.Sheet;
import org.apache.poi.hssf.usermodel.*;

import javax.mail.*;
import javax.mail.internet.*;
import javax.activation.*;
import java.util.Date;
import javax.mail.internet.MimeMessage;

public class ImsUtils{
   public static Properties getProperties(String path){
      Properties props = new Properties();
      FileInputStream input = null;
      String fileName = "reportGenerator.properties";
      try{
         input = new FileInputStream(path+"resources/"+fileName);
         if(input==null){
            System.out.println("Sorry, unable to find "+fileName);
            return props;
         }
         props.load(input);
      }catch(IOException e){
         System.out.println("No properties could be loaded:");
         e.printStackTrace();
      }finally{
         if (input != null) {
            try {
               input.close();
            } catch (IOException e) {
               e.printStackTrace();
            }
         }
      }
      return props;
   }

   public static Map<Integer,HSSFCellStyle> getStyleMap(HSSFSheet sheet, int cols){
      HSSFRow row = sheet.getRow(1);
      Map<Integer,HSSFCellStyle> map = new HashMap<Integer,HSSFCellStyle>();
      if(cols>row.getPhysicalNumberOfCells()){
         cols=row.getPhysicalNumberOfCells();
      }
      for(int i=0;i<=cols-1;i++){
         map.put(i,row.getCell(i).getCellStyle());
      }
      return map;
   }
   
   public static HSSFRow getNewRow(HSSFSheet sheet, Map<Integer,HSSFCellStyle> map, int rowNumber){
   	HSSFRow row = null;
      row = sheet.createRow(rowNumber);
      for(int i=0;i<=map.size()-1;i++){
         HSSFCell cell = row.createCell(i);
         cell.setCellStyle(map.get(i));                   
      }
      return row;
   }

	public static int stringToInt(String value, int _default) {
		 try {
			  return Integer.parseInt(value);
		 } catch (NumberFormatException e) {
			  return _default;
		 }
	}
	
   public static boolean sendEmail(Connection conn, Properties props, Map<String, String> emailMap){
   	boolean sent = false;
   	String attachment = emailMap.get("emailAttachmentFile");
		String to = emailMap.get("emailRecipientStr");
		String msgText = emailMap.get("emailBody");
		System.out.println("msgText - - - - ->"+msgText);
		String subject = emailMap.get("emailSubject");
		
		Map<String, String> map = DbUtils.getEmailConfig(conn);
		if(map!=null){
			String host = map.get("smtp");
			String from = map.get("fromEmail");
			String fromName = map.get("fromName");
			String mailer = map.get("fromName");
			String port = map.get("port");
			String pass = map.get("pass");
			Properties emailProps = System.getProperties();
			
			emailProps.put("mail.debug", "true");

			emailProps.put("mail.transport.protocol", "smtp");

			emailProps.put("mail.smtp.auth", "true");
			
			emailProps.put("mail.smtp.host", host);	
			emailProps.put("mail.from", from);
			emailProps.put("mail.smtp.starttls.enable", "true");
			emailProps.put("mail.smtp.port", port);
			//emailProps.put("mail.smtp.ssl.trust", host);
			emailProps.put("mail.smtp.ssl.protocols", "TLSv1.1 TLSv1.2");
			
			emailProps.setProperty("mail.debug", "false");
			Session session = Session.getInstance(emailProps, null);
			try {
				MimeMessage msg = new MimeMessage(session);
				msg.setFrom(new InternetAddress(from, fromName));
				msg.addRecipients(Message.RecipientType.TO,(InternetAddress.parse(to))); 
				msg.setSubject(subject);
				msg.setSentDate(new Date());
				//msg.settestText(msgText,"text/plain");
				 
				Multipart multipart = new MimeMultipart();
				MimeBodyPart textPart = new MimeBodyPart();
				textPart.setContent(msgText, "text/plain");
				MimeBodyPart attachementPart = new MimeBodyPart();
				attachementPart.attachFile(new File(attachment));
				multipart.addBodyPart(textPart);
				multipart.addBodyPart(attachementPart);
				msg.setContent(multipart);  
				 
				Transport transport = session.getTransport("smtp");
				transport.connect(from, pass);
				transport.sendMessage(msg, msg.getAllRecipients());
				transport.close();
				sent = true;
			}catch (MessagingException mEx) {
				mEx.printStackTrace();
				System.out.println();
			}catch (IOException ioEx){
				ioEx.printStackTrace();
				System.out.println();
			}
		}
		return sent;
   }
}
