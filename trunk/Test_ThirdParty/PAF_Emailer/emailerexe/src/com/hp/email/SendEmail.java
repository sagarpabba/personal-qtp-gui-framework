package com.hp.email;


import java.io.File;
import java.io.IOException;
import java.sql.SQLException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.Properties;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.Authenticator;
import javax.mail.BodyPart;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;



public class SendEmail {

	
	
	public static void sendMultiEmail(String templatedir,String logdir,String smtpserver, boolean authenuser,
			final String username, final String password, String from,
			String to, String subject) throws IOException, ParseException, SQLException {
		    final String CURRENT_TIME = new SimpleDateFormat("yyyy-MM-dd")
		.format(Calendar.getInstance().getTime());
// constant

          final String TODAY_REPORT = logdir+File.separator
		+ "TestingExecutionReport_" + CURRENT_TIME + ".htm";

		try {
			String currenttime = new SimpleDateFormat(
					"EEEE,MMMM dd,yyyy h:mm:ss a z").format(Calendar
					.getInstance().getTime());

			Properties prop = new Properties();
			// prop.put("mail.smtp.auth", "true");
			prop.put("mail.smtp.host", smtpserver);
			prop.put("mail.smtp.port", "25");
			prop.put("mail.debug", "true");
			Session session = null;
			if (authenuser) {
				prop.put("mail.smtp.auth", "true");
				// prop.put("mail.smtp.starttls.enable", "true");
				// prop.put("mail.smtp.ssl.enable","true");
				    session = Session.getDefaultInstance(prop, new Authenticator() {
					@Override
					protected PasswordAuthentication getPasswordAuthentication() {
						return new PasswordAuthentication(username, password);
					}
				});
			} else {
				session = Session.getDefaultInstance(prop);
			}

			String subjecttitle = JsoupUtils.generateHtmlReport(templatedir,logdir);
			// System.out.println("we had generate a test report file in this path:"+TODAY_REPORT);
			// email settings
			MimeMessage mime = new MimeMessage(session);
			mime.setFrom(new InternetAddress(from));
			mime.setRecipients(Message.RecipientType.TO,
					InternetAddress.parse(to));
			
			//get the environment value
			String executionid=System.getenv("ExecutionID");
			if(executionid==null){
				executionid="ScriptError";
			}
			mime.setSubject("[" + subjecttitle + "] Execution ID:" +executionid+" -" +subject);
			mime.setSentDate(new Date());

			// set multipart email
			MimeMultipart multipart = new MimeMultipart("related");

			// read the today's report content and send the email
			String htmlcontents = FileUtils.returnFileContents(TODAY_REPORT);
			// FileUtil.writeContents("testlog.log", htmlcontents);
			System.out.println("read all the today's report content into the string already now");

			BodyPart bodypart = new MimeBodyPart();
			bodypart.setContent(htmlcontents, "text/html;charset=UTF-8");
			System.out.println("now add the html content into the email's body content");
			bodypart.setHeader("Content-Type", "text/html;charset=UTF-8");
			// first set the body main content
			multipart.addBodyPart(bodypart);
			System.out.println("complete parsing the html body content");

			// add the image file
			int filesize = FileUtils.getSubFilesSize(logdir, ".png");
			File[] errorshotfile = FileUtils.listFilesEndsWith(logdir, ".png");
			if (filesize > 0) {
				System.out.println("the current testing met the error and generate an error screenshot we will attach it into the email");
				for (int fileindex = 0; fileindex < filesize; fileindex++) {
					int imagecount = fileindex + 1;
					String errorfilepath = errorshotfile[fileindex]
							.getAbsolutePath();
					bodypart = new MimeBodyPart();
					DataSource fds = new FileDataSource(errorfilepath);
					bodypart.setDataHandler(new DataHandler(fds));
					System.out.println("the error screenshot file long path is :"
							+ errorfilepath);
					System.out.println("the email's image content id is :image"
							+ imagecount + "@hp");
					// be careful this content must contain with <>
					bodypart.addHeader("Content-ID", "<image" + imagecount
							+ "@hp>");
					bodypart.setDisposition("inline");
					multipart.addBodyPart(bodypart);
					System.out.println("add the image embled into the html for this screenshot file :"
							+ errorfilepath);
				}
			}
			System.out.println("complete parsing the image body content");

			// Part two is attachment,set the email's body attachments
		//	bodypart = new MimeBodyPart();
		//	DataSource source = new FileDataSource();
		//	bodypart.setDataHandler(new DataHandler(source));
		///	bodypart.setFileName("reporter.log");
		//	multipart.addBodyPart(bodypart);
		//	System.out.println("complete parsing the attachment body content");
			
			//the company logo picture  companylog
			String logopath=templatedir+File.separator+"logo.jpg";
			bodypart = new MimeBodyPart();
			DataSource ds = new FileDataSource(logopath);
			bodypart.setDataHandler(new DataHandler(ds));
			System.out.println("the email logo file path is :"
					+ logopath);
			System.out.println("the email's image content id is :companylog");
			// be careful this content must contain with <>
			bodypart.addHeader("Content-ID", "<imagecompanylog@hp>");
			bodypart.setDisposition("inline");
			//bodypart.setFileName("logo.png");
			multipart.addBodyPart(bodypart);

			// // Send the complete message parts
			mime.setContent(multipart, "text/html;charset=UTF-8");
			System.out.println("now add  all the contents into html body,is sending the email now.....");
			Transport.send(mime);
			System.out.println("now  send the email successfully......");
			// transport.close();
		} catch (MessagingException e) {
			System.out.println("sorry met the MessagingException  error cannot send the email ");
			System.out.println(e);
		}
	}

}
