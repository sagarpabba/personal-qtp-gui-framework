package com.hp.email;

import java.io.File;
import java.io.IOException;
import java.sql.SQLException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.TimeZone;

//command line is like this :
//-reportdir C:\temp\reporter  -smtpserver smtp3.hp.com  -from chang-yuan.hu@hp.com  -to chang-yuan.hu@hp.com -templatedir C:\temp\Test_ThirtyParty\PAF_Emailer -subject ITG2 

public class PAFEmailer {

	
	@SuppressWarnings("deprecation")
	public static void main(String[] args) throws IOException, ParseException, SQLException {
		
		StringBuffer sb=new StringBuffer();
		for(int i=0;i<=args.length-1;i++){
		   	sb.append(args[i]);
		}
		System.out.println("the args we had passed is:\n"+sb);
		String argsstr=sb.toString();
		int reportlogbegin=argsstr.indexOf("-reportdir");
		//int reporttemplatebegin=argsstr.indexOf("-reporttemplate");
		int smtpbegin=argsstr.indexOf("-smtpserver");
		int frombegin=argsstr.indexOf("-from");
		int tobegin=argsstr.indexOf("-to");
		int templatedirindex=argsstr.indexOf("-templatedir");
		int subjectindex=argsstr.indexOf("-subject");
		
		
		String reportdir=argsstr.substring(reportlogbegin+10, smtpbegin);
		String reportlogfile=reportdir+File.separator+"stepoutput.log";
		String reporttemplate=reportdir+File.separator+"report_template.htm";
		String smtpserver=argsstr.substring(smtpbegin+11, frombegin);
		String fromemail=argsstr.substring(frombegin+5, tobegin);
		String toemail=argsstr.substring(tobegin+3,templatedirindex).trim();
	//	String subject=argsstr.substring().trim();
		String templatedir=argsstr.substring(templatedirindex+12,subjectindex).trim();
		String subject=argsstr.substring(subjectindex+8).trim();
		
		System.out.println("Get the email report file is:"+reportlogfile);
		System.out.println("Get the email report remplate file is:"+reporttemplate);
		System.out.println("Get the smtp server address is:"+smtpserver);
		System.out.println("Get the from email address is:"+fromemail);
		System.out.println("Get the to email address is:"+toemail);
		System.out.println("Get the template file directory is:"+templatedir);
		//System.out.println("get the subject is :"+subject);
		//System.out.println("Get the subject content is:"+subject);
		SimpleDateFormat  runtime=new SimpleDateFormat(
				"EEEE,MMMM dd,yyyy h:mm:ss a z");
		
		runtime.setTimeZone(TimeZone.getTimeZone("GMT"));
		String currenttime=runtime.format(Calendar.getInstance().getTime());
		
		SendEmail.sendMultiEmail(templatedir, reportdir, smtpserver, false, "", "", fromemail, toemail, subject+" On "+currenttime);
		
		//JsoupUtils.getValue(reportlogfile,"START_TIME");
	}
}
