package com.hp.email;


import java.io.BufferedReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.Calendar;

import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;

public class JsoupUtils {



//	//private static final String TODAY_REPORT_DIR = projectpath + "reporter";
//	private final static String CURRENT_TIME = new SimpleDateFormat(
//			"yyyy-MM-dd").format(Calendar.getInstance().getTime());
//	//overview chart
//	//private static File overviewfile=new File(projectpath+"resources"+separator+"Automation_Overview_Template.htm");
//   // private final static String OVERVIEW_REPORT=TODAY_REPORT_DIR+ File.separator + "Automation_Overview.htm";
//    
//    private final static String WEEK_CHART_NAME="3dchart_weekly.jpg";
//    private final static String MONTH_CHART_NAME="3dchart_month.jpg";
//    private final static String PROJECT_CHART_NAME="3dchar_project.jpg";
//	
	
	//we will take to create a new template for the report,the report column will generate automacially
	/**
	 * generate the html report as we had run ,if the test had not run the step will not generate
	 * @return the test execution status:passed ,failed ,no run or not completed
	 * @throws IOException
	 * @throws SQLException 
	 */
	public static String generateHtmlReport(String templatedir,String reportdir) throws IOException, SQLException{
		 //String report_dir=reportdir;
		 String outpuvaluetfile=reportdir+File.separator+"dataoutput.log";
		 String stepfile=reportdir+File.separator+"stepoutput.log";
		 String templatereport=templatedir+File.separator+"report_template.htm";
		 String reportfilename=reportdir+File.separator+"TestingExecutionReport_"+ 
		                      new SimpleDateFormat("yyyy-MM-dd").format(Calendar.getInstance().getTime())+".htm";
		 StringBuilder sb=new StringBuilder();
		 Document doc = Jsoup.parse(new File(templatereport), "UTF-8");
	   
		 //first table for the overview table, 
/**************************************************************************************************************/
		 Element overviewtable = doc.select("table.MsoNormalTable").get(1);
		 int passedcase=0;
		 int failedcase=0;
		 int noruncase=0;
		 int totalcases=0;
		 int warningcase=0;
		 BufferedReader br=null;
		 try {
			br=new BufferedReader(new FileReader(stepfile));
			String strline=null;
			while((strline=br.readLine())!=null){
				totalcases=totalcases+1;
				String status=strline.split("\\|")[2];
				if(status.toLowerCase().trim().contains("pass")){
					passedcase=passedcase+1;
				}
				if(status.toLowerCase().trim().contains("fail")){
					failedcase=failedcase+1;
				}
				if(status.toLowerCase().trim().contains("norun")){
					noruncase=noruncase+1;
				}
				if(status.toLowerCase().trim().contains("warning")){
					warningcase=warningcase+1;
				}
			}
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}finally{
		   br.close();
		}
		System.out.println("Get the case overview ,passed cases are:"+passedcase+",failed cases are:"+failedcase+",no run cases are:"+noruncase);
      
		//insert the record into the database
	//	Connection con=DatabaseUtils.getConnection();
		//first delete the duplicate record for now
	//	int deletenum=DatabaseUtils.updateRecord(con, "delete from qtp_status where date(RUN_TIME)=CURDATE()");
		//System.out.println("We found before we insert today's record there are "+deletenum+"  duplicated records already for today ,so we delete these record...");
	//	int insertednum=DatabaseUtils.updateRecord(con, "insert into qtp_status(PASS_TOTAL,FAIL_TOTAL,NORUN_TOTAL) Values("+passedcase+","+failedcase+","+noruncase+")");
		//System.out.println("we had inserted a new execution "+insertednum+" record for today");
		//con.close();
		//generate a execution run status graph
		
		//  int totalcases=totalcases-1;
        String starttime=FileUtils.getSpliteValue(outpuvaluetfile, "Execution Start Time", "|", 1);
        String totaltime=TimeUtils.howManyMinutes(starttime, TimeUtils.getCurrentTime(Calendar.getInstance().getTime()));
        
        String appendline=""
+"  <tr style=\"mso-yfti-irow:0;mso-yfti-lastrow:yes;height:46.75pt\">"
+"  <td width=\"277\" valign=\"top\" style=\"width:107.6pt;border:solid windowtext 1.0pt;"
+"  border-top:none;mso-border-top-alt:solid windowtext .5pt;mso-border-alt:solid windowtext .5pt;"
+"  background:#D9E2F3;mso-background-themecolor:accent5;mso-background-themetint:"
+"  51;padding:0in 5.4pt 0in 5.4pt;height:46.75pt\">"
+"  <p class=\"MsoListParagraphCxSpFirst\" style=\"margin-left:0in;mso-add-space:auto;"
+"  mso-yfti-cnfc:68\"><b style=\"mso-bidi-font-weight:normal\"><span style=\"font-size:13.0pt;line-height:105%;font-family:&quot;Cambria&quot;,&quot;serif&quot;;"
+"  mso-fareast-font-family:&quot;Cambria&quot;,&quot;serif&quot;;mso-fareast-theme-font:major-fareast;mso-bidi-font-family:"
+"  &quot;Times New Roman&quot;;mso-bidi-theme-font:major-bidi\"><o:p></o:p></span></b></p>"
+"  </td>"
+"  <td width=\"252\" valign=\"top\" style=\"width:117.0pt;border-top:none;border-left:"
+"  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;"
+"  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;"
+"  mso-border-alt:solid windowtext .5pt;background:#D9E2F3;mso-background-themecolor:"
+"  accent5;mso-background-themetint:51;padding:0in 5.4pt 0in 5.4pt;height:46.75pt\">"
+"  <p class=\"MsoListParagraphCxSpMiddle\" style=\"margin-left:0in;mso-add-space:"
+"  auto;mso-yfti-cnfc:64\"><b><span style=\"font-size:13.0pt;line-height:105%;"
+"  font-family:&quot;Cambria&quot;,&quot;serif&quot;;mso-fareast-font-family:&quot;Cambria&quot;,&quot;serif&quot;;mso-fareast-theme-font:"
+"  major-fareast;mso-bidi-font-family:&quot;Times New Roman&quot;;mso-bidi-theme-font:"
+"  major-bidi\">"+totaltime+"<o:p></o:p></span></b></p>"
+"  </td>"
+"  <td width=\"234\" valign=\"top\" style=\"width:2.5in;border-top:none;border-left:none;"
+"  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;"
+"  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;"
+"  mso-border-alt:solid windowtext .5pt;background:#D9E2F3;mso-background-themecolor:"
+"  accent5;mso-background-themetint:51;padding:0in 5.4pt 0in 5.4pt;height:46.75pt\">"
+"  <p class=\"MsoListParagraphCxSpMiddle\" style=\"margin-left:0in;mso-add-space:"
+"  auto;mso-yfti-cnfc:64\"><b><span style=\"font-size:13.0pt;line-height:105%;"
+"  font-family:&quot;Cambria&quot;,&quot;serif&quot;;mso-fareast-font-family:&quot;Cambria&quot;,&quot;serif&quot;;mso-fareast-theme-font:"
+"  major-fareast;mso-bidi-font-family:&quot;Times New Roman&quot;;mso-bidi-theme-font:"
+"  major-bidi\">"+totalcases+"<o:p></o:p></span></b></p>"
+"  </td>"
+"  <td width=\"216\" valign=\"top\" style=\"width:2.5in;border-top:none;border-left:none;"
+"  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;"
+"  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;"
+"  mso-border-alt:solid windowtext .5pt;background:#00B050;padding:0in 5.4pt 0in 5.4pt;"
+"  height:46.75pt\">"
+"  <p class=\"MsoListParagraphCxSpMiddle\" style=\"margin-left:0in;mso-add-space:"
+"  auto;mso-yfti-cnfc:64\"><b><span style=\"font-size:13.0pt;line-height:105%;"
+"  font-family:&quot;Cambria&quot;,&quot;serif&quot;;mso-fareast-font-family:&quot;Cambria&quot;,&quot;serif&quot;;mso-fareast-theme-font:"
+"  major-fareast;mso-bidi-font-family:&quot;Times New Roman&quot;;mso-bidi-theme-font:"
+"  major-bidi\">"+passedcase+"<o:p></o:p></span></b></p>"
+"  </td>"
+"  <td width=\"198\" valign=\"top\" style=\"width:215.85pt;border-top:none;border-left:"
+"  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;"
+"  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;"
+"  mso-border-alt:solid windowtext .5pt;background:#C00000;padding:0in 5.4pt 0in 5.4pt;"
+"  height:46.75pt\">"
+"  <p class=\"MsoListParagraphCxSpLast\" style=\"margin-left:0in;mso-add-space:auto;"
+"  mso-yfti-cnfc:64\"><b><span style=\"font-size:13.0pt;line-height:105%;"
+"  font-family:&quot;Cambria&quot;,&quot;serif&quot;;mso-fareast-font-family:&quot;Cambria&quot;,&quot;serif&quot;;mso-fareast-theme-font:"
+"  major-fareast;mso-bidi-font-family:&quot;Times New Roman&quot;;mso-bidi-theme-font:"
+"  major-bidi\">"+failedcase+"<o:p></o:p></span></b></p>"
+"  <p class=\"MsoNormal\" style=\"text-indent:.5in;mso-yfti-cnfc:64\"><o:p>&nbsp;</o:p></p>"
+"  </td>"
+"  <td width=\"216\" valign=\"top\" style=\"width:355.65pt;border-top:none;border-left:"
+"  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;"
+"  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;"
+"  mso-border-alt:solid windowtext .5pt;background:#FFC000;padding:0in 5.4pt 0in 5.4pt;"
+"  height:46.75pt\">"
+"  <p class=\"MsoListParagraph\" style=\"margin-left:0in;mso-add-space:auto;"
+"  mso-yfti-cnfc:64\"><b><span style=\"font-size:13.0pt;line-height:105%;"
+"  font-family:&quot;Cambria&quot;,&quot;serif&quot;;mso-fareast-font-family:&quot;Cambria&quot;,&quot;serif&quot;;mso-fareast-theme-font:"
+"  major-fareast;mso-bidi-font-family:&quot;Times New Roman&quot;;mso-bidi-theme-font:"
+"  major-bidi\">"+noruncase+"<o:p></o:p></span></b></p>"
+"  </td>"
+"  <td width=\"192\" valign=\"top\" style=\"width:355.65pt;border-top:none;border-left:"
+"  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;"
+"  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;"
+"  mso-border-alt:solid windowtext .5pt;background:#FFC000;padding:0in 5.4pt 0in 5.4pt;"
+"  height:46.75pt\">"
+"  <p class=\"MsoListParagraph\" style=\"margin-left:0in;mso-add-space:auto;"
+"  mso-yfti-cnfc:64\"><b><span style=\"font-size:13.0pt;line-height:105%;"
+"  font-family:&quot;Cambria&quot;,&quot;serif&quot;;mso-fareast-font-family:&quot;Cambria&quot;,&quot;serif&quot;;mso-fareast-theme-font:"
+"  major-fareast;mso-bidi-font-family:&quot;Times New Roman&quot;;mso-bidi-theme-font:"
+"  major-bidi\">"+warningcase+"<o:p></o:p></span></b></p>"
+"  </td>"
+" </tr>";
        overviewtable.append(appendline);
               
/**************************************************************************************************************/
        //the step table,this is the index 1
        Element steptable=doc.select("table.MsoNormalTable").get(2).select("tbody").first();
        System.out.println("the table text is:"+steptable.text());
        BufferedReader stepbr=null;
        stepbr=new BufferedReader(new FileReader(stepfile));
        String stepline=null;
        int stepnumber=0;
        while((stepline=stepbr.readLine())!=null){
        	System.out.println("line is:"+stepline);
        	stepnumber=stepnumber+1;
        	String[] rowdata=stepline.split("\\|");
        	String stepname=rowdata[0].trim();
        	String checkername=rowdata[1].trim();
        	String stepstatus=rowdata[2].trim();
        	String commentdata=rowdata[3].trim();
        	
        	String backgroundcolor="";
        	if(stepstatus.toLowerCase().contains("pass")){
        		backgroundcolor="background:#00B050;";
        	}
        	if(stepstatus.toLowerCase().contains("failed")){
        		backgroundcolor="background:#C00000;";
        	}
        	if(stepstatus.toLowerCase().contains("norun")){
        		backgroundcolor="background:#FFC000;";
        	}
        	if(stepstatus.toLowerCase().contains("warning")){
        		backgroundcolor="background:#FFC000;";
        	}
            
        String tablerow=""
+ " <tr style=\"mso-yfti-irow:"+(stepnumber+1)+";height:15.75pt\">"
+"  <td width=\"67\" valign=\"top\" style=\"width:49.05pt;border:solid windowtext 1.0pt;"
+"   border-top:none;mso-border-top-alt:solid windowtext .5pt;mso-border-alt:solid windowtext .5pt;"
+"   padding:0in 0in 0in 0in;height:15.75pt\">"
+"   <p style=\"margin-left:0in;mso-add-space:auto\" class=\"MsoListParagraphCxSpFirst\"><b><span style=\"font-size:13.0pt;line-height:105%;font-family:&quot;Cambria&quot;,&quot;serif&quot;;"
+"   mso-fareast-font-family:  &quot;Times New Roman&quot;;mso-fareast-theme-font:major-fareast;mso-bidi-font-family:"
+"   &quot;Times New Roman&quot;;mso-bidi-theme-font:major-bidi\">"+stepnumber+"<o:p></o:p></span></b></p>"
+"   </td>"
+"   <td width=\"354\" valign=\"bottom\" style=\"width:148.55pt;border-top:none;border-left:"
+"   none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;"
+"   mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;"
+"   mso-border-alt:solid windowtext .5pt;padding:0in 5.4pt 0in 5.4pt;height:15.75pt\">"
+"   <p style=\"margin-left:0in;mso-add-space:"
+"   auto\" class=\"MsoListParagraphCxSpMiddle\"><b><span style=\"font-size:13.0pt;line-height:105%;font-family:&quot;Cambria&quot;,&quot;serif&quot;;"
+"   mso-fareast-font-family:  &quot;Times New Roman&quot;;mso-fareast-theme-font:major-fareast;mso-bidi-font-family:"
+"   &quot;Times New Roman&quot;;mso-bidi-theme-font:major-bidi\">"+stepname+"<o:p></o:p></span></b></p>"
+"   </td>"
+"   <td width=\"426\" valign=\"bottom\" style=\"width:324.45pt;border-top:none;border-left:"
+"   none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;"
+"   mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;"
+"   mso-border-alt:solid windowtext .5pt;padding:0in 5.4pt 0in 5.4pt;height:15.75pt\">"
+"   <p style=\"margin-left:0in;mso-add-space:"
+"   auto\" class=\"MsoListParagraphCxSpMiddle\"><span style=\"font-size:13.0pt;line-height:105%;font-family:&quot;Cambria&quot;,&quot;serif&quot;;"
+"   mso-fareast-font-family: &quot;Times New Roman&quot;;;mso-fareast-theme-font:major-fareast;mso-bidi-font-family:"
+"   &quot;Times New Roman&quot;;mso-bidi-theme-font:major-bidi;mso-bidi-font-weight:bold\">"+checkername+"<o:p></o:p></span></p>"
+"   </td>"
+"   <td width=\"126\" valign=\"bottom\" nowrap=\"\" style=\"width:242.55pt;border-top:none;"+backgroundcolor
+"   border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;"
+"   mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;"
+"   mso-border-alt:solid windowtext .5pt;padding:0in 5.4pt 0in 5.4pt;height:15.75pt\">"
+"   <p style=\"margin-left:0in;mso-add-space:"
+"   auto\" class=\"MsoListParagraphCxSpMiddle\"><b><span style=\"font-size:13.0pt;line-height:105%;font-family:&quot;Cambria&quot;,&quot;serif&quot;;"
+"   mso-fareast-font-family: &quot;Times New Roman&quot;;mso-fareast-theme-font:major-fareast;mso-bidi-font-family:"
+"   &quot;Times New Roman&quot;;mso-bidi-theme-font:major-bidi\">"+stepstatus+"<o:p></o:p></span></b></p>"
+"   </td>"
+"   <td width=\"612\" valign=\"bottom\" nowrap=\"\" style=\"width:5.5in;border-top:none;"
+"   border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;"
+"   mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;"
+"   mso-border-alt:solid windowtext .5pt;padding:0in 5.4pt 0in 5.4pt;height:15.75pt\">"
+"   <p style=\"margin-left:0in;mso-add-space:auto\" class=\"MsoListParagraphCxSpLast\"><span style=\"font-size:11.0pt;line-height:105%;font-family:&quot;Cambria&quot;,&quot;serif&quot;;"
+"   mso-fareast-font-family: &quot;Times New Roman&quot;;mso-fareast-theme-font:major-fareast;mso-bidi-font-family:"
+"   &quot;Times New Roman&quot;;mso-bidi-theme-font:major-bidi\">"+commentdata+"<o:p></o:p></span></p>"
+"   </td>"
+"   </tr>";
        
        steptable.append(tablerow);
       }
       stepbr.close();
/**************************************************************************************************************/
        //update the data value table
        
        Element valuetable=doc.select("table.MsoNormalTable").get(3).select("tbody").first();
        BufferedReader valuebr=null;
        valuebr=new BufferedReader(new FileReader(outpuvaluetfile));
        String valueline=null;
        int valuenumber=0;
        while((valueline=valuebr.readLine())!=null){
        	valuenumber=valuenumber+1;
        	String[] rowdatavalue=valueline.split("\\|");
        	String datades=rowdatavalue[0];
        	String datausedvalue=rowdatavalue[1];
        String valuerow=""
+"  <tr style=\"mso-yfti-irow:0;height:18.2pt\">"
+"  <td width=\"299\" valign=\"top\" style=\"width:202.1pt;border:solid white 1.0pt;"
+"  mso-border-themecolor:background1;border-top:none;mso-border-top-alt:solid white .5pt;"
+"  mso-border-top-themecolor:background1;mso-border-alt:solid white .5pt;"
+"  mso-border-themecolor:background1;background:#4472C4;mso-background-themecolor:"
+"  accent5;padding:0in 5.4pt 0in 5.4pt;height:18.2pt\">"
+"  <p style=\"tab-stops:center 95.6pt;mso-yfti-cnfc:68\" class=\"MsoNormal\"><b><span style=\"font-size:9.0pt;line-height:105%;font-family:&quot;Cambria&quot;,&quot;serif&quot;;"
+"  mso-bidi-font-family:Arial;color:white;mso-themecolor:background1\">"+datades+"<o:p></o:p></span></b></p>"
+"  </td>"
+"  <td width=\"1286\" valign=\"top\" style=\"width:958.5pt;border-top:none;border-left:"
+"  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;"
+"  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;"
+"  mso-border-alt:solid windowtext .5pt;background:#B4C6E7;mso-background-themecolor:"
+"  accent5;mso-background-themetint:102;padding:0in 5.4pt 0in 5.4pt;height:18.2pt\">"
+"  <p style=\"mso-yfti-cnfc:64\" class=\"MsoNormal\"><span style=\"font-size:9.0pt;"
+"  line-height:105%;font-family:&quot;Cambria&quot;,&quot;serif&quot;;mso-bidi-font-family:Arial\">"+datausedvalue+"<o:p></o:p></span></p>"
+"  </td>"
+" </tr>";
        valuetable.append(valuerow);
        }
        valuebr.close();
        
/**************************************************************************************************************/
    	// this is the insertimage code MsoTableMediumShading2Accent5
		Element imagetbodynode = doc.select("table.MsoNormalTable").get(4).select("tbody").first();
		System.out.println("the image node is :" + imagetbodynode.text());


		int filesize = FileUtils.getSubFilesSize(reportdir, ".png");
		File[] errorshotfile = FileUtils.listFilesEndsWith(reportdir, ".png");
		if (filesize > 0) {
			System.out.println("this testing had generated the error screenshot, so we will put into this screenshot into email ");
			for (int fileindex = 0; fileindex < filesize; fileindex++) {
				int screenshotnum = fileindex + 1;
				
				String imagecode=""
						+ "<tr><td><p>"+errorshotfile[fileindex].getName() +"</p></td><td><span> <img src=\"cid:image"+screenshotnum+"@hp\" alt=\" "
						+ "this image cannot show,maybe blocked by your email setting \"><br></span></td></tr>";
				imagetbodynode.append(imagecode);
			}
		}
		System.out.println("We had completed pop all the prior data into a Map type ,we will use it later");
		
		//emabled the company logo picture  ,the image CID must be started with image ...
		//image code is:<img width=79 height=79 src=\"cid:companylogo\" align=right hspace=12 alt=\"this image cannot show,maybe blocked by your email setting \" v:shapes=\"Picture_x0020_2\">
		Element notetable=doc.select("p.MsoNormal").first();
        notetable.prepend("<img width=79 height=79"
        		+ " src=\"cid:imagecompanylog@hp\""
        		+ "align=right hspace=12 alt=\"this image cannot show,maybe blocked by your email setting\">");
        System.out.println("we had attached the company logo into the email ...");
        
        //System.out.println("notetable html is:"+notetable.toString());
		// generate the html report
	/**************************************************************************************************************/
		sb.append(doc.body().html());
		FileUtils.writeContents(reportfilename, sb.toString());
		String filepath="\\\\pdeauto06.fc.hp.com\\PAF_Automation\\Test_Reports\\";
		//FileUtils.copyFile(reportfilename,filepath+ "TestingExecutionReport_"+ 
         //       new SimpleDateFormat("yyyy-MM-dd").format(Calendar.getInstance().getTime())+".htm");
		System.out.println("Copied today's report to the share path:"+filepath);
		// get the total run-time status
		if (passedcase == totalcases) {
			return "Passed";
		} else if (failedcase > 0) {
			return "Failed";
		} else if (noruncase == totalcases) {
			return "No Run";
		} else if((warningcase>0)&&(failedcase==0)&&(noruncase==0)){
			return "Warning";
		}else{
			return "Not Completed";
		}
	}
	
}
