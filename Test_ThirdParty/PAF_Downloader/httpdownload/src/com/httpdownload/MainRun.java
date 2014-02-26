package com.httpdownload;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Random;

import net.sf.json.JSONArray;
import net.sf.json.JSONObject;

import org.apache.http.Header;
import org.apache.http.HttpResponse;
import org.apache.http.message.BasicHeader;
import org.kohsuke.args4j.Argument;
import org.kohsuke.args4j.CmdLineException;
import org.kohsuke.args4j.CmdLineParser;
import org.kohsuke.args4j.Option;

import com.google.common.collect.ArrayListMultimap;
import com.google.common.collect.Multimap;

import static org.kohsuke.args4j.ExampleMode.ALL;


/*
 * 
 *                 Author:Alter(chang-yuan.hu@hp.com)
 *                 Create time:2014-01-16 16:00
 *                 usage:
 */
public class MainRun {

	
		@Option(name="-type",required=false,usage="specified which PAF environment you need to run the assessment,like it1,it2,fut1.by default run in itg1,etc")
	    private String type="itg1";
	
	    @Option(name="-username",required=false,usage="specified the PAF login username,by default using chang-yuan.hu@hp.com account")
	    private String username="chang-yuan.hu@hp.com";	
	    @Option(name="-password",required=false,usage="specified the PAF login password")
	    private String password="gu.chan-1026";   
	    @Option(name="-downloaddir",required=false,usage="specified the folder where you put the downloaded assessment")
	    private String downloaddir="C:\\PAF_Run_Result\\"; 
	    
	
	 
	    // receives other command line parameters than options
	    @Argument
	    private  List<String> arguments = new ArrayList<String>();

	    
	public static void main(String[] args) throws Exception {
		    long tStart = System.currentTimeMillis();
			new MainRun().run(args);
			long tEnd = System.currentTimeMillis();
			System.out.println("Download Assessment Total Time Takes: "+((tEnd-tStart) / 1000.0)+" Seconds.Enjoy it! :)");
	}
	
	@SuppressWarnings("deprecation")
	public void run(String[] args) throws Exception {
		
		 CmdLineParser parser = new CmdLineParser(this);  
         
	        // if you have a wider console, you could increase the value;  
	        // here 80 is also the default  
	       // parser.setUsageWidth(80);  	  
	        try {  
		            //System.out.println("parse the cmd line string is:"+args);  
		            parser.parseArgument(args);  
		    
		            if( arguments.isEmpty() )  
		                 throw new CmdLineException("No argument is given");  
	  
	        } catch( CmdLineException e ) {  
		            // if there's a problem in the command line,  
		            // you'll get this exception. this will report  
		            // an error message.  
		            System.err.println(e.getMessage());  
		            System.err.println("java -jar PAFDownloader.jar [options...] arguments...");  
		            // print the list of available options  
		            parser.printUsage(System.err);  
		            System.err.println();  
		  
		            // print option sample. This is useful some time  
		            System.err.println(" Example: java -jar PAFDownloader.jar "+parser.printExample(ALL));  	  
		            return;  
	        } 
	        
	       String runid=arguments.get(0);
	       System.out.println("PAF Envrionment type is:"+type);	     
	       System.out.println("PAF Run ID is:"+runid);
	       System.out.println("PAF login username is:"+username);
	      // System.out.println("PAF login password is:"+password);
	       
	       String pafbaseurl = null;
	       //String assessmentcode=null;
	       String loginurl=null;
	    
	      // get the environemnt
	       if(type.equalsIgnoreCase("itg1")){
	    	   pafbaseurl="https://proactive-assessments-itg1.corp.hp.com/web";
	    	   loginurl="https://login-itg.houston.hp.com/rdlogin/fcc/rdloginb.fcc?TYPE=33554433&REALMOID=06-35e88573-7409-4e04-ac92-50b6645dd51b&GUID=&SMAUTHREASON=0&METHOD=GET&SMAGENTNAME=Y29Uura4LKM36BZRFK7jkiddzgVt2oHymRDCpQlnc3HFiFU8OhjzQoOr9Qh0jk0Z&TARGET=-SM-https%3a%2f%2fproactive--assessments--itg1%2ecorp%2ehp%2ecom%2fweb";

	    	  
	       }else if(type.equalsIgnoreCase("itg2")){
	    	   pafbaseurl="https://proactive-assessments-itg2.corp.hp.com:1181/web";
	    	   loginurl="https://login-itg.houston.hp.com/rdlogin/fcc/rdloginb.fcc?TYPE=33554433&REALMOID=06-35e88573-7409-4e04-ac92-50b6645dd51b&GUID=&SMAUTHREASON=0&METHOD=GET&SMAGENTNAME=-SM-bEKH4IauPeWGucWPMmhf%2bG91h3nTUNhWBqAOms0CXkGAmAch6XyiJHADF6oIvS0G&TARGET=-SM-https%3a%2f%2fproactive--assessments--itg2%2ecorp%2ehp%2ecom%3a1181%2fweb";
	       }else if(type.equalsIgnoreCase("fut1")){
	    	   pafbaseurl="https://proactive-assessments-fut1.corp.hp.com/web";
	    	   loginurl="https://login-itg.houston.hp.com/rdlogin/fcc/rdloginb.fcc?TYPE=33554433&REALMOID=06-35e88573-7409-4e04-ac92-50b6645dd51b&GUID=&SMAUTHREASON=0&METHOD=GET&SMAGENTNAME=-SM-ImXGLY1uGJ5OBnMSr8o%2foXqArhiyGe4xelZQkVcCLvqWts6gAuV5heeO98mvlHyW&TARGET=-SM-https%3a%2f%2fproactive--assessments--fut1%2ecorp%2ehp%2ecom%2fweb";
	       }else if(type.equalsIgnoreCase("fut2")){
	    	   pafbaseurl="https://proactive-assessments-fut2.corp.hp.com/web";
	    	   loginurl="https://proactive-assessments-fut2.corp.hp.com/web/spring_security_login;jsessionid=7095B5D75A57187E913D7AFDC84D8A65.IRSLFCHOST02_WHA-General-Inst";
	       }else{
	    	   pafbaseurl="https://proactive-assessments-itg1.corp.hp.com/web"; 
	    	   loginurl="https://login-itg.houston.hp.com/rdlogin/fcc/rdloginb.fcc?TYPE=33554433&REALMOID=06-35e88573-7409-4e04-ac92-50b6645dd51b&GUID=&SMAUTHREASON=0&METHOD=GET&SMAGENTNAME=Y29Uura4LKM36BZRFK7jkiddzgVt2oHymRDCpQlnc3HFiFU8OhjzQoOr9Qh0jk0Z&TARGET=-SM-https%3a%2f%2fproactive--assessments--itg1%2ecorp%2ehp%2ecom%2fweb";
	       }
	    
	       String pafrunidurl=pafbaseurl+"/run/"+runid;
    	   String startdownloadurl=pafbaseurl+"/ajax/startDownload";
    	   String completedownloadurl=pafbaseurl+"/ajax/downloadCompletedReport";
	     
    	   
    	 //  Header[] initpage=new BasicHeader[1];
    	  // initpage[0]=new BasicHeader("Host",loginhost);
    	   
	     //  HttpResponse loginlocation=HttpClientUtils.sendGet(pafbaseurl, initpage);
	     //  Header[] redirectheader=loginlocation.getAllHeaders();
	    //   String loginurl=HttpClientUtils.getContent(loginlocation);
	       // this is hte login form header
	       Multimap<String, String> userform=ArrayListMultimap.create();
	       userform.put("SMENC", "ISO-8859-1");
	       userform.put("postpreservationdata", "");
	       userform.put("target", pafrunidurl);
	       userform.put("smauthreason", "0");
	       userform.put("smtryno", "");
	       userform.put("smretries", "1");
	       userform.put("bacmonitoring","LOG ON");
	       userform.put("USER",username);
	       userform.put("PASSWORD", password);
	      
	       // this will navigate to the run assessment detail page
	      HttpResponse loginresponse=HttpClientUtils.sendPost(loginurl, userform, null);
	      //String smhauthr=HttpClientUtils.getContent(loginresponse);
	   
	       //get the side reminder cookies in this page    
	       Header[] login_responseHeaders = loginresponse.getHeaders("Set-Cookie");
			if(login_responseHeaders == null || login_responseHeaders.length <= 0){
				System.out.println("the request set cookies is null");
				System.exit(1);
			}
			Header[] login_responseCookies = new BasicHeader[login_responseHeaders.length];
			for(int i=0;i<login_responseHeaders.length;i++){
				login_responseCookies[i] = new BasicHeader("Cookie",login_responseHeaders[i].getValue());
			}
			System.out.println("1),get the login interface set cookies is: "+java.util.Arrays.toString(login_responseCookies));
	     
		   //get the run assessment detail page request
		   Header[] buildupheader=HttpClientUtils.buildRequestHeader(loginurl, login_responseCookies);
	       HttpResponse rundetailget=HttpClientUtils.sendGet(pafbaseurl, buildupheader);
	       String detailcontent=HttpClientUtils.getContent(rundetailget);
	       //System.out.println("home page content  is:"+home_responsecontent);
	       
	       Header[] header_detailresponse = rundetailget.getHeaders("Set-Cookie");
			if(header_detailresponse == null || header_detailresponse.length <= 0){
					System.out.println("the response assessment detail page  set_cookies header  is null");
					System.exit(1);
			}
			
			// this is the most important cookies we need to carried in the blow steps
			Header[] cookies_detailcookies=new BasicHeader[1];
			cookies_detailcookies[0]=new BasicHeader("Cookie",header_detailresponse[1].getValue());
		
			System.out.println("2),get the site minder cookies we need  is: "+java.util.Arrays.toString(cookies_detailcookies));
	      
			
		   Header[] header_startdownloader=HttpClientUtils.buildRequestHeader(pafrunidurl, cookies_detailcookies);
		   HttpResponse response_rundetail=HttpClientUtils.sendGet(pafrunidurl, header_startdownloader);
		   String content_rundetail=HttpClientUtils.getContent(response_rundetail);
		 //  System.out.println("get the run id assessment page content is:"+runidcontent);
		   
	       //get the listed report ids       
	       String reportidpattern="\\{\"saDocKey.*\\}\\];";
	       RegExpTester tester=new RegExpTester(reportidpattern, true);
	       String sadocKey=tester.validateString(content_rundetail);
	       System.out.println("Get the report ID  JSON String is:"+sadocKey);
	       
	     
	       //build the run form data,submit by ajax	
	       Multimap<String, String> multimap = ArrayListMultimap.create();
	       JSONArray sadocKeys=JSONArray.fromObject(sadocKey);
	       int reports=sadocKeys.size();
	       for(int i=0;i<reports;i++){
	    	   JSONObject sadocreport=sadocKeys.getJSONObject(i);
	    	   String downloadreportid=sadocreport.getString("saDocKey");
	    	   System.out.println("put the report id is:"+downloadreportid);
	    	   multimap.put("reportIds", downloadreportid);
	       }
	       
	       //get the customer name
	       String customernamepattern="(\\{id:downloadId.*customerName:'.*.'\\};)";
	       RegExpTester customertester=new RegExpTester(customernamepattern, true);
	       String ajaxdata=customertester.validateString(content_rundetail);
	       int firstindex=ajaxdata.indexOf("$.ajax({");
	       String firstdata=ajaxdata.substring(1, firstindex-2) .trim();
	       String seconddata=firstdata.substring(0, firstdata.length()-1);
	       
	     //  JSONObject customerinfo=JSONObject.fromObject(seconddata);
	       int startassesindex=seconddata.indexOf("assesName:'");
	       int endassessindex=seconddata.indexOf("', customerName:'");
	       int startcustomerindex= seconddata.indexOf("'}");
	       
	       
	       String assessmentname=seconddata.substring(startassesindex+11,endassessindex);
	      // assessmentname=assessmentname.replaceAll(" ", "_");
	       String customername=seconddata.substring(endassessindex+17,startcustomerindex-1);
	       String replacedcustomername= customername.replaceAll(" ", "_");
	       
	       System.out.println("Get the assessment name is:"+assessmentname);
	       System.out.println("Get the customer name is:"+replacedcustomername);
	       //get the guid
	       String pattern="xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx";
	       for(int k=0;k<pattern.length()-1;k++){
	           pattern=pattern.replaceFirst("x", getguid());
	           pattern=pattern.replaceFirst("y", getguid());
	        }
	       System.out.println("get the report GUID is:"+pattern);
	       multimap.put("id", pattern);
	       multimap.put("reasonCode", "DT");
	       multimap.put("analysisRunId",runid);
	       
	       multimap.put("assesName", assessmentname);
	       multimap.put("customerName",customername);
	       
	      
	       
	       System.out.println("the ajax download need to carried the data map is:"+multimap);
	       
	       Header[] smhcookiesheader=HttpClientUtils.buildRequestHeader(pafrunidurl, cookies_detailcookies);
	     //  Thread.sleep(10000);
	       HttpResponse download=HttpClientUtils.sendPost(startdownloadurl, multimap, smhcookiesheader);
	 	//  Thread.sleep(10000);    
	       boolean isfinished=false;
	       int downloadtime=0;
	       
	       Header[] completedcookiesheader=null;
	       while(!isfinished){
	    	   downloadtime=downloadtime+1;
		       String downloadverify=HttpClientUtils.getContent(download);
		       System.out.println("retry "+downloadtime+" time to startdownload json return content is:"+downloadverify);
		       JSONObject successmsg=JSONObject.fromObject(downloadverify);
		       isfinished=(Boolean) successmsg.get("success");
		       
		       Header[] downloadcookiesheader = download.getHeaders("Set-Cookie");
		       System.out.println("4) get the start completed set_cookies response is: "+java.util.Arrays.toString(downloadcookiesheader));
				if(downloadcookiesheader == null || downloadcookiesheader.length <= 0){
					System.out.println("the start download set cookies is null");
					System.exit(1);
				}
				completedcookiesheader= new BasicHeader[downloadcookiesheader.length];
				for(int i=0;i<downloadcookiesheader.length;i++){
					completedcookiesheader[i] = new BasicHeader("Cookie",downloadcookiesheader[i].getValue());
				}
				System.out.println("5) get the download completed cookies is: "+java.util.Arrays.toString(completedcookiesheader));
	       }
	       // build the refer header
	     
	       Header[] completedownloadheader= HttpClientUtils.buildRequestHeader(pafrunidurl, cookies_detailcookies);
	       System.out.println("the download refer header is: "+java.util.Arrays.toString(completedownloadheader));
	       
	       Multimap<String, String> downloadzipguid = ArrayListMultimap.create();
	       downloadzipguid.put("id",pattern);
	       HttpResponse zipfile= HttpClientUtils.sendPost(completedownloadurl, downloadzipguid, completedownloadheader);
	       
	       
	      
	       String zipfilename=assessmentname+"-"+replacedcustomername+"_"+GetCurrentTime.getTime()+".zip";
	       File downloadfolder=new File(downloaddir);
	       if(!downloadfolder.exists()){
	    	   System.out.println("We found the download folder not existing:"+downloaddir+",we will create it");
	    	   downloadfolder.mkdirs();
	       }
	       String filePath =downloaddir+zipfilename;
	       zipfile.addHeader("Content-Disposition", "attachment;filename="+zipfilename);
	       zipfile.addHeader("Content-Transfer-Encoding", "binary");
	       zipfile.addHeader("Content-Type", ".zip;charset=UTF-8");
	       zipfile.addHeader("Transfer-Encoding", "chunked");
	      
	     //  System.out.println("reponse code is:"+zipfile.getStatusLine().getStatusCode());
	     //  System.out.println("response download content is:"+HttpClientUtils.getContent(zipfile));
	     //  zipfile.getEntity().writeTo(new java.io.FileOutputStream(filePath));
	       File downloadzipfile=new File(filePath);
	      
	       if(downloadzipfile.createNewFile()){
		        System.out.println("File is created!");
		      }else{
		        System.out.println("File already exists.");
		      }
	       FileOutputStream fos = new FileOutputStream(downloadzipfile);
	       InputStream  in=zipfile.getEntity().getContent();
	       
	       byte[] buffer = new byte[4096];
           int len1 = 0;
           while ((len1 = in.read(buffer)) != -1) {
                   fos.write(buffer, 0, len1);
                   System.out.println("Now is writing the data with the buffer: "+len1+" byte...");
           }
           in.close();
           fos.flush();
           fos.close();
           System.out.println("Download the Assessment report: "+filePath+" Finished and Successfully !");

	       
	}
   	
	private static long copyLarge(InputStream input, OutputStream output)
            throws IOException {
        byte[] buffer = new byte[4096];
        long count = 0L;
        int n = 0;
        while (-1 != (n = input.read(buffer))) {
            output.write(buffer, 0, n);
            count += n;
        }
        return count;
    }

     /*
      * get the guid for the PAF report	
      */
	 public static String getguid(){
		   
		   int rnd =new Random().nextInt(60 - 1) + 1|0,v=1=='x'?rnd:rnd&0x3|0x8;
			String replacestr= Integer.toHexString(v);
			//System.out.println("random number is:"+replacestr);
			return replacestr;
	   }


}
