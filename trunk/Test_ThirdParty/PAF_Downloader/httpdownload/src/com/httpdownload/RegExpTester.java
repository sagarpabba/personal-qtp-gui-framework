package com.httpdownload;

import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class RegExpTester{
	 
	  private Pattern pattern;
	  private Matcher matcher;

	 

	  public RegExpTester(String patterns){
		  pattern = Pattern.compile(patterns);
	  }
	  public RegExpTester(String patterns,boolean mulitlyline){
		  pattern = Pattern.compile(patterns,pattern.DOTALL);
	  }
	 

	  /**
	   * Validate password with regular expression
	   * @param password password for validation
	   * @return true valid password, false invalid password
	   */
	  public boolean validate(final String inputtester){

		  matcher = pattern.matcher(inputtester);
		  return matcher.matches();

	  }
	  public String validateString(final String inputtester){
		  matcher = pattern.matcher(inputtester);
		  List<String> matchedstr=new ArrayList<String>();
		  while(matcher.find()){
			 // matcher.
			  System.out.println(matcher.groupCount());
			   matchedstr.add(matcher.group());
			 //  System.out.println("Matched content is:"+matcher.group());
		  }
		  
		  return matchedstr.toString();
	  }
}