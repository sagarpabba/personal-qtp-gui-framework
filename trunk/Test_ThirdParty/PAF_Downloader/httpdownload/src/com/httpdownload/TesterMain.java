package com.httpdownload;

import java.util.Random;

public class TesterMain {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		
		String customer="{id:downloadId, reasonCode:reasonCode, analysisRunId:analysisRunId, reportIds:reportIds, assesName:'PRO_SCAN', customerName:'PAF validation5'};";
		String customernamepattern="\\{id:downloadId.*customerName:'.*.'\\};$";
		RegExpTester customertester=new RegExpTester(customernamepattern, true);
	       String ajaxdata=customertester.validateString(customer);
	       System.out.println(ajaxdata);
   }
}
