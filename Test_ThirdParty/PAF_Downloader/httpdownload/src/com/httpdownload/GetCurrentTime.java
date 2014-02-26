package com.httpdownload;

import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Random;

public class GetCurrentTime {

	
	
	public static String getTime(){
		
		String returntime=new SimpleDateFormat("yyyy-MM-dd_HH-mm-ss").format(Calendar.getInstance().getTime());
		return returntime;
	}
	
	public static int randomnumber(int min,int max){
		
		   Random rand = new Random();

		    // nextInt is normally exclusive of the top value,
		    // so add 1 to make it inclusive
		    int randomNum = rand.nextInt((max - min) + 1) + min;

		    return randomNum;
	}
}
