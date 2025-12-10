package de.srs.importExcelMRP.app;

import java.util.TimerTask;

public class ImportExcelMRPTask extends TimerTask {

	    @Override
	    public void run() {
	        long currennTime = System.currentTimeMillis();
	        long stopTime = currennTime + 2000;//provide the 2hrs time it should execute 1000*60*60*2
	          //while(stopTime != System.currentTimeMillis()){
	              // Do your Job Here
	            System.out.println("Start Job"+stopTime);
	            System.out.println("End Job"+System.currentTimeMillis());
	          //}
	    }


}
