package de.srs.importExcelMRP.app.utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.attribute.FileTime;
import java.util.Properties;
import java.util.TreeMap;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import de.srs.importExcelMRP.Exception.FileNotFound;
import de.srs.importExcelMRP.app.App;

public class Utils {
	static Logger logger = LogManager.getLogger(Utils.class);
	
	public static String readFileProperties() throws FileNotFoundException, IOException {
		
		
		
		String rootPath = Thread.currentThread().getContextClassLoader().getResource("").getPath();
		String appConfigPath = rootPath + "app.properties";
		logger.debug("appConfigPath:: "+appConfigPath);

		Properties appProps = new Properties();
		appProps.load(new FileInputStream(appConfigPath));

		logger.debug("pathExcelFiles:: "+ appProps.getProperty("pathExcelFiles"));

		return  appProps.getProperty("pathExcelFiles");
	}

	public static String recognizeLastFileFromDisk(String folderExcelFiles) throws FileNotFound {
		TreeMap<String, Long> files = new TreeMap<String, Long>();
		
		final File folder = new File(folderExcelFiles);
	    for (final File fileEntry : folder.listFiles()) {
	        if (fileEntry.isDirectory()) {
	        	recognizeLastFileFromDisk(folderExcelFiles);
	        } else {
	        	try {
	        		
	        	    FileTime creationTime = (FileTime) Files.getAttribute(fileEntry.getAbsoluteFile().toPath(), "creationTime");
	        	    files.put(fileEntry.getName(), creationTime.toMillis());
	        	    
	        	} catch (IOException ex) {
					throw new FileNotFound("Error when i try to recognizeLastFileFromDisk " + ex);
	        	}
	        }
	    }
	    logger.info("The last file is:: "+files.lastKey());
	    return  files.lastKey();
	}
}
