package misuExcel;

import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.logging.*;

public class Log {
	
	static Logger logger = null;
	
	private static void init(){
		if(logger==null){
			logger = Logger.getLogger("excel");
	        FileHandler fileHandler;
			try {
				fileHandler = new FileHandler("./ExcelLog.txt");
				fileHandler.setLevel(Level.ALL);
			    fileHandler.setFormatter(new MyLogHander());
			    logger.addHandler(fileHandler);
			} catch (SecurityException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			} 
		}
	}
	
	static void init(String savePath){
		if(logger==null){
			logger = Logger.getLogger("excel");
	        FileHandler fileHandler;
			try {
				fileHandler = new FileHandler(savePath + "/ExcelTool文件/ExcelLog.txt");
				fileHandler.setLevel(Level.ALL);
			    fileHandler.setFormatter(new MyLogHander());
			    logger.addHandler(fileHandler);
			} catch (SecurityException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			} 
		}
	}
	
	
	static void info(String str){
		init();
		logger.info(str);
	}
	
	static void warm(String str){
		init();
		logger.warning(str);
	}
}

class MyLogHander extends Formatter { 
    @Override 
    public String format(LogRecord record) { 
    		Date date = new Date(record.getMillis());
    		DateFormat format  = new SimpleDateFormat("HH:mm:ss");
            return format.format(date) + " "+ record.getLevel()+ ": " + record.getMessage()+"\n"; 
    } 
} 