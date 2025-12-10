package de.srs.importExcelMRP.app;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.SQLException;
import java.time.LocalDate;
import java.time.ZoneId;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.Locale;
import java.util.Timer;

import de.srs.importExcelMRP.Exception.ErrorReadExcelException;
import jakarta.ejb.EJB;
import jakarta.ejb.Lock;
import jakarta.ejb.LockType;
import jakarta.ejb.Schedule;

public class TimerEx {
    @EJB
    private WorkerBean workerBean;

    @Lock(LockType.READ)
    @Schedule(second = "*/5", minute = "*", hour = "*", persistent = false)
    public void atSchedule() throws InterruptedException {
        workerBean.doTimerWork();
    }
    /*
    private final static int SIX_AM = 6;
    private final static int THIRDY_MINUTES = 0;

	public static void main(String[] args) throws ErrorReadExcelException, FileNotFoundException, IOException {
		Timer timer = new Timer();
		
		Date date = new Date();
		LocalDate localDate = date.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
		int year  = localDate.getYear();
		int month = localDate.getMonthValue();
		int day   = localDate.getDayOfMonth();		
		
		
		Calendar date1 = Calendar.getInstance();
        date1.set(year, month-1, day, SIX_AM, THIRDY_MINUTES, 0);
        timer.schedule(new ImportExcelMRPTask(), date1.getTime(), 86400000);
		
	}
	*/
}