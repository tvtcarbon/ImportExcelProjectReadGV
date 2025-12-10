package de.srs.importExcelMRP.app.timer;

import java.util.Calendar;

import de.srs.importExcelMRP.enums.TimerType;
import jakarta.ejb.Schedule;
import jakarta.ejb.Timer;

public class TimerSchedulerImpl implements TimerScheduler {

	public boolean isDayOfWeek(final int dayOfWeek) {
		// support for "dayOfWeek" in "@Schedule" is buggy
		// so we check the day of week programmatically!
		Calendar calendar = Calendar.getInstance();
		return calendar.get(Calendar.DAY_OF_WEEK) == dayOfWeek;
	}

	public boolean isSaturday() {
		return isDayOfWeek(Calendar.SATURDAY);
	}

	public boolean isSunday() {
		return isDayOfWeek(Calendar.SUNDAY);
	}

	@Schedule(persistent = false, hour = "6", minute = "0")
	public void scheduleToIslaBonita(Timer timer) {
		if (isSaturday()) {
			run(TimerType.AT_6_AM, false);
		}
	}

	@Override
	public void run(TimerType type, boolean forceRun) {
		try {
			switch (type) {
			case AT_6_AM:
				break;
			default:
				break;
			}
		} catch (Exception e) {
		}
	}

	@Override
	public boolean isTimerConfigured(TimerType type) {
		// TODO Auto-generated method stub
		return false;
	}

}
