package de.srs.importExcelMRP.app.timer;

import de.srs.importExcelMRP.enums.TimerType;

public interface TimerScheduler {

	  void run(TimerType type, boolean forceRun);

	  boolean isTimerConfigured(TimerType type);
}
