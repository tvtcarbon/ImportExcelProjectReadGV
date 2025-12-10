package de.srs.importExcelMRP.app;

import java.util.concurrent.atomic.AtomicBoolean;

import jakarta.ejb.Lock;
import jakarta.ejb.LockType;
import jakarta.ejb.Singleton;

@Singleton
public class WorkerBean {
    private AtomicBoolean busy = new AtomicBoolean(false);

    @Lock(LockType.READ)
    public void doTimerWork() throws InterruptedException {
    	System.out.println("doTimerWork!");
        if (!busy.compareAndSet(false, true)) {
            return;
        }
        try {
            Thread.sleep(20000L);
        } finally {
            busy.set(false);
        }
    }
}
