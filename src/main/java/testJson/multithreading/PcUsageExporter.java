package testJson.multithreading;

import com.sun.management.OperatingSystemMXBean;
import org.apache.commons.lang3.tuple.Pair;
import org.jetbrains.annotations.NotNull;

import java.lang.management.ManagementFactory;
import java.util.LinkedList;
import java.util.List;

public class PcUsageExporter implements Runnable {

    @NotNull private static final OperatingSystemMXBean os = (OperatingSystemMXBean) ManagementFactory.getOperatingSystemMXBean();
    private static final Runtime runtime = Runtime.getRuntime();

    @SuppressWarnings("NotNullFieldNotInitialized")
    @NotNull private List<Pair<Long, Long>> usages;

    public PcUsageExporter() {
        this.init();
    }

    private void init() {
        this.usages = new LinkedList<>();
    }

    @NotNull public synchronized List<Pair<Long, Long>> getUsages() {
        @NotNull final List<Pair<Long, Long>> usages = this.usages;
        this.init();
        return usages;
    }

    @Override
    public synchronized void run() {
        final long cpuUsage = Math.round(os.getProcessCpuLoad() * 100);
        final long ramUsageBytes = runtime.totalMemory() - runtime.freeMemory();
        final long ramUsage = ramUsageBytes / (1024 * 1024);
        this.usages.add(Pair.of(cpuUsage, ramUsage));
    }

}
