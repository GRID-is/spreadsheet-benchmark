// The JVM's heap figures for a shim (such as replay/PoiEvaluate.java): call
// watchFullCollections() first, and memoryLine() while the workbook is still reachable. The line
//   #memory \t M \t <retained heap> <retained direct buffers> <peak live heap>
// gives the retained heap and peak live heap (CONTEXT.md), in MiB; replay/jvm.ts parses it.
import com.sun.management.GarbageCollectionNotificationInfo;
import java.lang.management.BufferPoolMXBean;
import java.lang.management.GarbageCollectorMXBean;
import java.lang.management.ManagementFactory;
import java.lang.management.MemoryPoolMXBean;
import java.lang.management.MemoryType;
import java.lang.management.MemoryUsage;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.atomic.AtomicLong;
import javax.management.NotificationEmitter;
import javax.management.openmbean.CompositeData;

public class JvmHeap {
    // The most heap in use after a full collection, in bytes.
    static final AtomicLong peakLiveHeap = new AtomicLong();

    // Tracks peakLiveHeap from the collectors' notifications. On G1, the default collector, a full
    // collection ("end of major GC") happens only under memory pressure or on System.gc(); its
    // young and mixed collections leave garbage in the old generation, so they are not counted.
    static void watchFullCollections () {
        Set<String> heapPools = new HashSet<>();
        for (MemoryPoolMXBean pool : ManagementFactory.getMemoryPoolMXBeans()) {
            if (pool.getType() == MemoryType.HEAP) heapPools.add(pool.getName());
        }
        for (GarbageCollectorMXBean collector : ManagementFactory.getGarbageCollectorMXBeans()) {
            if (!(collector instanceof NotificationEmitter)) continue;
            ((NotificationEmitter) collector).addNotificationListener((notification, handback) -> {
                if (!GarbageCollectionNotificationInfo.GARBAGE_COLLECTION_NOTIFICATION.equals(notification.getType())) return;
                GarbageCollectionNotificationInfo info = GarbageCollectionNotificationInfo.from((CompositeData) notification.getUserData());
                if (!"end of major GC".equals(info.getGcAction())) return;
                long used = 0;
                for (Map.Entry<String, MemoryUsage> pool : info.getGcInfo().getMemoryUsageAfterGc().entrySet()) {
                    if (heapPools.contains(pool.getKey())) used += pool.getValue().getUsed();
                }
                peakLiveHeap.accumulateAndGet(used, Math::max);
            }, null, null);
        }
    }

    static long mib (long bytes) {
        return Math.round(bytes / 1048576.0);
    }

    // Collects garbage and reads the heap. The forced collection's notification may arrive later,
    // so it is counted directly.
    static String memoryLine () {
        System.gc();
        long retainedHeap = ManagementFactory.getMemoryMXBean().getHeapMemoryUsage().getUsed();
        long retainedDirect = 0;
        for (BufferPoolMXBean pool : ManagementFactory.getPlatformMXBeans(BufferPoolMXBean.class)) {
            if ("direct".equals(pool.getName())) retainedDirect += pool.getMemoryUsed();
        }
        long peakLive = Math.max(peakLiveHeap.get(), retainedHeap);
        return "#memory\tM\t" + mib(retainedHeap) + ' ' + mib(retainedDirect) + ' ' + mib(peakLive) + '\n';
    }
}
