package testJson;

import com.google.common.util.concurrent.ThreadFactoryBuilder;
import jsonGenerator.Generator;
import org.jetbrains.annotations.NotNull;
import searchTree.BreadthFirstSearch;
import searchTree.DepthFirstSearch;
import testJson.multithreading.PcUsageExporter;
import utils.GsonClass;

import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Map;
import java.util.concurrent.ScheduledThreadPoolExecutor;
import java.util.concurrent.TimeUnit;

public class Main {

    @NotNull private static final String characterPoll = "AaBbCcDdEeFfGgHhIiJjKkLlMmNnOoPpQqRrSsTtUuVvWwXxYyZz!@#$%&";

    // Add the following arguments to VM options: `-Xmx8G -Xms16M -XX:MinHeapFreeRatio=10 -XX:MaxHeapFreeRatio=20`

    public static void main(String[] args) throws Exception {
        //#region Test input
        final int testCounter = Integer.parseInt(args[1]);
        @NotNull final PcUsageExporter pcUsageExporter = new PcUsageExporter();
        @NotNull final String jsonPath = args[0];
        final int numberOfLetters = Integer.parseInt(args[3]);
        final int depth = Integer.parseInt(args[4]);
        final int numberOfChildren = Integer.parseInt(args[5]);
        final int samplingInterval = Integer.parseInt(args[6]);
        //#endregion

        //#region Getting ready for testing
        @NotNull final String rawJson = new String(Files.readAllBytes(Paths.get(jsonPath)));
        ScheduledThreadPoolExecutor scheduler = new ScheduledThreadPoolExecutor(1,
                new ThreadFactoryBuilder().setNameFormat("PcUsageExporter-%d").build());
        scheduler.scheduleAtFixedRate(pcUsageExporter, 0, samplingInterval, TimeUnit.MILLISECONDS);
        //#endregion

        //#region Testing
        try (@NotNull final ExcelGenerator excelGenerator = new ExcelGenerator(
                args[2],
                jsonPath,
                samplingInterval,
                numberOfLetters,
                depth,
                numberOfChildren
        )) {
            for (int count = 0; count < testCounter; ++count) {
                Reporter reporter = new Reporter();

                reporter.startMeasuring(Reporter.MeasurementType.GENERATE_JSON);
                Generator.generateJson(characterPoll, numberOfLetters, depth, numberOfChildren);
                reporter.finishMeasuring(Reporter.MeasurementType.GENERATE_JSON);

                reporter.startMeasuring(Reporter.MeasurementType.DESERIALIZE_JSON);
                @NotNull final Map<String, Object> json = GsonClass.fromJson(rawJson);
                reporter.finishMeasuring(Reporter.MeasurementType.DESERIALIZE_JSON);

                final long valueToSearch = 2_000_000_000L;

                reporter.startMeasuring(Reporter.MeasurementType.ITERATE_ITERATIVELY);
                if (BreadthFirstSearch.run(json, valueToSearch))
                    throw new RuntimeException("BFS the tree found value that shouldn't be in it: " + valueToSearch);
                reporter.finishMeasuring(Reporter.MeasurementType.ITERATE_ITERATIVELY);

                reporter.startMeasuring(Reporter.MeasurementType.ITERATE_RECURSIVELY);
                if (DepthFirstSearch.run(json, valueToSearch))
                    throw new RuntimeException("DFS the tree found value that shouldn't be in it: " + valueToSearch);
                reporter.finishMeasuring(Reporter.MeasurementType.ITERATE_RECURSIVELY);

                reporter.startMeasuring(Reporter.MeasurementType.SERIALIZE_JSON);
                GsonClass.toJson(json);
                reporter.finishMeasuring(Reporter.MeasurementType.SERIALIZE_JSON);

                excelGenerator.appendWorksheet("Test " + (count + 1), reporter.getMeasurementDuration(), pcUsageExporter.getUsages());
            }
        }
        //#endregion

        scheduler.shutdown();
    }

}
