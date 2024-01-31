package testJson.multithreading;

import com.google.gson.Gson;
import jsonGenerator.Generator;
import org.jetbrains.annotations.NotNull;
import searchTree.BreadthFirstSearch;
import searchTree.DepthFirstSearch;
import testJson.Reporter;
import utils.GsonClass;

import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Map;

public class RunTestLoop implements Runnable {

    @NotNull private static final String characterPoll = "AaBbCcDdEeFfGgHhIiJjKkLlMmNnOoPpQqRrSsTtUuVvWwXxYyZz!@#$%&";

    private final int testCounter;
    @NotNull private final PcUsageExporter pcUsageExporter;
    @NotNull private final String jsonPath;
    private final int numberOfLetters;
    private final int depth;
    private final int numberOfChildren;

    public RunTestLoop(int testCounter, @NotNull final PcUsageExporter pcUsageExporter, @NotNull final String jsonPath, int numberOfLetters, int depth, int numberOfChildren) {
        this.testCounter = testCounter;
        this.pcUsageExporter = pcUsageExporter;
        this.jsonPath = jsonPath;
        this.numberOfLetters = numberOfLetters;
        this.depth = depth;
        this.numberOfChildren = numberOfChildren;
    }

    @Override
    public void run() {
        try {
            @NotNull final String rawJson = new String(Files.readAllBytes(Paths.get(this.jsonPath)));

            for (int count = 0; count < this.testCounter; ++count) {
                Reporter reporter = new Reporter();

                reporter.startMeasuring(Reporter.MeasurementType.GENERATE_JSON);
                Generator.generateJson(characterPoll, this.numberOfLetters, this.depth, this.numberOfChildren);
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

                reporter.getMeasurementDuration();
                this.pcUsageExporter.getUsages();
            }
        } catch (Exception exception) {
            System.err.println("Test failed: " + exception);
        }
    }

}
