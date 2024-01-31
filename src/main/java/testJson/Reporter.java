package testJson;

import org.jetbrains.annotations.NotNull;

import java.util.HashMap;
import java.util.Map;
import java.util.Map.Entry;
import java.util.stream.Stream;

public class Reporter {

    private final Map<String, Long> measurementStartTime;
    private final Map<String, Long> measurementDuration;

    //#region Constructor
    public Reporter() {
        this.measurementStartTime = new HashMap<>();
        this.measurementDuration = new HashMap<>();
    }
    //#endregion

    public @NotNull Stream<Entry<String, Long>> getMeasurementDuration() {
        return measurementDuration.entrySet().stream();
    }

    //#region Generic Measurement
    public void startMeasuring(@NotNull final MeasurementType measurementType) {
        this.measurementStartTime.put(measurementType.name(), System.nanoTime());
    }

    public void finishMeasuring(@NotNull final MeasurementType measurementType) {
        @NotNull final String measurementName = measurementType.name();
        final long durationNano = System.nanoTime() - this.measurementStartTime.get(measurementName);
        final long duration = durationNano / 1_000_000;
        this.measurementDuration.put(measurementName, duration);
    }
    //#endregion

    public enum MeasurementType {
        GENERATE_JSON,
        DESERIALIZE_JSON,
        ITERATE_ITERATIVELY,
        ITERATE_RECURSIVELY,
        SERIALIZE_JSON
    }

}
