package utils;

import org.jetbrains.annotations.Nullable;

public class MathDataCollector {

    @Nullable private Double minimum;
    @Nullable private Double maximum;
    private double sum;
    private long count;

    public MathDataCollector() {
        this.minimum = null;
        this.maximum = null;
        this.sum = 0;
        this.count = 0;
    }

    public void add(final double data) {
        this.sum += data;
        ++this.count;

        this.minimum = this.minimum == null ? data : Math.min(this.minimum, data);
        this.maximum = this.maximum == null ? data : Math.max(this.maximum, data);
    }

    public @Nullable Double getMinimum() {
        return this.minimum;
    }

    public @Nullable Double getMaximum() {
        return this.maximum;
    }

    public double getSum() {
        return this.sum;
    }

    public long getCount() {
        return this.count;
    }

    public @Nullable Double getAverage() {
        if (this.count == 0)
            return null;
        else
            return this.sum / this.count;
    }
}
