package utils;

import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;

import java.util.List;
import java.util.Map;

public enum GsonType {

    Null,
    Boolean,
    Long,
    Double,
    String,
    Array,
    Object;

    public static @Nullable GsonType getNoneLeafType(@Nullable final Object element) {
        if (element instanceof Map)
            return GsonType.Object;
        else if (element instanceof List)
            return GsonType.Array;
        else
            return null;
    }
}
