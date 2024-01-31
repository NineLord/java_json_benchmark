package utils;

import com.google.gson.Gson;
import com.google.gson.GsonBuilder;
import com.google.gson.ToNumberPolicy;
import com.google.gson.internal.LinkedTreeMap;
import org.jetbrains.annotations.NotNull;

import java.util.Map;

public class GsonClass {

    private static final Gson parser = new GsonBuilder()
            .setObjectToNumberStrategy(ToNumberPolicy.LONG_OR_DOUBLE)
            .create();

    public static Map<String, Object> fromJson(@NotNull final String json) {
        //noinspection unchecked
        return (Map<String, Object>) parser.fromJson(json, LinkedTreeMap.class);
    }

    public static String toJson(@NotNull final Object json) {
        return parser.toJson(json);
    }
}
