package utils;

import org.jetbrains.annotations.NotNull;
import org.json.JSONArray;
import org.json.JSONObject;

public enum JSONType {

    Null,
    Boolean,
    Long,
    Double,
    String,
    Array,
    Object;

    @NotNull private static final char[] alphabet = "AaBbCcDdEeFfGgHhIiJjKkLlMmNnOoPpQqRrSsTtUuVvWwXxYyZz".toCharArray();
    @NotNull private static final JSONType[] leafJSONTypes = new JSONType[]{Null, Boolean, Long, Double, String};
    @NotNull private static final JSONType[] noneLeafJSONTypes = new JSONType[]{Array, Object};

    public static @NotNull JSONType getRandomLeafJsonType() {
        return Randomizer.getRandomValueFromArray(leafJSONTypes);
    }

    public static @NotNull JSONType getRandomNoneLeafJsonType() {
        return Randomizer.getRandomValueFromArray(noneLeafJSONTypes);
    }

    public static @NotNull Object getRandomNodeValue(@NotNull final JSONType type) {
        switch (type) {
            case Null:
                return JSONObject.NULL;
            case Boolean:
                return Randomizer.getRandom();
            case Long:
                return Randomizer.getRandom(-1_000_000_000L, 1_000_000_000L);
            case Double:
                return Randomizer.getRandom(-1_000_000_000d, 1_000_000_000d);
            case String: {
                final int stringLength = Randomizer.getRandom(0, 32);
                @NotNull final StringBuilder stringBuilder = new StringBuilder();

                for (int count = 0; count < stringLength; ++count)
                    stringBuilder.append(Randomizer.getRandomValueFromArray(JSONType.alphabet));

                return stringBuilder.toString();
            }
            case Array:
                return new JSONArray();
            case Object:
                return new JSONObject();
            default:
                throw new RuntimeException("Invalid JSONType: " + type);
        }
    }
}
