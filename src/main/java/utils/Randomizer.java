package utils;

import org.jetbrains.annotations.NotNull;
import org.json.JSONException;
import org.json.JSONObject;

import java.util.Iterator;

public class Randomizer {

    private Randomizer() {}

    public static boolean getRandom() {
        return Randomizer.getRandom(0, 1) == 0;
    }

    public static int getRandom(int minimum, int maximum) {
        return (int) Randomizer.getRandom((long) minimum, maximum);
    }

    public static long getRandom(long minimum, long maximum) {
        return (long) Math.floor(Randomizer.getRandom((double) minimum, maximum));
    }

    public static double getRandom(double minimum, double maximum) {
        return (Math.random() * (maximum + 1 - minimum)) + minimum;
    }

    public static <T> T getRandomValueFromArray(T[] array) {
        if (array.length == 0)
            throw new IndexOutOfBoundsException("Can't get random value from empty array");

        int index = Randomizer.getRandom(0, array.length - 1);
        return array[index];
    }

    public static char getRandomValueFromArray(char[] array) {
        if (array.length == 0)
            throw new IndexOutOfBoundsException("Can't get random value from empty array");

        int index = Randomizer.getRandom(0, array.length - 1);
        return array[index];
    }

    public static Object getRandomValueFromJSONObject(JSONObject object) throws JSONException {
        int index = getRandom(0, object.length());

        //noinspection unchecked
        Iterator<String> keys = object.keys();
        for (int count = 0; keys.hasNext(); ++count) {
            @NotNull final String key = keys.next();

            if (count == index)
                return object.get(key);
        }

        throw new RuntimeException("Can't get random value from empty object");
    }
}
