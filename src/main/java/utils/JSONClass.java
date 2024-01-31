package utils;

import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;

public class JSONClass {

    private @Nullable JSONObject object;
    private @Nullable JSONArray array;
    @SuppressWarnings("NotNullFieldNotInitialized")
    private @NotNull JSONType type;

    public JSONClass(@NotNull JSONObject object) {
        this.setObject(object);
    }

    public JSONClass(@NotNull JSONArray array) {
        this.setArray(array);
    }

    //#region Getters
    public @Nullable JSONObject getObject() {
        return object;
    }

    public @Nullable JSONArray getArray() {
        return array;
    }

    public @NotNull JSONType getType() {
        return type;
    }
    //#endregion

    //#region Setters
    public void setObject(@NotNull JSONObject object) {
        this.object = object;
        this.array = null;
        this.type = JSONType.Object;
    }

    public void setArray(@Nullable JSONArray array) {
        this.object = null;
        this.array = array;
        this.type = JSONType.Array;
    }
    //#endregion

    //#region Class Functionality
    public boolean getBoolean(@NotNull String key) throws JSONException, NumberFormatException {
        switch (this.type) {
            case Object:
                //noinspection ConstantConditions
                return this.object.getBoolean(key);
            case Array:
                int index = Integer.parseInt(key);
                //noinspection ConstantConditions
                return this.array.getBoolean(index);
            default:
                throw new RuntimeException("Unknown type: " + this.type);
        }
    }
    public double getDouble(@NotNull String key) throws JSONException, NumberFormatException {
        switch (this.type) {
            case Object:
                //noinspection ConstantConditions
                return this.object.getDouble(key);
            case Array:
                int index = Integer.parseInt(key);
                //noinspection ConstantConditions
                return this.array.getDouble(index);
            default:
                throw new RuntimeException("Unknown type: " + this.type);
        }
    }
    public int getInt(@NotNull String key) throws JSONException, NumberFormatException {
        switch (this.type) {
            case Object:
                //noinspection ConstantConditions
                return this.object.getInt(key);
            case Array:
                int index = Integer.parseInt(key);
                //noinspection ConstantConditions
                return this.array.getInt(index);
            default:
                throw new RuntimeException("Unknown type: " + this.type);
        }
    }
    public long getLong(@NotNull String key) throws JSONException, NumberFormatException {
        switch (this.type) {
            case Object:
                //noinspection ConstantConditions
                return this.object.getLong(key);
            case Array:
                int index = Integer.parseInt(key);
                //noinspection ConstantConditions
                return this.array.getLong(index);
            default:
                throw new RuntimeException("Unknown type: " + this.type);
        }
    }
    public @NotNull Object get(@NotNull String key) throws JSONException, NumberFormatException {
        switch (this.type) {
            case Object:
                //noinspection ConstantConditions
                return this.object.get(key);
            case Array:
                int index = Integer.parseInt(key);
                //noinspection ConstantConditions
                return this.array.get(index);
            default:
                throw new RuntimeException("Unknown type: " + this.type);
        }
    }
    public @NotNull String getString(@NotNull String key) throws JSONException, NumberFormatException {
        switch (this.type) {
            case Object:
                //noinspection ConstantConditions
                return this.object.getString(key);
            case Array:
                int index = Integer.parseInt(key);
                //noinspection ConstantConditions
                return this.array.getString(index);
            default:
                throw new RuntimeException("Unknown type: " + this.type);
        }
    }
    public @NotNull JSONObject getJSONObject(@NotNull String key) throws JSONException, NumberFormatException {
        switch (this.type) {
            case Object:
                //noinspection ConstantConditions
                return this.object.getJSONObject(key);
            case Array:
                int index = Integer.parseInt(key);
                //noinspection ConstantConditions
                return this.array.getJSONObject(index);
            default:
                throw new RuntimeException("Unknown type: " + this.type);
        }
    }
    public @NotNull JSONArray getJSONArray(@NotNull String key) throws JSONException, NumberFormatException {
        switch (this.type) {
            case Object:
                //noinspection ConstantConditions
                return this.object.getJSONArray(key);
            case Array:
                int index = Integer.parseInt(key);
                //noinspection ConstantConditions
                return this.array.getJSONArray(index);
            default:
                throw new RuntimeException("Unknown type: " + this.type);
        }
    }

    public @NotNull JSONClass getJSONClass(@NotNull String key) throws JSONException, NumberFormatException {
        switch (this.type) {
            case Object:
                try {
                    //noinspection ConstantConditions
                    return new JSONClass(this.object.getJSONObject(key));
                } catch (JSONException ignored) {}

                return new JSONClass(this.object.getJSONArray(key));
            case Array:
                int index = Integer.parseInt(key);
                try {
                    //noinspection ConstantConditions
                    return new JSONClass(this.array.getJSONObject(index));
                } catch (JSONException ignored) {}

                return new JSONClass(this.array.getJSONArray(index));
            default:
                throw new RuntimeException("Unknown type: " + this.type);
        }
    }

    public void set(@NotNull final String key, @NotNull final Object value) throws JSONException, NumberFormatException {
        switch (this.type) {
            case Object:
                //noinspection ConstantConditions
                this.object.put(key, value);
                break;
            case Array:
                int index = Integer.parseInt(key);
                //noinspection ConstantConditions
                this.array.put(index, value);
                break;
            default:
                throw new RuntimeException("Unknown type: " + this.type);
        }
    }
    public void set(final int index, @NotNull final Object value) throws JSONException, NumberFormatException {
        switch (this.type) {
            case Array:
                //noinspection ConstantConditions
                this.array.put(index, value);
                break;
            case Object:
                @NotNull final String key = Integer.toString(index);
                //noinspection ConstantConditions
                this.object.put(key, value);
                break;
            default:
                throw new RuntimeException("Unknown type: " + this.type);
        }
    }
    //#endregion

    @Override
    public String toString() {
        switch (this.type) {
            case Object:
                return "JSONClass" + this.object;
            case Array:
                return "JSONClass" + this.array;
            default:
                throw new RuntimeException("Unknown type: " + this.type);
        }
    }

}
