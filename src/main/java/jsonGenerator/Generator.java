package jsonGenerator;

import org.jetbrains.annotations.NotNull;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;
import utils.JSONClass;
import utils.JSONType;
import utils.NotYetImplemented;
import utils.Randomizer;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Stack;

public class Generator {

    @SuppressWarnings("ConstantConditions")
    public static void main(String[] args) throws JSONException {
        @NotNull final String alphabet = "abcdefghijklmnopqrstuvwxyz";
        @NotNull final String specialCharacters = "!@#$%&";
        final int alphabetCount = 26;
        final int numberOfLetters = 8;
        final int depth = 10;
        final int numberOfChildren = 5;
        @NotNull final String pathToSaveFile = System.getProperty("user.home") + File.separator + "generatedJson.json";
        final boolean isSpecialCharacters = true;
        final boolean debug = false;

        JSONObject result = Generator.generateJson(
                (isSpecialCharacters ? specialCharacters : "") +
                        alphabet.substring(0, alphabetCount),
                numberOfLetters,
                depth,
                numberOfChildren
        );

        if (debug)
            System.out.println(result.toString(2));
        else {
            try (Writer writer = new BufferedWriter(new OutputStreamWriter(Files.newOutputStream(Paths.get(pathToSaveFile)), StandardCharsets.UTF_8))) {
                writer.write(result.toString());
            } catch (IOException exception) {
                throw new RuntimeException(exception);
            }
        }
    }

    private final char[] charactersPoll;
    private final int numberOfLetters;
    private final int depth;
    private final int numberOfChildren;

    //#region Constructor
    private Generator(@NotNull String charactersPoll, int numberOfLetters, int depth, int minimumChildren) {
        this.charactersPoll = charactersPoll.toCharArray();
        this.numberOfLetters = numberOfLetters;
        this.depth = depth;
        this.numberOfChildren = minimumChildren;
    }

    public static @NotNull JSONObject generateJson(@NotNull String charactersPoll, int numberOfLetters, int depth, int minimumChildren) throws JSONException {
        @NotNull final Generator generator = new Generator(charactersPoll, numberOfLetters, depth, minimumChildren);

        return generator.generateFullTree();
    }
    //#endregion

    public @NotNull JSONObject generateFullTree() throws JSONException {
        @NotNull final JSONObject result = new JSONObject();

        @NotNull Stack<JSONClass> currentNodes = new Stack<>();
        currentNodes.push(new JSONClass(result));
        @NotNull Stack<JSONClass> nextLevelNodes = new Stack<>();
        final int lastLevel = this.depth - 1;

        for (int level = 0; level < lastLevel; ++level) {

            while (!currentNodes.isEmpty()) {
                @NotNull final JSONClass currentNode = currentNodes.pop();
                @NotNull final JSONType currentNodeType = currentNode.getType();

                switch (currentNodeType) {
                    case Array: {
                        for (int nodeCount = 0; nodeCount < this.numberOfChildren; ++nodeCount) {
                            @NotNull final JSONType childNodeType = JSONType.getRandomNoneLeafJsonType();

                            switch (childNodeType) {
                                case Array: {
                                    @NotNull final JSONArray childNodeValue = new JSONArray();
                                    currentNode.set(nodeCount, childNodeValue);
                                    nextLevelNodes.push(new JSONClass(childNodeValue));
                                    break;
                                }
                                case Object: {
                                    @NotNull final JSONObject childNodeValue = new JSONObject();
                                    currentNode.set(nodeCount, childNodeValue);
                                    nextLevelNodes.push(new JSONClass(childNodeValue));
                                    break;
                                }
                                default:
                                    throw new RuntimeException("Invalid child node type: " + childNodeType);
                            }
                        }
                        break;
                    }
                    case Object: {
                        for (int nodeCount = 0; nodeCount < this.numberOfChildren; ++nodeCount) {
                            @NotNull final String childNodeName = this.getRandomNodeName();
                            @NotNull final JSONType childNodeType = JSONType.getRandomNoneLeafJsonType();

                            switch (childNodeType) {
                                case Array: {
                                    @NotNull final JSONArray childNodeValue = new JSONArray();
                                    currentNode.set(childNodeName, childNodeValue);
                                    nextLevelNodes.push(new JSONClass(childNodeValue));
                                    break;
                                }
                                case Object: {
                                    @NotNull final JSONObject childNodeValue = new JSONObject();
                                    currentNode.set(childNodeName, childNodeValue);
                                    nextLevelNodes.push(new JSONClass(childNodeValue));
                                    break;
                                }
                                default:
                                    throw new RuntimeException("Invalid child node type: " + childNodeType);
                            }
                        }
                        break;
                    }
                    default:
                        throw new RuntimeException("Invalid current node type: " + currentNodeType);
                }
            }

            currentNodes = nextLevelNodes;
            nextLevelNodes = new Stack<>();
        }

        if (this.depth > 0) {
            while (!currentNodes.isEmpty()) {
                @NotNull final JSONClass currentNode = currentNodes.pop();
                @NotNull final JSONType currentNodeType = currentNode.getType();

                switch (currentNodeType) {
                    case Array: {
                        for (int nodeCount = 0; nodeCount < this.numberOfChildren; ++nodeCount)
                            currentNode.set(nodeCount, JSONType.getRandomNodeValue(JSONType.getRandomLeafJsonType()));
                        break;
                    }
                    case Object: {
                        for (int nodeCount = 0; nodeCount < this.numberOfChildren; ++nodeCount)
                            currentNode.set(this.getRandomNodeName(), JSONType.getRandomNodeValue(JSONType.getRandomLeafJsonType()));
                        break;
                    }
                    default:
                        throw new RuntimeException("Invalid current node type: " + currentNodeType);
                }
            }
        }

        return result;
    }

    //#region Helper methods
    //#region Node Names
    private char getRandomNodeCharacter() {
        return Randomizer.getRandomValueFromArray(this.charactersPoll);
    }

    private @NotNull String getRandomNodeName() {
        @NotNull final StringBuilder stringBuilder = new StringBuilder();
        for (int count = 0; count < this.numberOfLetters; ++count)
            stringBuilder.append(this.getRandomNodeCharacter());

        return stringBuilder.toString();
    }
    //#endregion
    //#endregion
}
