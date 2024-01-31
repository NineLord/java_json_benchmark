package searchTree;

import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;
import utils.GsonClass;
import utils.GsonType;

import java.util.List;
import java.util.Map;
import java.util.Stack;

public class BreadthFirstSearch {

    public static void main(String[] args) {
        @NotNull final String rawTree = "{\"a\":2,\"b\":1,\"c\":[null,{\"d\":3}]}";
        @NotNull final Object tree = GsonClass.fromJson(rawTree);
        @Nullable final Object value = 3L;

        System.out.println(BreadthFirstSearch.run(tree, value));
    }

    public static boolean run(@NotNull Object tree, @Nullable Object value) {
        @NotNull Stack<Object> currentNodes = new Stack<>();
        currentNodes.push(tree);
        @NotNull Stack<Object> nextLevelNodes = new Stack<>();
        boolean isValueString = value instanceof String;

        while (!currentNodes.isEmpty()) {
            @Nullable final Object currentNode = currentNodes.pop();

            @Nullable GsonType currentNodeType = GsonType.getNoneLeafType(currentNode);
            if (currentNodeType == null) {
                if (currentNode == null) {
                    if (value == null)
                        return true;
                } else {
                    if (currentNode.equals(value))
                        return true;
                }
            } else {
                switch (currentNodeType) {
                    case Object: {
                        //noinspection unchecked
                        @NotNull final Map<String, Object> actualCurrentNode = (Map<String, Object>) currentNode;
                        if (isValueString) {
                            if (actualCurrentNode.containsKey(value))
                                return true;
                        }
                        nextLevelNodes.addAll(actualCurrentNode.values());
                        break;
                    }
                    case Array: {
                        //noinspection unchecked
                        @NotNull final List<Object> actualCurrentNode = (List<Object>) currentNode;
                        nextLevelNodes.addAll(actualCurrentNode);
                        break;
                    }
                    default:
                        throw new RuntimeException("Invalid node type: " + currentNodeType);
                }
            }

            if (currentNodes.isEmpty()) {
                currentNodes = nextLevelNodes;
                nextLevelNodes = new Stack<>();
            }
        }

        return false;
    }
}
