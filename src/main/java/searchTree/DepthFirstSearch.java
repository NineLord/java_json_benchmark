package searchTree;

import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;
import utils.GsonClass;
import utils.GsonType;

import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

public class DepthFirstSearch {

    public static void main(String[] args) {
        @NotNull final String rawTree = "{\"a\":2,\"b\":1,\"c\":[null,{\"d\":3}]}";
        @NotNull final Object tree = GsonClass.fromJson(rawTree);
        @Nullable final Object value = "a";

        System.out.println(DepthFirstSearch.run(tree, value));
    }

    public static boolean run(@Nullable Object node, @Nullable Object value) {
        @Nullable final GsonType nodeType = GsonType.getNoneLeafType(node);

        if (nodeType == null) {
            if (node == null)
                return value == null;
            else
                return node.equals(value);
        }

        switch (nodeType) {
            case Object: {
                //noinspection unchecked
                @NotNull final Map<String, Object> actualNode = (Map<String, Object>) node;
                for (@NotNull final Entry<String, Object> entry : actualNode.entrySet()) {
                    if (entry.getKey().equals(value) || DepthFirstSearch.run(entry.getValue(), value))
                        return true;
                }
                break;
            }
            case Array: {
                //noinspection unchecked
                @NotNull final List<Object> actualNode = (List<Object>) node;
                for (@NotNull final Object nodeValue : actualNode) {
                    if (DepthFirstSearch.run(nodeValue, value))
                        return true;
                }
                break;
            }
            default:
                throw new RuntimeException("Invalid node type: " + nodeType);
        }

        return false;
    }
}
