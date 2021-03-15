import java.util.ArrayList;

public class ExelSheet {
    String name;
    public static ArrayList<ArrayList<String>> table = new ArrayList<>();

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public static ArrayList<ArrayList<String>> getTable() {
        return table;
    }

    public static void setTable(ArrayList<ArrayList<String>> table) {
        ExelSheet.table = table;
    }
}
