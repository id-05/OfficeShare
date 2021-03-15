import java.util.ArrayList;

public class ExelSheet {
    String name;
    ArrayList<ArrayList<String>> table = new ArrayList<>();

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public ArrayList<ArrayList<String>> getTable() {
        return table;
    }

    public void setTable(ArrayList<ArrayList<String>> table) {
        this.table = table;
    }
}
