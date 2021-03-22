import java.util.ArrayList;

public class ExelDocInList {
    String name;
    ArrayList<ExelSheet> Sheets = new ArrayList<>();

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public ArrayList<ExelSheet> getSheets() {
        return Sheets;
    }

    public void setSheets(ArrayList<ExelSheet> sheets) {
        Sheets = sheets;
    }
}