package freemahn;

import java.util.ArrayList;

/**
 * Created by pgordon on 27.02.2017.
 */
public class Project {
    String name;
    String filename;
    ArrayList<Double> totalWithTravel;
    double total;

    public Project(String name, String filename, ArrayList<Double> totalWithTravel) {
        this.name = name;
        this.filename = filename;
        this.totalWithTravel = totalWithTravel;

    }

    public Project(String name, String filename) {
        this.name = name;
        this.filename = filename;
        totalWithTravel = new ArrayList<>();
    }

    @Override
    public String toString() {
        return "[name=" + name + ", filename=" + filename + ", total=" + totalWithTravel + "=" + total +"]";
    }
}
