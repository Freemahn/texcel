package freemahn;

import java.util.ArrayList;

/**
 * Created by pgordon on 27.02.2017.
 */
public class Project {
    public String name;
    public String filename;
    public ArrayList<Double> revenueMonthly = new ArrayList<>();
    public ArrayList<Double> chargingMonthly = new ArrayList<>();
    public Double revenueTotal;

    public Project(String name, String filename, ArrayList<Double> revenueMonthly) {
        this.name = name;
        this.filename = filename;
        this.revenueMonthly = revenueMonthly;

    }

    public Project(String name, String filename) {
        this.name = name;
        this.filename = filename;
        revenueMonthly = new ArrayList<>();
    }

    @Override
    public String toString() {
        return "[name=" + name + ", filename=" + filename + "\nrevenueMonthly=" + revenueMonthly + "=revenueTotal" + revenueTotal + "\nchargingMonthly=" + chargingMonthly + "]";
    }
}
