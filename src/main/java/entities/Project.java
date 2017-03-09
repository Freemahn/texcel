package entities;

import java.util.ArrayList;

/**
 * Created by pgordon on 27.02.2017.
 */
public class Project {
    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getFilename() {
        return filename;
    }

    public void setFilename(String filename) {
        this.filename = filename;
    }

    public ArrayList<Double> getRevenueMonthly() {
        return revenueMonthly;
    }

    public void setRevenueMonthly(ArrayList<Double> revenueMonthly) {
        this.revenueMonthly = revenueMonthly;
    }

    public ArrayList<Double> getChargingMonthly() {
        return chargingMonthly;
    }

    public void setChargingMonthly(ArrayList<Double> chargingMonthly) {
        this.chargingMonthly = chargingMonthly;
    }

    public Double getRevenueTotal() {
        return revenueTotal;
    }

    public void setRevenueTotal(Double revenueTotal) {
        this.revenueTotal = revenueTotal;
    }

    private String name;
    private String filename;
    private ArrayList<Double> revenueMonthly = new ArrayList<>();
    private ArrayList<Double> chargingMonthly = new ArrayList<>();
    private Double revenueTotal;


    public Project(String name, String filename, ArrayList<Double> revenueMonthly) {
        this.name = name;
        this.filename = filename;
        this.revenueMonthly = revenueMonthly;

    }

    Project(String name, String filename) {
        this.name = name;
        this.filename = filename;
        revenueMonthly = new ArrayList<>();
    }

    @Override
    public String toString() {
        return "[name=" + name + ", filename=" + filename + "\nrevenueMonthly=" + revenueMonthly + "=revenueTotal" + revenueTotal + "\nchargingMonthly=" + chargingMonthly + "]";
    }
}
