package entities;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.InputMismatchException;

/**
 * Created by pgordon on 27.02.2017.
 */
public class Project {
    public Project(String name, String filename) {
        this.name = name;
        this.filename = filename;
        revenueMonthly = new ArrayList<>();
    }

    // Integer - costcenter for this project (project can include multiple costcenters)
    private HashMap<Integer, ArrayList<Double>> costCenterMonthly = new HashMap<Integer, ArrayList<Double>>();
    private ArrayList<Double> revenueMonthly = new ArrayList<>();
    private ArrayList<Double> chargingMonthly = new ArrayList<>();
    private ArrayList<Double> travelMonthly = new ArrayList<>();
    private String name;
    private String filename;
    private Double revenueTotal;

    public String getName() {
        return name;
    }

    public String getFilename() {
        return filename;
    }

    public ArrayList<Double> getRevenueMonthly() {
        return revenueMonthly;
    }

    public ArrayList<Double> getChargingMonthly() {
        return chargingMonthly;
    }

    public Double getRevenueTotal() {
        return revenueTotal;
    }

    public void setRevenueTotal(Double revenueTotal) {
        this.revenueTotal = revenueTotal;
    }

    public ArrayList<Double> getTravelMonthly() {
        return travelMonthly;
    }

    public HashMap<Integer, ArrayList<Double>> getCostCenterMonthly() {
        return costCenterMonthly;
    }

    public void setCostCenterMonthly(HashMap<Integer, ArrayList<Double>> costCenterMonthly) {
        this.costCenterMonthly = costCenterMonthly;
    }

    @Override
    public String toString() {
        return "[name=" + name + ", filename=" + filename + "\nrevenueMonthly=" + revenueMonthly + "=revenueTotal" + revenueTotal + "\nchargingMonthly=" + chargingMonthly + "]";
    }

}
