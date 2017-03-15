package main;

import entities.Project;
import gui.GUI;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;
import java.util.List;
import java.util.function.BiPredicate;
import java.util.function.Predicate;

import static main.Utils.*;

/**
 * Created by prozhdes on 08.03.2017.
 */

public class Executor {
    private HashMap<String, Project> projects;
    private static boolean debug = true;
    static CellStyle style;

    public void parse(List<String> files, File destinationFile, int currentMonth) throws Exception {
       // try {
        projects = new HashMap<>();
        String projectWorldFile = destinationFile.getAbsolutePath();
        String name[] = destinationFile.getName().split("\\.(?=[^\\.]+$)");
        String temp_name = name[0] + "_temp." + name[1];
        String destinationFolder = destinationFile.getParent() + File.separator;

        for (String path : files) {
            Project p = parseForecastProject(path);
            projects.put(p.getName(), p);
        }

        //fill revenue
        XSSFWorkbook workbook = null;
        try {
            workbook = new XSSFWorkbook(OPCPackage.open(projectWorldFile));
            style = workbook.createCellStyle();
            style.setFillForegroundColor(IndexedColors.AQUA.index);
            style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        } catch (InvalidFormatException | IOException e) {
           // return "Cannot open forecast project world";
        }

            XSSFSheet revenueSheet = workbook.getSheet("Revenue");
            XSSFSheet chargingSheet = workbook.getSheet("Charging_projects");
            XSSFSheet costCentersSheet = workbook.getSheet("Charging_CostCenters");
            XSSFSheet projectsTravelSheet = workbook.getSheet("Projects_Travel");

            fillRevenue(revenueSheet, currentMonth);
            fillChargingProjects(chargingSheet, currentMonth);
            fillCostCenters(costCentersSheet, currentMonth);
            fillTravel(projectsTravelSheet, currentMonth);

            //fill chargingMonthly
            try (FileOutputStream out =
                         new FileOutputStream(new File(destinationFolder + temp_name))) {

                //recalculate all formulas
                XSSFFormulaEvaluator.evaluateAllFormulaCells(workbook);
                workbook.write(out);
            } catch (IOException e) {
                e.printStackTrace();
            }
            workbook.close();
            delete(projectWorldFile);
            rename(destinationFolder + temp_name, projectWorldFile);
            Desktop dt = Desktop.getDesktop();
            dt.open(new File(projectWorldFile));
       // } catch (Exception e) {
       //     return e.getMessage();
       // }
       // return "finished";
    }

    private void fillCostCenters(XSSFSheet chargingSheet, int currentMonth) {
        Iterator<Row> rowIterator = chargingSheet.iterator();
        Row row = null;
        //skip 3  headers row
        rowIterator.next();
        rowIterator.next();
        rowIterator.next();
        while (rowIterator.hasNext()) {
            Row currentRow = rowIterator.next();
            String content = getCellContent(currentRow.getCell(0));
            if (content.isEmpty())
                break;
            Project project = projects.get(content);
            if (project == null) {
                //skip row
                continue;//TODO
            }
            int costcenter = (int) Double.parseDouble(getCellContent(currentRow.getCell(3)));//int) currentRow.getCell(3).getNumericCellValue();
            ArrayList<Double> values = project.getCostCenterMonthly().get(costcenter);

            int monthsColOffset = 4;
            for (int i = currentMonth; i < 12; i++) {


                double monthValue = values.get(i);

                Cell cell = currentRow.getCell(monthsColOffset + i);
                cell.setCellValue(monthValue);
                cell.setCellStyle(style);
            }
        }

    }

    /**
     *
     */
    private ArrayList<Integer> parseAccrualMonths(XSSFSheet revenueSheet) {
        Row accrualsRow = revenueSheet.getRow(1);
        ArrayList<Integer> cols = new ArrayList<>();
        for (Cell cell : accrualsRow) {
            if (cell.getStringCellValue().contains("Accruals"))
                cols.add(cell.getColumnIndex());
        }
        return cols;
    }

    private void fillRevenue(XSSFSheet revenueSheet, int currentMonth) {
        ArrayList<Integer> cols = parseAccrualMonths(revenueSheet);
        Iterator<Row> rowIterator = revenueSheet.iterator();
        Row row = null;
        //skip 3  headers row
        rowIterator.next();
        rowIterator.next();
        rowIterator.next();
        while (rowIterator.hasNext()) {
            Row currentRow = rowIterator.next();
            String content = getCellContent(currentRow.getCell(0));
            if (content.isEmpty())
                break;
            Project project = projects.get(content);

            if (project == null) {
                //skip row
                continue;//TODO
            }
            for (int i = currentMonth; i < 12; i++) {
                double monthValue = project.getRevenueMonthly().get(i);
                int colIndex = cols.get(i);
                Cell cell = currentRow.getCell(colIndex);
                cell.setCellValue(monthValue);
                cell.setCellStyle(style);

            }
        }
    }

    private void fillTravel(XSSFSheet travelSheet, int currentMonth) {
        Iterator<Row> rowIterator = travelSheet.iterator();
        Row row = null;
        //skip 3  headers row
        rowIterator.next();
        rowIterator.next();
        rowIterator.next();

        while (rowIterator.hasNext()) {
            Row currentRow = rowIterator.next();
            //System.out.println(" == " + getCellContent(currentRow.getCell(0)));
            String content = getCellContent(currentRow.getCell(0));
            if (content.isEmpty())
                break;
            Project project = projects.get(content);

            if (project == null) {
                //skip row
                continue;//TODO
            }
            for (int i = currentMonth + 4; i < 16; i++) {
                double monthValue = project.getTravelMonthly().get(i - 4);
                Cell cell = currentRow.getCell(i);
                cell.setCellValue(monthValue);
                cell.setCellStyle(style);

            }
        }
    }

    private void fillChargingProjects(XSSFSheet chargingSheet,int currentMonth) {
        Iterator<Row> rowIterator = chargingSheet.iterator();
        Row row = null;
        //skip 3  headers row
        rowIterator.next();
        rowIterator.next();
        rowIterator.next();
        while (rowIterator.hasNext()) {
            Row currentRow = rowIterator.next();
            String content = getCellContent(currentRow.getCell(0));
            if (content.isEmpty())
                break;
            Project project = projects.get(content);

            if (project == null) {
                //skip row
                continue;//TODO
            }
            int monthsColOffset = 4;
            for (int i = currentMonth; i < 12; i++) {
                double monthValue = project.getChargingMonthly().get(i);
                Cell cell = currentRow.getCell(monthsColOffset+i);
                cell.setCellValue(monthValue);
                cell.setCellStyle(style);
            }
        }
    }

    public Project parseForecastProject(String filePath) throws Exception {
        XSSFWorkbook workbook;
        //parsing Forecast_PROJECTNAME.xslx
        int ind = filePath.lastIndexOf(File.separator);
        String fileName = filePath.substring(ind + 1, filePath.lastIndexOf('.'));
        String projectName = fileName.replace("Forecast_", "");
        Project project = new Project(projectName, filePath);
        try (FileInputStream file = new FileInputStream(new File(project.getFilename()))) {
            //Get the workbook instance for XLS file
            workbook = new XSSFWorkbook(OPCPackage.open(file));
        } catch (InvalidFormatException | IOException e) {
            throw new Exception(project.getFilename() + " has wrong format");
        }

        XSSFSheet workHoursSheet = workbook.getSheet("Work hours");
        XSSFSheet chargingSheet = workbook.getSheet("Charging");
        if (workHoursSheet == null || chargingSheet == null)
            throw new Exception("No such tab in " + projectName + " file");

        Row totalWithTravel = getRowStartsWith(workHoursSheet, "Total with travel");
        Row travel = getRowStartsWith(workHoursSheet, "Travel");

        Row total = getRowStartsWith(chargingSheet, "TOTAL");
        List<Row> CC = getRowsStartsWith(chargingSheet, "CC");
        if (totalWithTravel == null || travel == null || total == null || CC.size() == 0)
            throw new Exception("No such string in " + projectName + " file");

        for (int i = 2; i <= 13; i++) {
            Double monthValue = getThCell(travel, i).getNumericCellValue();
            project.getTravelMonthly().add(monthValue);

            monthValue = getThCell(totalWithTravel, i).getNumericCellValue();
            project.getRevenueMonthly().add(monthValue);

            monthValue = getThCell(total, i).getNumericCellValue();
            project.getChargingMonthly().add(monthValue);
        }

        HashMap<Integer, ArrayList<Double>> map = new HashMap<Integer, ArrayList<Double>>();
        for (Row row : CC) { // for each costcenter
            Integer center = (int) getThCell(row, 2).getNumericCellValue();
            ArrayList<Double> values = new ArrayList<>();
            for (int i = 3; i <= 14; i++) {
                Double monthValue = getThCell(row, i).getNumericCellValue();
                values.add(monthValue);
            }
            map.put(center, values);
        }
        project.setCostCenterMonthly(map);

        return project;
    }
}
