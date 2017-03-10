package console;

import entities.Project;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

/**
 * Created by prozhdes on 08.03.2017.
 */

public class Executor {
    private static HashMap<String, Project> projects;
    private static boolean debug = true;


    public static int parse(List<String> files, File destinationFile, int currentMonth) {
        projects = new HashMap<>();
        String projectWorldFile = destinationFile.getAbsolutePath();
        String name[] = destinationFile.getName().split("\\.(?=[^\\.]+$)");
        String temp_name = name[0] + "_temp." + name[1];
        String destinationFolder = destinationFile.getParent() + File.separator;
        //parse projects
        files.stream()
                .map((v) -> parseForecastProject(v))
                .forEach(project -> projects.put(project.getName(), project));

        projects.forEach((k, v) -> System.out.println(v));

        //fill revenue
        XSSFWorkbook workbook;
        try {
            workbook = new XSSFWorkbook(OPCPackage.open(projectWorldFile));
        } catch (InvalidFormatException | IOException e) {
            System.err.println("Cannot open forecast project world");
            //e.printStackTrace();
            return -1;
        }

        XSSFSheet revenueSheet = workbook.getSheet("Revenue");
        XSSFSheet chargingSheet = workbook.getSheet("Charging_projects");
        XSSFSheet costCentersSheet = workbook.getSheet("Charging_CostCenters");

        fillRevenue(revenueSheet, currentMonth);
        fillChargingProjects(chargingSheet,currentMonth);
        fillCostCenters(costCentersSheet,currentMonth);

        //fill chargingMonthly
        try (FileOutputStream out =
                     new FileOutputStream(new File(destinationFolder + temp_name))) {

            //recalculate all formulas
            XSSFFormulaEvaluator.evaluateAllFormulaCells(workbook);
            workbook.write(out);
        } catch (IOException e) {
            e.printStackTrace();
        }
        try {
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        delete(projectWorldFile);
        rename(destinationFolder + temp_name, projectWorldFile);
        return 0;
    }

    private static void fillCostCenters(XSSFSheet chargingSheet, int currentMonth) {
        System.out.println("Parsing projects world: cost centers started");
        System.out.println("Setting values..");
        Iterator<Row> rowIterator = chargingSheet.iterator();
        Row row = null;
        //skip 3  headers row
        rowIterator.next();
        rowIterator.next();
        rowIterator.next();
        while (rowIterator.hasNext()) {
            Row currentRow = rowIterator.next();
            Cell currentCell = currentRow.getCell(0);

            if (currentCell == null) {
                //means we have already parsed all project lines
                break;
            }

            //System.out.println(currentCell.getCellTypeEnum().name());
            String projectName = currentCell.getStringCellValue();

            Project project = projects.get(projectName);

            if (project == null) {
                //skip row
                //if (debug) System.out.println("No project " + projectName);
                continue;//TODO
            }
            System.err.println("projectName " + projectName);

            int costcenter = (int) currentRow.getCell(3).getNumericCellValue();
            ArrayList<Double> values = project.getCostCenterMonthly().get(costcenter);


            int monthsColOffset = 4;
            for (int i = currentMonth; i < 12; i++) {
                double monthValue = values.get(i);

                Cell cell = currentRow.getCell(monthsColOffset + i);
                cell.setCellValue(monthValue);
                //if (debug) System.out.println(",become " + cell.getNumericCellValue());
            }
            System.out.println("----------------------------------------------");
        }
        System.out.println("Parsing projects world: costcenters finished");

    }

    /**
     *
     */
    private static ArrayList<Integer> parseAccrualMonths(XSSFSheet revenueSheet) {
        System.out.print("Parsing accrual positions...");
        Row accrualsRow = revenueSheet.getRow(1);
        ArrayList<Integer> cols = new ArrayList<>();
        for (Cell cell : accrualsRow) {
            System.out.print(".");
            if (cell.getStringCellValue().contains("Accruals"))
                cols.add(cell.getColumnIndex());
        }
        System.out.println("finished");
        return cols;
    }

    public static boolean delete(String filepath) {
        File file = null;
        try{
            file = new File(filepath);
        }catch(Exception e){
            e.printStackTrace();
        }
        return file.delete();
    }

    public static boolean rename(String tempName, String name) {
        File tempFile = new File(tempName);
        File file = new File(name);
        return tempFile.renameTo(file);
    }

    private static void fillRevenue(XSSFSheet revenueSheet, int currentMonth) {
        System.out.println("Parsing projects world: revenue started");
        ArrayList<Integer> cols = parseAccrualMonths(revenueSheet);
        System.out.println("Setting values..");

        Iterator<Row> rowIterator = revenueSheet.iterator();
        Row row = null;
        //skip 3  headers row
        rowIterator.next();
        rowIterator.next();
        rowIterator.next();
        while (rowIterator.hasNext()) {

            Row currentRow = rowIterator.next();
            Cell currentCell = currentRow.getCell(0);

            if (currentCell == null) {
                //means we have already parsed all project lines
                break;
            }

            String projectName = currentCell.getStringCellValue();
            Project project = projects.get(projectName);

            if (project == null) {
                //skip row
                //if (debug) System.out.println("No project " + projectName);
                continue;//TODO
            }
            System.out.println("projectName " + projectName);
            for (int i = currentMonth; i < 12; i++) {
                double monthValue = project.getRevenueMonthly().get(i);
                int colIndex = cols.get(i);
                Cell cell = currentRow.getCell(colIndex);
                if (debug) System.out.print("was " + cell.getNumericCellValue());
                cell.setCellValue(monthValue);
                if (debug) System.out.println(",become " + cell.getNumericCellValue());
            }
            System.out.println("----------------------------------------------");
        }
        System.out.println("Parsing projects world: revenue finished");

    }

    private static void fillChargingProjects(XSSFSheet chargingSheet,int currentMonth) {
        System.out.println("Parsing projects world: charging started");
        System.out.println("Setting values..");
        Iterator<Row> rowIterator = chargingSheet.iterator();
        Row row = null;
        //skip 3  headers row
        rowIterator.next();
        rowIterator.next();
        rowIterator.next();
        while (rowIterator.hasNext()) {

            Row currentRow = rowIterator.next();
            Cell currentCell = currentRow.getCell(0);

            if (currentCell == null) {
                //means we have already parsed all project lines
                break;
            }

            System.out.println(currentCell.getCellTypeEnum().name());
            String projectName = currentCell.getStringCellValue();

            Project project = projects.get(projectName);

            if (project == null) {
                //skip row
                //if (debug) System.out.println("No project " + projectName);
                continue;//TODO
            }
            System.err.println("projectName " + projectName);
            int monthsColOffset = 4;
            for (int i = currentMonth; i < 12; i++) {
                double monthValue = project.getChargingMonthly().get(i);

                Cell cell = currentRow.getCell(monthsColOffset+i);
                //if (debug) System.out.print("was " + cell.getNumericCellValue());
                cell.setCellValue(monthValue);
                if (debug) System.out.println(",become " + cell.getNumericCellValue());
            }
            System.out.println("----------------------------------------------");
        }
        System.out.println("Parsing projects world: charging finished");


    }

    public static Project parseForecastProject(String filePath) {
        XSSFWorkbook workbook;
        //parsing Forecast_PROJECTNAME.xslx
        int ind = filePath.lastIndexOf(File.separator);
        String fileName = filePath.substring(ind + 1, filePath.lastIndexOf('.')); //filename.substring("Forecast_".length(), filename.lastIndexOf("."));
        String projectName = fileName.replace("Forecast_", "");
        Project project = new Project(projectName, filePath);

        int rowHeaderColumn = 1;
        try (FileInputStream file = new FileInputStream(new File(project.getFilename()))) {
            //Get the workbook instance for XLS file
            workbook = new XSSFWorkbook(OPCPackage.open(file));
        } catch (InvalidFormatException | IOException e) {
            e.printStackTrace();
            return null;
        }

        //parse revenue sheet
        XSSFSheet workHoursSheet = workbook.getSheet("Work hours");
        Iterator<Row> rowIterator = workHoursSheet.iterator();
        int monthsColOffset = 3;
        //search for "Total with travel" row
        Row totalWithTravelRow = null;
        Row travel = null;
        while (rowIterator.hasNext()) {
            Row currentRow = rowIterator.next();
            Cell currentCell = currentRow.getCell(rowHeaderColumn);
            //System.out.println(currentCell);
            if (currentCell != null && currentCell.getStringCellValue().equals("Total with travel")) {
                totalWithTravelRow = currentRow;
            }
            if (currentCell != null && currentCell.getStringCellValue().equals("Travel")) {
                travel = currentRow;
            }
        }
        if (totalWithTravelRow == null) {
            System.err.println("Not found \"Total with travel\" line");
            return null;
        }
        for (int i = monthsColOffset; i < monthsColOffset + 12; i++) {
            Double monthValue = totalWithTravelRow.getCell(i).getNumericCellValue();
            project.getRevenueMonthly().add(monthValue);
            monthValue = travel.getCell(i).getNumericCellValue();
            project.getTravelMonthly().add(monthValue);
        }

        Cell total = totalWithTravelRow.getCell(monthsColOffset + 12);
        project.setRevenueTotal(total.getNumericCellValue());
        Cell totalTravel = travel.getCell(monthsColOffset + 12);
        project.setTravelTotal(total.getNumericCellValue());

        for (int i = monthsColOffset; i < monthsColOffset + 12; i++) {
            Double monthValue = totalWithTravelRow.getCell(i).getNumericCellValue();
            project.getRevenueMonthly().add(monthValue);
        }
        XSSFSheet chargingSheet = workbook.getSheet("Charging");
        //search for "TOTAL" and "CC" (CostCenter) row
        Row totalRow = null;
        List<Row> ccRows = new ArrayList<Row>();
        Iterator<Row> chargingIterator = chargingSheet.iterator();
        while (chargingIterator.hasNext()) {
            Row currentRow = chargingIterator.next();

            Cell currentCell = currentRow.getCell(0);


            if (currentCell != null && currentCell.getStringCellValue().equals("TOTAL")) {
                totalRow = currentRow;
            }
            if (currentCell != null && currentCell.getStringCellValue().contains("CC")) {
                ccRows.add(currentRow);
            }
        }
        monthsColOffset = 4;
        for (int i = monthsColOffset; i < monthsColOffset + 12; i++) {
            Double monthValue = totalRow.getCell(i).getNumericCellValue();
            project.getChargingMonthly().add(monthValue);
        }

        HashMap<Integer, ArrayList<Double>> map = new HashMap<Integer, ArrayList<Double>>();
        for (Row row : ccRows) { // for each costcenter
            Integer costcenter;
            //System.err.println(" == ccs " + row.getPhysicalNumberOfCells());
            costcenter = (int) row.getCell(1).getNumericCellValue();
            monthsColOffset = 4;
            ArrayList<Double> values = new ArrayList<>();
            for (int i = monthsColOffset; i < monthsColOffset + 12; i++) {
                Double monthValue = row.getCell(i).getNumericCellValue();
                //System.err.println(" == val " + monthValue);
                values.add(monthValue); //project.getCostCenterMonthly().put(costcenter, monthValue);
            }
            map.put(costcenter, values);
        }
        System.err.println(" == ccs " + map.keySet());
        project.setCostCenterMonthly(map);

        return project;
    }
}
