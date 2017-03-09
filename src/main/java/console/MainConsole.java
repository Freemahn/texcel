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
import java.time.LocalDateTime;
import java.util.*;

public class MainConsole {
    private static final String PREFIX = "C:\\texcel\\docs1\\";
    private static HashMap<String, Project> projects = new HashMap<>();
    private static boolean debug = true;


    public static void main(String[] args) {

        String projectWorldFile;
        List<String> files = new ArrayList<>();
        if(args.length  == 0){
            projectWorldFile = "C:\\texcel\\docs1\\Forecast 2017 project world.xlsx";
            files.add("Forecast_Project 1.xlsm");
            files.add("Forecast_Project 2.xlsx");
            files.add("Forecast_Project 3.xlsx");
        }else{
            projectWorldFile = args[0];
            files.addAll(Arrays.asList(args));
            files.remove(0);//remove projectWorldFile
        }

        //parse projects
        //TODO input via properties/gui

        files
                .stream()
                .map((v) -> parseForecastProject(PREFIX,v))
                .forEach(project -> projects.put(project.getName(), project));

        projects.forEach((k, v) -> System.out.println(v));




        //fill revenue
        XSSFWorkbook workbook;
        try {
            workbook = new XSSFWorkbook(OPCPackage.open(PREFIX +"Forecast 2017 project world.xlsx"));
        } catch (InvalidFormatException | IOException e) {
            System.err.println("Cannot open forecast project world");
            //e.printStackTrace();
            return;
        }
        LocalDateTime timePoint = LocalDateTime.now();
        int currentMonth = timePoint.getMonthValue()-1;
        XSSFSheet revenueSheet = workbook.getSheet("Revenue");
        XSSFSheet chargingSheet = workbook.getSheet("Charging_projects");
        fillRevenue(revenueSheet, currentMonth);
        fillChargingProjects(chargingSheet,currentMonth);

        //fill chargingMonthly
        try (FileOutputStream out =
                     new FileOutputStream(new File("output\\test.xlsx"))) {

            //recalculate all formulas
            XSSFFormulaEvaluator.evaluateAllFormulaCells(workbook);
            workbook.write(out);
        } catch (IOException e) {
            e.printStackTrace();

        }
        System.out.println("Work is done, time to get some tea");
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

            String projectName = currentCell.getStringCellValue();
            Project project = projects.get(projectName);

            if (project == null) {
                //skip row
                //if (debug) System.out.println("No project " + projectName);
                continue;//TODO
            }
            System.out.println("projectName " + projectName);
            int monthsColOffset = 4;
            for (int i = currentMonth; i < 12; i++) {
                double monthValue = project.getChargingMonthly().get(i);

                Cell cell = currentRow.getCell(monthsColOffset+i);
                //if (debug) System.out.print("was " + cell.getNumericCellValue());
                cell.setCellValue(monthValue);
            }
            System.out.println("----------------------------------------------");
        }
        System.out.println("Parsing projects world: charging finished");


    }

    public static Project parseForecastProject(String path, String filename) {
        XSSFWorkbook workbook;
        //parsing Forecast_PROJECTNAME.xslx
        String projectName = filename.substring("Forecast_".length(), filename.lastIndexOf("."));
        Project project = new Project(projectName, filename);

        int rowHeaderColumn = 1;
        try (FileInputStream file = new FileInputStream(new File(path + project.getFilename()))) {
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
        while (rowIterator.hasNext() && totalWithTravelRow == null) {
            Row currentRow = rowIterator.next();
            Cell currentCell = currentRow.getCell(rowHeaderColumn);
            //System.out.println(currentCell);
            if (currentCell != null && currentCell.getStringCellValue().equals("Total with travel")) {
                totalWithTravelRow = currentRow;
            }
        }
        if (totalWithTravelRow == null) {
            System.err.println("Not found \"Total with travel\" line");
            return null;
        }
        for (int i = monthsColOffset; i < monthsColOffset + 12; i++) {
            Double monthValue = totalWithTravelRow.getCell(i).getNumericCellValue();
            project.getRevenueMonthly().add(monthValue);
        }
        Cell total = totalWithTravelRow.getCell(monthsColOffset + 12);
        project.setRevenueTotal(total.getNumericCellValue());

        XSSFSheet chargingSheet = workbook.getSheet("Charging");
        //search for "TOTAL" row
        Row totalRow = null;
        Iterator<Row> chargingIterator = chargingSheet.iterator();
        while (chargingIterator.hasNext() && totalRow == null) {
            Row currentRow = chargingIterator.next();
            Cell currentCell = currentRow.getCell(0);
            //System.out.println(currentCell);
            if (currentCell != null && currentCell.getStringCellValue().equals("TOTAL")) {
                totalRow = currentRow;
            }
        }
        monthsColOffset = 4;
        for (int i = monthsColOffset; i < monthsColOffset + 12; i++) {
            Double monthValue = totalRow.getCell(i).getNumericCellValue();
            project.getChargingMonthly().add(monthValue);
        }

        return project;
    }

}
