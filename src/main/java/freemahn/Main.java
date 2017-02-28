package freemahn;

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

public class Main {
    private static final String PREFIX = "C:\\texcel\\docs\\";
    private static HashMap<String, Project> projects = new HashMap<>();
    private static boolean debug = true;


    public static void main(String[] args) {


        List<String> files = new ArrayList<>();

        //TODO input via properties/gui
        files.add("Forecast_Project 1.xlsm");
        files.add("Forecast_Project 2.xlsx");
        files.add("Forecast_Project 3.xlsx");
        files
                .stream()
                .map(Main::parseForecastProject)
                .forEach(project -> projects.put(project.name, project));

        projects.forEach((k, v) -> System.out.println(v));
        fillRevenue();
    }

    /**
     *
     */

    private static void fillRevenue() {
        System.out.println("Parsing Forecast 2017 project world.xlsx: started");
        XSSFWorkbook workbook = null;
        try {
            workbook = new XSSFWorkbook(OPCPackage.open(PREFIX + "Forecast 2017 project world.xlsx"));
        } catch (InvalidFormatException | IOException e) {
            e.printStackTrace();
            return;
        }

        XSSFSheet revenueSheet = workbook.getSheet("Revenue");

        System.out.print("Parsing accrual positions...");
        int accrualsRow = 1;
        Row accruals = revenueSheet.getRow(accrualsRow);
        ArrayList<Integer> cols = new ArrayList<>();
        for (Cell cell : accruals) {
            System.out.print(".");
            if (cell.getStringCellValue().contains("Accruals"))
                cols.add(cell.getColumnIndex());
        }
        System.out.println("finished");
        System.out.println("Setting values..");


        Iterator<Row> rowIterator = revenueSheet.iterator();
        Row row = null;
        //skip 3  headers row
        rowIterator.next();
        rowIterator.next();
        rowIterator.next();
        while (rowIterator.hasNext()) {
            System.out.println("----------------------------------------------");
            Row currentRow = rowIterator.next();
            Cell currentCell = currentRow.getCell(0);

            if (currentCell == null) {
                //means we have already parsed all project lines(?)
                System.out.println("null row at index " + currentRow.getRowNum());
                break;
            }
            if (debug)
                System.out.println("currentCell " + Arrays.toString(currentCell.getStringCellValue().getBytes()));
            String projectName = currentCell.getStringCellValue();
            Project p = projects.get(projectName);
            //System.out.println("projectName " + projectName + "" + Arrays.toString(projectName.getBytes()));
            if (p == null) {
                if (debug) System.out.println("No project " + projectName);
                continue;//TODO
            }

            int monthsColOffset = 4;
            for (int i = 0; i < 12; i++) {
                double monthValue = p.totalWithTravel.get(i);
                int colIndex = cols.get(i);
                Cell cell = currentRow.getCell(colIndex);
                if (debug) System.out.print("was " + cell.getNumericCellValue());
                cell.setCellValue(monthValue);
                if (debug) System.out.println(",become " + cell.getNumericCellValue());
            }


        }
        try (FileOutputStream out =
                     new FileOutputStream(new File("output\\test.xlsx"));) {
            //recalculate all formulas
            XSSFFormulaEvaluator.evaluateAllFormulaCells(workbook);
            workbook.write(out);
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("Parsing Forecast 2017 project world.xlsx: finished");

    }

    private static void fillChargingProjects() {

    }

    private static Project parseForecastProject(String filename) {
        XSSFWorkbook workbook = null;
        //parsing Forecast_PROJECTNAME.xslx
        String projectName = filename.substring("Forecast_".length(), filename.lastIndexOf("."));
        Project project = new Project(projectName, filename);

        int employeeRowOffset = 2;
        int monthsColOffset = 3;
        int rowHeaderColumn = 1;
        try (FileInputStream file = new FileInputStream(new File(PREFIX + project.filename));
        ) {
            //Get the workbook instance for XLS file
            workbook = new XSSFWorkbook(OPCPackage.open(file));
            XSSFSheet workHoursSheet = workbook.getSheet("Work hours");
            Iterator<Row> rowIterator = workHoursSheet.iterator();

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
                Cell monthValue = totalWithTravelRow.getCell(i);
                project.totalWithTravel.add(monthValue.getNumericCellValue());
            }
            Cell total = totalWithTravelRow.getCell(monthsColOffset + 12);
            project.total = total.getNumericCellValue();

        } catch (InvalidFormatException | IOException e) {
            e.printStackTrace();
        }
        return project;
    }

}
