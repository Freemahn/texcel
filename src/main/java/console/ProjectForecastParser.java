package console;

import entities.Project;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;

/**
 * Created by pgordon on 28.02.2017.
 */
public class ProjectForecastParser {


    private HashMap<String, Project> projects = new HashMap<>();
    public ProjectForecastParser(){


    }
    public void parseForecastProjects(List<File> files){
        files.forEach(this::parseForecastProject);

    }
    public void parseForecastProjects(File... files){
        Arrays.stream(files).forEach(this::parseForecastProject);

    }
    public void parseForecastProject(File file) {
        XSSFWorkbook workbook ;
        //parsing Forecast_PROJECTNAME.xslx
        String filename = file.getName();
        String projectName = filename.substring("Forecast_".length(), filename.lastIndexOf("."));
        Project project = new Project(projectName, filename);

        int rowHeaderColumn = 1;
        try (FileInputStream file1 = new FileInputStream(file)) {
            //Get the workbook instance for XLS file
            workbook = new XSSFWorkbook(OPCPackage.open(file1));
        } catch (InvalidFormatException | IOException e) {
            e.printStackTrace();
            return;
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
            return;
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
        projects.put(project.getName(), project);
    }

}
