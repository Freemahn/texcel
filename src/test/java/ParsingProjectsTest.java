import freemahn.Main;
import freemahn.Project;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Before;
import org.junit.Test;

import java.io.IOException;
import java.util.Arrays;
import java.util.List;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
public class ParsingProjectsTest {

    private Project actualProject;
    @Before
    public void openFile() throws InvalidFormatException, IOException{
        XSSFWorkbook workbook= new XSSFWorkbook(OPCPackage.open("test_files\\Forecast 2017 project world.xlsx"));
        actualProject =  Main.parseForecastProject("test_files\\", "Forecast_Project 1.xlsm");
        assertNotNull(actualProject);
    }
    @Test
    public void testParseProjectName() {
        assertEquals("Project 1", actualProject.getName());
    }
    @Test
    public void testTotalMonthly()  {
        Double [] revenues = {3265904.0, 3023520.0, 3788720.0, 3225200.0, 4105803.3949999996, 3350860.0,
                3766708.9299999997, 3482340.0, 4102796.1624999996, 3665160.0, 4268116.1625, 3429880.0};
        List expected = Arrays.asList(revenues);
        List<Double> actual = actualProject.getRevenueMonthly();
        assertEquals(expected, actual);
    }
    @Test
    public void testRevenueTotal()  {
        Double expected = 4.347500865E7;
        Double actual =  actualProject.getRevenueTotal();
        assertEquals(expected, actual);
    }

    @Test
    public void testChargingMonthly()  {
        Double [] revenues = {3146672.3500000006, 2913165.9199999995, 3650437.6, 3107484.7999999993, 3231545.5999999996,
                3228558.64, 3146294.9999999995, 3355240.400000001, 3349381.72, 3531387.5999999996, 3508667.7199999997, 3304694.92};
        List expected = Arrays.asList(revenues);
        List<Double> actual = actualProject.getChargingMonthly();
        assertEquals(expected, actual);
    }

}