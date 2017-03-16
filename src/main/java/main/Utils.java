package main;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.io.File;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.function.BiPredicate;

/**
 * Created by prozhdes on 14.03.2017.
 */
public class Utils {

    private static BiPredicate<String, String> sameContent = (stored, cell) -> stored.replaceAll(" ", "").equalsIgnoreCase(cell.replaceAll(" ", ""));


    public static Row getRowStartsWith(XSSFSheet from, String cellContent) {
        Iterator<Row> iterator = from.iterator();
        Row target = null;
        while (iterator.hasNext() && target == null) {
            Row currentRow = iterator.next();
            Cell first = getThCell(currentRow, 1);
            if (sameContent.test(getCellContent(first), cellContent))
                target = currentRow;
        }
        return target;
    }

    public static List<Row> getRowsStartsWith(XSSFSheet from, String cellContent) {
        Iterator<Row> iterator = from.iterator();
        List<Row> target = new ArrayList<Row>();
        while (iterator.hasNext()) {
            Row currentRow = iterator.next();
            Cell first = getThCell(currentRow, 1);
            if (sameContent.test(getCellContent(first), cellContent))
                target.add(currentRow);
        }
        return target;
    }

    public static int getIndexOfCell(Row row, int th) {
        int index = -1;
        int num = row.getPhysicalNumberOfCells();
        for (int i = 0; (i < num) && (th > 0); i++) {
            if (!getCellContent(row.getCell(i)).isEmpty()) {
                th--;
                index = i;
            }
        }
        return index;
    }

    public static Cell getThCell(Row row, int th) {
        Cell c = null;
        int num = row.getPhysicalNumberOfCells();
        for (int i = 0; (i < num) && (th > 0); i++) {
            if (!getCellContent(row.getCell(i)).isEmpty()) {
                th--;
                c = row.getCell(i);
            }
        }
        return c;
    }

    public static String getCellContent(Cell c) {
        if (c == null) {
            return "";
        }
        try {
            return c.getStringCellValue();
        } catch (Exception e) {
            return String.valueOf(c.getNumericCellValue());
        }
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
}
