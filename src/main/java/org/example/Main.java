package org.example;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Main {
    public static void main(String[] args) throws IOException {

        readExcel("src/main/resources/Modulhandbuecher/HDH-Informatik-Allgemeine Informatik.xls");
    }

    public static void readExcel(String fileLocation) throws IOException {

        FileInputStream file = new FileInputStream(new File(fileLocation));
        Workbook workbook = new HSSFWorkbook(file);

        Sheet sheet = workbook.getSheetAt(0);

        Map<Integer, List<String>> data = new HashMap<>();
        int i = 0;
        for (Row row : sheet) {

            data.put(i, new ArrayList<String>());
            for (Cell cell : row) {

                boolean isBlankCell = CellType.BLANK == cell.getCellType();
                boolean isEmptyStringCell = CellType.STRING == cell.getCellType() && cell.getStringCellValue().trim().isEmpty();

                if (isBlankCell || isEmptyStringCell) {
//                if (cell == null || cell.getCellType() == CellType.BLANK) {
                    System.out.print("blank; ");
                } else {

                    switch (cell.getCellType()) {
                        case STRING:
                            System.out.print(cell.getRichStringCellValue().getString() + ": " + cell.getCellType().toString() + ", "
                                    + "FontIndex: " + cell.getCellStyle().getFontIndex() + ", "
                                    + "FillBackgroundColor: " + cell.getCellStyle().getFillBackgroundColor()
                                    + "; ");
                            break;
                        case NUMERIC:
                            System.out.print(cell.getNumericCellValue() + ": " + cell.getCellType().toString() + "; "
                                    + "FontIndex: " + cell.getCellStyle().getFontIndex() + ", "
                                    + "FillBackgroundColor: " + cell.getCellStyle().getFillBackgroundColor()
                                    + "; ");
                            break;
                        default:
                            System.out.print("default");
                    }
                }
            }
            System.out.println();
            i++;
        }

    }

}