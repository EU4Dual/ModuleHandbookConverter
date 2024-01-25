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

import static org.apache.poi.ss.usermodel.Row.MissingCellPolicy.CREATE_NULL_AS_BLANK;

public class Main {
    public static void main(String[] args) throws IOException {

//        readAllRows("src/main/resources/Modulhandbuecher/HDH-Informatik-Allgemeine Informatik.xls");

        readExcel("src/main/resources/Modulhandbuecher/HDH-Informatik-Allgemeine Informatik.xls");
//        readExcel("src/main/resources/Modulhandbuecher/HDH-Informatik-Informationstechnik.xls");
//        readExcel("src/main/resources/Modulhandbuecher/KA-Informatik-Informatik.xls");
//        readExcel("src/main/resources/Modulhandbuecher/MOS-Informatik-Angewandte-Informatik.xls");
//        readExcel("src/main/resources/Modulhandbuecher/STG-Informatik-Informatik.xls");

    }

    public static void getKeyList() {
        ArrayList keyList = new ArrayList();
        keyList.add("MODULNUMMER");
        keyList.add("VERORTUNG IM STUDIENVERLAUF");
        keyList.add("MODULDAUER (SEMESTER)");
        keyList.add("MODULVERANTWORTUNG");
        keyList.add("SPRACHE");
        keyList.add("LEHRFORMEN");
        keyList.add("LEHRMETHODEN");
        keyList.add("PRÜFUNGSLEISTUNG");
        keyList.add("PRÜFUNGSUMFANG (IN MINUTEN)");
        keyList.add("BENOTUNG");
        keyList.add("WORKLOAD INSGESAMT (IN H)");
        keyList.add("DAVON PRÄSENZZEIT (IN H)");
        keyList.add("DAVON SELBSTSTUDIUM (IN H)");
        keyList.add("ECTS-LEISTUNGSPUNKTE");
        keyList.add("FACHKOMPETENZ");
        keyList.add("METHODENKOMPETENZ");
        keyList.add("PERSONALE UND SOZIALE KOMPETENZ");
        keyList.add("ÜBERGREIFENDE HANDLUNGSKOMPETENZ");
    }

    public static void readExcel(String fileLocation) throws IOException, NullPointerException {

        FileInputStream file = new FileInputStream(fileLocation);
        Workbook workbook = new HSSFWorkbook(file);
        workbook.setMissingCellPolicy(CREATE_NULL_AS_BLANK);

        Sheet sheet = workbook.getSheetAt(0);

        ArrayList<Integer> moduleRowNumList = getModuleRowNum(sheet);

        ArrayList<Map<String, String>> moduleList = new ArrayList<>(moduleRowNumList.size());

        for (int i = 0; i < moduleRowNumList.size(); i++) {
            readModule(sheet, moduleRowNumList, moduleList);
        }

    }

    public static ArrayList readModuleList(Sheet sheet, ArrayList list) throws IOException {

        Map<String, String> module = new HashMap<>();

        // Starting from Row 22: List of all modules
        for (int i = 22; i < sheet.getLastRowNum(); i++) {

            Row row = sheet.getRow(i);

            if (row.getCell(1) == null) { break; }

            module.put("MODULNUMMER", row.getCell(1).getRichStringCellValue().getString());
            module.put("MODULBEZEICHNUNG", row.getCell(2).getRichStringCellValue().getString());
//            module.put("VERORTUNG", row.getCell(4).getRichStringCellValue().getString());
//            module.put("ECTS", Double.toString(row.getCell(5).getNumericCellValue()));

            list.add(module);
            System.out.println(module);
        }

        return list;
    }

    public static ArrayList<Map<String, String>> readModule (Sheet sheet, ArrayList<Integer> moduleRowNumList, ArrayList<Map<String, String>> moduleList) throws IOException {

        Map<String, String> module = new HashMap<>();

        boolean eof = false;
        int moduleIndex = 0;
        while (!eof && moduleIndex < moduleRowNumList.size()) {
            int rowNum = moduleRowNumList.get(moduleIndex);
            int nextModuleIndex = (rowNum < moduleRowNumList.getLast() ? moduleRowNumList.get(moduleIndex + 1) : sheet.getLastRowNum());
            while (!eof && rowNum < nextModuleIndex) {

                Row row = sheet.getRow(rowNum);

                if (row == null) {
                    eof = true;
                    break;
                }

                for (Cell cell : row) {

                    // bold burgundy - FontIndex: 9
                    // bold black    - FontIndex: 10
                    // regular black - FontIndex: 12

                    if (cell.getCellStyle().getFontIndex() == 10) {

                        String key = cell.getStringCellValue();

                        int columnIndex = cell.getColumnIndex();
                        String value = "";
                        Cell currentCell = sheet.getRow(rowNum + 1).getCell(columnIndex);
                        if (currentCell.getCellStyle().getFontIndex() == 12) {
                            if (currentCell.getCellType().equals(CellType.STRING)) {
                                value = currentCell.getStringCellValue();
                            } else if (currentCell.getCellType().equals(CellType.NUMERIC)) {
                                value = Double.toString(currentCell.getNumericCellValue());
                            }
                            if (rowNum == 1316) {
                                System.out.println("xxx");
                            }
                            int i = 2;
//                        Cell nextCell = sheet.getRow(rowNum + i++).getCell(columnIndex);
                            if (rowNum < sheet.getLastRowNum()) {
//                                System.out.println("rowNum: " + rowNum + ", lastRowNum: " + sheet.getLastRowNum() + ", i: " + i);
                                while (!eof && sheet.getRow(rowNum + i++).getCell(columnIndex).getCellStyle().getFontIndex() == 12) {
                                    if (sheet.getRow(rowNum + i++).getCell(columnIndex).getCellType().equals(CellType.STRING)) {
                                        value += "; " + currentCell.getStringCellValue();
                                    } else if (sheet.getRow(rowNum + i++).getCell(columnIndex).getCellType().equals(CellType.NUMERIC)) {
                                        value += "; " + currentCell.getNumericCellValue();
                                    }
                                }
                                module.put(key, value);
                            }
                        }
                    }
                }

                rowNum++;

//                module.put("MODULNUMMER", row.getCell(0).getRichStringCellValue().getString());
//                module.put("VERORTUNG IM STUDIENVERLAUF", row.getCell(1).getRichStringCellValue().getString());
//                module.put("MODULDAUER (SEMESTER)", Double.toString(row.getCell(2).getNumericCellValue()));
//                module.put("MODULVERANTWORTUNG", row.getCell(3).getRichStringCellValue().getString());
//                module.put("SPRACHE", row.getCell(5).getRichStringCellValue().getString());

            }

            moduleList.add(module);
            System.out.println(module);

            moduleIndex++;
        }

        return moduleList;
    }

    // Loop for each module
    public static ArrayList<Integer> getModuleRowNum(Sheet sheet) throws IOException {

        ArrayList<Integer> moduleRowNumList = new ArrayList<>();

        for (Row row : sheet) {
            for (Cell cell : row) {

                if (cell.getCellType().equals(CellType.STRING) && cell.getRichStringCellValue().getString().equals("MODULNUMMER")) {
                    moduleRowNumList.add(row.getRowNum());
                }
            }
        }

        return moduleRowNumList;
    }

    public static void readAllRows(String fileLocation) throws IOException {

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