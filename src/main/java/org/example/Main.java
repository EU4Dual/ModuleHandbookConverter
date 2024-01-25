package org.example;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static org.apache.poi.ss.usermodel.Row.MissingCellPolicy.CREATE_NULL_AS_BLANK;

public class Main {
    public static void main(String[] args) throws IOException {

//        readAllRows("src/main/resources/Modulhandbuecher/HDH-Informatik-Allgemeine Informatik.xls");

        ArrayList<Map<String, String>> moduleList = readExcel("src/main/resources/Modulhandbuecher/HDH-Informatik-Allgemeine Informatik.xls");
//        readExcel("src/main/resources/Modulhandbuecher/HDH-Informatik-Informationstechnik.xls");
//        readExcel("src/main/resources/Modulhandbuecher/KA-Informatik-Informatik.xls");
//        readExcel("src/main/resources/Modulhandbuecher/MOS-Informatik-Angewandte-Informatik.xls");
//        readExcel("src/main/resources/Modulhandbuecher/STG-Informatik-Informatik.xls");

        writeExcel(moduleList);

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

    public static ArrayList<Map<String, String>> readExcel(String fileLocation) throws IOException, NullPointerException {

        FileInputStream file = new FileInputStream(fileLocation);
        Workbook workbook = new HSSFWorkbook(file);
        workbook.setMissingCellPolicy(CREATE_NULL_AS_BLANK);

        Sheet sheet = workbook.getSheetAt(0);

        //ArrayList<Integer> moduleRowNumList = getModuleRowNum(sheet);

        ArrayList<Map<String, String>> moduleList = new ArrayList<>();

//        for (int i = 0; i < moduleRowNumList.size(); i++) {
//            readModule(sheet, moduleRowNumList, moduleList);
//        }

        String key = "";
        String value = "";

        int rowIndex = 0;

        boolean titleFound = false;

        while (rowIndex < sheet.getLastRowNum()) {

            Row row = sheet.getRow(rowIndex);

            if (row == null) {
                rowIndex++;
                continue;
            }

            for (int cellIndex = 0; cellIndex < row.getLastCellNum(); cellIndex++) {

                Cell cell = row.getCell(cellIndex);

                // find module title
                if (cell.getCellStyle().getFontIndex() == 13 && !cell.getStringCellValue().equals("Curriculum (Pflicht und Wahlmodule)")) {

                    titleFound = true;

                    Map<String, String> module = new HashMap<>();

                    // german module title
                    key = "MODULBEZEICHNUNG (Deutsch)";
                    value = cell.getStringCellValue().split(" \\(")[0];
                    module.put(key, value);
                    rowIndex++;
                    // english module title
                    cell = sheet.getRow(rowIndex).getCell(cellIndex);
                    key = "MODULBEZEICHNUNG (Englisch)";
                    value = cell.getStringCellValue();
                    module.put(key, value);
                    rowIndex++;

                    // scan module details
                    while (rowIndex < sheet.getLastRowNum() && cell.getCellStyle().getFontIndex() != 13) {

                        row = sheet.getRow(rowIndex);

                        if (row == null) {
                            rowIndex++;
                            continue;
                        }

                        int scanLineNum = 1;

                        for (cellIndex = 0; cellIndex < row.getLastCellNum(); cellIndex++) {

                            cell = row.getCell(cellIndex);

                            // bold black
                            if (cell.getCellStyle().getFontIndex() == 10 && sheet.getRow(rowIndex+1).getCell(cellIndex).getCellStyle().getFontIndex() == 12) {
                                key = cell.getStringCellValue();
                                Cell valueCell = sheet.getRow(rowIndex + 1).getCell(cellIndex);
                                if (valueCell.getCellType().equals(CellType.STRING)) {
                                    value = valueCell.getStringCellValue();
                                } else if (valueCell.getCellType().equals(CellType.NUMERIC)) {
                                    value = Double.toString(valueCell.getNumericCellValue());
                                }

                                int i = 2;
                                if (i > scanLineNum) {
                                    while (sheet.getRow(rowIndex + i).getCell(cellIndex).getCellStyle().getFontIndex() == 12) {
                                        valueCell = sheet.getRow(rowIndex + i++).getCell(cellIndex);
                                        if (valueCell.getCellType().equals(CellType.STRING)) {
                                            value += "; " + valueCell.getStringCellValue();
                                        } else if (valueCell.getCellType().equals(CellType.NUMERIC)) {
                                            value += "; " + Double.toString(valueCell.getNumericCellValue());
                                        }
                                    }
                                } else {
                                    for (int counter = 2; counter < scanLineNum; counter++) {
                                        if (sheet.getRow(rowIndex + counter).getCell(cellIndex).getCellStyle().getFontIndex() == 12) {
                                            valueCell = sheet.getRow(rowIndex + counter).getCell(cellIndex);
                                            if (valueCell.getCellType().equals(CellType.STRING)) {
                                                value += "; " + valueCell.getStringCellValue();
                                            } else if (valueCell.getCellType().equals(CellType.NUMERIC)) {
                                                value += "; " + Double.toString(valueCell.getNumericCellValue());
                                            }
                                        }
                                    }
                                }

                                module.put(key, value);

                                scanLineNum = (Math.max(i, scanLineNum));
                            } else

                            // bold burgundy
                            if (cell.getCellStyle().getFontIndex() == 9 && sheet.getRow(rowIndex+1).getCell(cellIndex).getCellStyle().getFontIndex() == 12) {
                                key = cell.getStringCellValue();
                                Cell valueCell = sheet.getRow(rowIndex + 1).getCell(cellIndex);
                                if (valueCell.getCellType().equals(CellType.STRING)) {
                                    value = valueCell.getStringCellValue();
                                } else if (valueCell.getCellType().equals(CellType.NUMERIC)) {
                                    value = Double.toString(valueCell.getNumericCellValue());
                                }

                                int i = 2;
                                while (sheet.getRow(rowIndex + i).getCell(cellIndex).getCellStyle().getFontIndex() == 12) {
                                    valueCell = sheet.getRow(rowIndex + i++).getCell(cellIndex);
                                    if (valueCell.getCellType().equals(CellType.STRING)) {
                                        value += "; " + valueCell.getStringCellValue();
                                    } else if (valueCell.getCellType().equals(CellType.NUMERIC)) {
                                        value += "; " + Double.toString(valueCell.getNumericCellValue());
                                    }
                                }

                                module.put(key, value);

                                scanLineNum = (Math.max(i, scanLineNum));
                            }


                        }

                        rowIndex += scanLineNum;
                        cell = sheet.getRow(rowIndex).getCell(0);

                    }

                    //System.out.println(module);

                    moduleList.add(module);
                } else {

                }
            }
            if (!titleFound) {
                rowIndex++;
            }
        }
        System.out.println(moduleList);

        return moduleList;
    }

    public static void writeExcel(ArrayList<Map<String, String>> moduleList) throws IOException {

        Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet("sheet1");

        int totalRow = moduleList.size();
        int totalColumn = moduleList.getFirst().size();

        // Write the header row
        Row headerRow = sheet.createRow(0);
        int cellIndex = 0;
        for (String key : moduleList.get(0).keySet()) {
            Cell cell = headerRow.createCell(cellIndex++);
            cell.setCellValue(key);
        }

        // Write data rows
        int rowIndex = 1;
        for (Map<String, String> module : moduleList) {
            Row dataRow = sheet.createRow(rowIndex++);
            cellIndex = 0;
            for (String key : module.keySet()) {
                Cell cell = dataRow.createCell(cellIndex++);
                cell.setCellValue(module.get(key));
            }
        }

        FileOutputStream file = new FileOutputStream("src/main/outputs/output.xls");
        workbook.write(file);
        file.close();
        System.out.println("Data Copied to Excel");

    }
//    public static ArrayList readModuleList(Sheet sheet, ArrayList list) throws IOException {
//
//        Map<String, String> module = new HashMap<>();
//
//        // Starting from Row 22: List of all modules
//        for (int i = 22; i < sheet.getLastRowNum(); i++) {
//
//            Row row = sheet.getRow(i);
//
//            if (row.getCell(1) == null) { break; }
//
//            module.put("MODULNUMMER", row.getCell(1).getRichStringCellValue().getString());
//            module.put("MODULBEZEICHNUNG", row.getCell(2).getRichStringCellValue().getString());
////            module.put("VERORTUNG", row.getCell(4).getRichStringCellValue().getString());
////            module.put("ECTS", Double.toString(row.getCell(5).getNumericCellValue()));
//
//            list.add(module);
//            System.out.println(module);
//        }
//
//        return list;
//    }
//
//    public static ArrayList<Map<String, String>> readModule (Sheet sheet, ArrayList<Integer> moduleRowNumList, ArrayList<Map<String, String>> moduleList) throws IOException {
//
//        Map<String, String> module = new HashMap<>();
//
//        boolean eof = false;
//        int moduleIndex = 0;
//        while (!eof && moduleIndex < moduleRowNumList.size()) {
//            int rowNum = moduleRowNumList.get(moduleIndex);
//            int nextModuleIndex = (rowNum < moduleRowNumList.getLast() ? moduleRowNumList.get(moduleIndex + 1) : sheet.getLastRowNum());
//            while (!eof && rowNum < nextModuleIndex) {
//
//                Row row = sheet.getRow(rowNum);
//
//                if (row == null) {
//                    eof = true;
//                    break;
//                }
//
//                for (Cell cell : row) {
//
//                    // bold burgundy - FontIndex: 9
//                    // bold black    - FontIndex: 10
//                    // regular black - FontIndex: 12
//                    // German module title - 13
//                    // English module title - 15
//
//                    if (cell.getCellStyle().getFontIndex() == 10) {
//
//                        String key = cell.getStringCellValue();
//
//                        int columnIndex = cell.getColumnIndex();
//                        String value = "";
//                        Cell currentCell = sheet.getRow(rowNum + 1).getCell(columnIndex);
//                        if (currentCell.getCellStyle().getFontIndex() == 12) {
//                            if (currentCell.getCellType().equals(CellType.STRING)) {
//                                value = currentCell.getStringCellValue();
//                            } else if (currentCell.getCellType().equals(CellType.NUMERIC)) {
//                                value = Double.toString(currentCell.getNumericCellValue());
//                            }
//                            if (rowNum == 1316) {
//                                System.out.println("xxx");
//                            }
//                            int i = 2;
////                        Cell nextCell = sheet.getRow(rowNum + i++).getCell(columnIndex);
//                            if (rowNum < sheet.getLastRowNum()) {
////                                System.out.println("rowNum: " + rowNum + ", lastRowNum: " + sheet.getLastRowNum() + ", i: " + i);
//                                while (!eof && sheet.getRow(rowNum + i++).getCell(columnIndex).getCellStyle().getFontIndex() == 12) {
//                                    if (sheet.getRow(rowNum + i++).getCell(columnIndex).getCellType().equals(CellType.STRING)) {
//                                        value += "; " + currentCell.getStringCellValue();
//                                    } else if (sheet.getRow(rowNum + i++).getCell(columnIndex).getCellType().equals(CellType.NUMERIC)) {
//                                        value += "; " + currentCell.getNumericCellValue();
//                                    }
//                                }
//                                module.put(key, value);
//                            }
//                        }
//                    }
//                }
//
//                rowNum++;
//
////                module.put("MODULNUMMER", row.getCell(0).getRichStringCellValue().getString());
////                module.put("VERORTUNG IM STUDIENVERLAUF", row.getCell(1).getRichStringCellValue().getString());
////                module.put("MODULDAUER (SEMESTER)", Double.toString(row.getCell(2).getNumericCellValue()));
////                module.put("MODULVERANTWORTUNG", row.getCell(3).getRichStringCellValue().getString());
////                module.put("SPRACHE", row.getCell(5).getRichStringCellValue().getString());
//
//            }
//
//            moduleList.add(module);
//            System.out.println(module);
//
//            moduleIndex++;
//        }
//
//        return moduleList;
//    }
//
//    // Loop for each module
//    public static ArrayList<Integer> getModuleRowNum(Sheet sheet) throws IOException {
//
//        ArrayList<Integer> moduleRowNumList = new ArrayList<>();
//
//        for (Row row : sheet) {
//            for (Cell cell : row) {
//
//                if (cell.getCellType().equals(CellType.STRING) && cell.getRichStringCellValue().getString().equals("MODULNUMMER")) {
//                    moduleRowNumList.add(row.getRowNum());
//                }
//            }
//        }
//
//        return moduleRowNumList;
//    }
//
//    public static void readAllRows(String fileLocation) throws IOException {
//
//        FileInputStream file = new FileInputStream(new File(fileLocation));
//        Workbook workbook = new HSSFWorkbook(file);
//
//        Sheet sheet = workbook.getSheetAt(0);
//
//        Map<Integer, List<String>> data = new HashMap<>();
//        int i = 0;
//        for (Row row : sheet) {
//
//            data.put(i, new ArrayList<String>());
//            for (Cell cell : row) {
//
//                boolean isBlankCell = CellType.BLANK == cell.getCellType();
//                boolean isEmptyStringCell = CellType.STRING == cell.getCellType() && cell.getStringCellValue().trim().isEmpty();
//
//                if (isBlankCell || isEmptyStringCell) {
////                if (cell == null || cell.getCellType() == CellType.BLANK) {
//                    System.out.print("blank; ");
//                } else {
//
//                    switch (cell.getCellType()) {
//                        case STRING:
//                            System.out.print(cell.getRichStringCellValue().getString() + ": " + cell.getCellType().toString() + ", "
//                                    + "FontIndex: " + cell.getCellStyle().getFontIndex() + ", "
//                                    + "FillBackgroundColor: " + cell.getCellStyle().getFillBackgroundColor()
//                                    + "; ");
//                            break;
//                        case NUMERIC:
//                            System.out.print(cell.getNumericCellValue() + ": " + cell.getCellType().toString() + "; "
//                                    + "FontIndex: " + cell.getCellStyle().getFontIndex() + ", "
//                                    + "FillBackgroundColor: " + cell.getCellStyle().getFillBackgroundColor()
//                                    + "; ");
//                            break;
//                        default:
//                            System.out.print("default");
//                    }
//                }
//            }
//            System.out.println();
//            i++;
//        }
//
//    }

}