package org.example;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

import static org.apache.poi.ss.usermodel.Row.MissingCellPolicy.CREATE_NULL_AS_BLANK;

public class ExcelConverter {

    public static void rewriteXls(ValidFile inputFile, String outputFolderPath) throws IOException {

        try {

            String type = inputFile.type;
            File file = inputFile.file;
            // get module list
            String fileLocation = file.getAbsolutePath();
            ArrayList<Map<String, String>> moduleList = readXlsModule(type, file);

            // create new workbook and sheet
            Workbook workbook = new HSSFWorkbook();
            Sheet sheet = workbook.createSheet("sheet1");

            // write the header row
            Row headerRow = sheet.createRow(0);
            int cellIndex = 0;
            for (String key : moduleList.get(0).keySet()) {
                Cell cell = headerRow.createCell(cellIndex++);
                cell.setCellValue(key);
            }

            // write data rows
            int rowIndex = 1;
            for (Map<String, String> module : moduleList) {
                Row dataRow = sheet.createRow(rowIndex++);
                cellIndex = 0;
                for (String key : module.keySet()) {
                    Cell cell = dataRow.createCell(cellIndex++);
                    cell.setCellValue(module.get(key));
                }
            }

            // create the output file name
            String outputFileName = getOutputFileName(file.getName());

            // write to output file
            try (FileOutputStream fileOut = new FileOutputStream(outputFolderPath + File.separator + outputFileName)) {
                workbook.write(fileOut);
                System.out.println(outputFileName + " created");
            } catch (IOException e) {
                e.printStackTrace();
            } finally {
                workbook.close();
            }

        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    public static ArrayList<Map<String, String>> readXlsModule(String type, File file) throws IOException, NullPointerException {

        FileInputStream fileInputStream = new FileInputStream(file.getAbsolutePath());
        Workbook workbook = new HSSFWorkbook(fileInputStream);
        workbook.setMissingCellPolicy(CREATE_NULL_AS_BLANK);
        Sheet sheet = workbook.getSheetAt(0);

        ArrayList<Map<String, String>> moduleList = new ArrayList<>();

        String key = "";
        String value = "";

        int rowIndex = 0;

        boolean titleFound = false;

        // loop all rows
        while (rowIndex < sheet.getLastRowNum()) {

            Row row = sheet.getRow(rowIndex);

            // skip empty row
            if (row == null) {
                rowIndex++;
                continue;
            }

            // loop all cells in a row
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

                            // found bold black cell (10) + regular cell (12) below
                            if (cell.getCellStyle().getFontIndex() == 10 && sheet.getRow(rowIndex+1).getCell(cellIndex).getCellStyle().getFontIndex() == 12) {
                                key = cell.getStringCellValue();
                                Cell valueCell = sheet.getRow(rowIndex + 1).getCell(cellIndex);
                                if (valueCell.getCellType().equals(CellType.STRING)) {
                                    value = valueCell.getStringCellValue();
                                } else if (valueCell.getCellType().equals(CellType.NUMERIC)) {
                                    value = Double.toString(valueCell.getNumericCellValue());
                                }

                                // check if more than one regular cell is under this key cell
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

                                // found bold burgundy cell (9) and regular cell (12) below
                                if (cell.getCellStyle().getFontIndex() == 9 && sheet.getRow(rowIndex+1).getCell(cellIndex).getCellStyle().getFontIndex() == 12) {
                                    key = cell.getStringCellValue();
                                    Cell valueCell = sheet.getRow(rowIndex + 1).getCell(cellIndex);
                                    if (valueCell.getCellType().equals(CellType.STRING)) {
                                        value = valueCell.getStringCellValue();
                                    } else if (valueCell.getCellType().equals(CellType.NUMERIC)) {
                                        value = Double.toString(valueCell.getNumericCellValue());
                                    }

                                    // check if more than one regular cell is under this key cell
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
                        // skip scanned rows
                        rowIndex += scanLineNum;
                        cell = sheet.getRow(rowIndex).getCell(0);
                    }
                    moduleList.add(module);
                }
            }
            if (!titleFound) {
                rowIndex++;
            }
        }
        // System.out.println(moduleList);
        return moduleList;
    }

    public static void rewriteXlsx(ValidFile inputFile, String outputFolderPath) throws IOException {

        try {

            String type = inputFile.type;
            File file = inputFile.file;
            // get module list
            String fileLocation = file.getAbsolutePath();
            ArrayList<Map<String, String>> moduleList = readXlsxModule(type, file);

            // create new workbook and sheet
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = workbook.createSheet("sheet1");

            // write the header row
            Row headerRow = sheet.createRow(0);
            int cellIndex = 0;
            for (String key : moduleList.get(0).keySet()) {
                Cell cell = headerRow.createCell(cellIndex++);
                cell.setCellValue(key);
            }

            // write data rows
            int rowIndex = 1;
            for (Map<String, String> module : moduleList) {
                Row dataRow = sheet.createRow(rowIndex++);
                cellIndex = 0;
                for (String key : module.keySet()) {
                    Cell cell = dataRow.createCell(cellIndex++);
                    cell.setCellValue(module.get(key));
                }
            }

            // create the output file name
            String outputFileName = getOutputFileName(file.getName());

            // write to output file
            try (FileOutputStream fileOut = new FileOutputStream(outputFolderPath + File.separator + outputFileName)) {
                workbook.write(fileOut);
                System.out.println(outputFileName + " created");
            } catch (IOException e) {
                e.printStackTrace();
            } finally {
                workbook.close();
            }

        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    public static ArrayList<Map<String, String>> readXlsxModule(String type, File file) throws IOException, NullPointerException {

        FileInputStream fileInputStream = new FileInputStream(file.getAbsolutePath());
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        workbook.setMissingCellPolicy(CREATE_NULL_AS_BLANK);
        XSSFSheet sheet = workbook.getSheetAt(0);

        ArrayList<Map<String, String>> moduleList = new ArrayList<>();

        String key = "";
        String value = "";

        int rowIndex = 0;

        boolean titleFound = false;

        // loop all rows
        while (rowIndex < sheet.getLastRowNum()) {

            Row row = sheet.getRow(rowIndex);

            // skip empty row
            if (row == null) {
                rowIndex++;
                continue;
            }

            // loop all cells in a row
            for (int cellIndex = 0; cellIndex < row.getLastCellNum(); cellIndex++) {

                Cell cell = row.getCell(cellIndex);
                int cellStyle = cell.getCellStyle().getFontIndex();
                // find module title
                if (cell.getCellStyle().getFontIndex() == 9 && !cell.getStringCellValue().equals("Curriculum (Pflicht und Wahlmodule)")) {

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
                    while (rowIndex < sheet.getLastRowNum() && cell.getCellStyle().getFontIndex() != 9) {

                        row = sheet.getRow(rowIndex);

                        if (row == null) {
                            rowIndex++;
                            continue;
                        }

                        int scanLineNum = 1;

                        for (cellIndex = 0; cellIndex < row.getLastCellNum(); cellIndex++) {

                            cell = row.getCell(cellIndex);

                            // found bold red cell (7) + regular cell (7) below
                            if (cell.getCellStyle().getFontIndex() == 7 && sheet.getRow(rowIndex+1).getCell(cellIndex).getCellStyle().getFontIndex() == 10) {
                                key = cell.getStringCellValue();
                                Cell valueCell = sheet.getRow(rowIndex + 1).getCell(cellIndex);
                                if (valueCell.getCellType().equals(CellType.STRING)) {
                                    value = valueCell.getStringCellValue();
                                } else if (valueCell.getCellType().equals(CellType.NUMERIC)) {
                                    value = Double.toString(valueCell.getNumericCellValue());
                                }

                                // check if more than one regular cell is under this key cell
                                int i = 2;
                                if (i > scanLineNum) {
                                    while (sheet.getRow(rowIndex + i).getCell(cellIndex).getCellStyle().getFontIndex() == 10) {
                                        valueCell = sheet.getRow(rowIndex + i++).getCell(cellIndex);
                                        if (valueCell.getCellType().equals(CellType.STRING)) {
                                            value += "; " + valueCell.getStringCellValue();
                                        } else if (valueCell.getCellType().equals(CellType.NUMERIC)) {
                                            value += "; " + Double.toString(valueCell.getNumericCellValue());
                                        }
                                    }
                                } else {
                                    for (int counter = 2; counter < scanLineNum; counter++) {
                                        if (sheet.getRow(rowIndex + counter).getCell(cellIndex).getCellStyle().getFontIndex() == 10) {
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

                                // found bold green cell (8) and regular cell (12) below
                                if (cell.getCellStyle().getFontIndex() == 8 && sheet.getRow(rowIndex+1).getCell(cellIndex).getCellStyle().getFontIndex() == 12) {
                                    key = cell.getStringCellValue();
                                    Cell valueCell = sheet.getRow(rowIndex + 1).getCell(cellIndex);
                                    if (valueCell.getCellType().equals(CellType.STRING)) {
                                        value = valueCell.getStringCellValue();
                                    } else if (valueCell.getCellType().equals(CellType.NUMERIC)) {
                                        value = Double.toString(valueCell.getNumericCellValue());
                                    }

                                    // check if more than one regular cell is under this key cell
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
                        // skip scanned rows
                        rowIndex += scanLineNum;
                        cell = sheet.getRow(rowIndex).getCell(0);
                    }
                    moduleList.add(module);
                }
            }
            if (!titleFound) {
                rowIndex++;
            }
        }
        // System.out.println(moduleList);
        return moduleList;
    }

    private static String getOutputFileName(String originalFileName) {
        // output filename = {original filename}-rewritten.xls
        int lastDotIndex = originalFileName.lastIndexOf('.');
        if (lastDotIndex != -1) {
            String baseName = originalFileName.substring(0, lastDotIndex);
            String extension = originalFileName.substring(lastDotIndex);
            return baseName + "-rewritten" + extension;
        } else {
            return originalFileName + "-rewritten";
        }
    }
}
