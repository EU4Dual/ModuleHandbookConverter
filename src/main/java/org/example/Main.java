package org.example;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

import static org.apache.poi.ss.usermodel.Row.MissingCellPolicy.CREATE_NULL_AS_BLANK;

public class Main {
    public static void main(String[] args) throws IOException {

        String inputFolderPath = "src/main/resources/Modulhandbuecher";
        String outputFolderPath = "src/main/outputs";

        File inputFolder = new File(inputFolderPath);
        File[] listOfFiles = inputFolder.listFiles();

        if (listOfFiles != null) {
            for (File file : listOfFiles) {
                if (file.isFile() && file.getName().endsWith(".xls")) {
                    rewriteExcel(file, outputFolderPath);
                }
            }
        } else {
            System.out.println("Invalid input folder path.");
        }
    }


    public static ArrayList<Map<String, String>> readModule(String fileLocation) throws IOException, NullPointerException {

        FileInputStream file = new FileInputStream(fileLocation);
        Workbook workbook = new HSSFWorkbook(file);
        workbook.setMissingCellPolicy(CREATE_NULL_AS_BLANK);

        Sheet sheet = workbook.getSheetAt(0);

        ArrayList<Map<String, String>> moduleList = new ArrayList<>();

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

    public static void rewriteExcel(File inputFile, String outputFolderPath) throws IOException {

        try {

            String fileLocation = inputFile.getAbsolutePath();
            ArrayList<Map<String, String>> moduleList = readModule(fileLocation);

            Workbook workbook = new HSSFWorkbook();
            Sheet sheet = workbook.createSheet("sheet1");

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

            // Create the output file name
            String outputFileName = getOutputFileName(inputFile.getName());

            try (FileOutputStream fileOut = new FileOutputStream(outputFolderPath + File.separator + outputFileName)) {
                workbook.write(fileOut);
                System.out.println("Processed file: " + inputFile.getName());
            } catch (IOException e) {
                e.printStackTrace();
            } finally {
                workbook.close();
            }

        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    private static String getOutputFileName(String originalFileName) {
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