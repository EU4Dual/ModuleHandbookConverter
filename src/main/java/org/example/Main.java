package org.example;

import java.io.*;
import java.util.ArrayList;

class ValidFile {
    String type;
    File file;
}

public class Main {

    public static final String TypeXls = ".xls";
    public static final String TypeXlsx = ".xlsx";
    public static final String TypeCsv = ".csv";

    public static void main(String[] args) throws IOException {

//        String inputFolderPath = "src/main/resources/excel";
//        String outputFolderPath = "src/main/outputs/excel";

        String inputFolderPath = "src/main/resources/csv";
        String outputFolderPath = "src/main/outputs/csv";

        File inputFolder = new File(inputFolderPath);
        File outputFolder = new File(outputFolderPath);
        File[] listOfFiles = inputFolder.listFiles();

        if (!outputFolder.exists()) {
            // create the output folder if not already exist
            if (outputFolder.mkdirs()) {
                System.out.println("Output folder created: " + outputFolder.getAbsolutePath());
            } else {
                // exit the program if the folder creation fails
                System.out.println("Failed to create output folder");
                return;
            }
        }

        // loop all files in input folder
        if (listOfFiles != null) {
            ArrayList<ValidFile> validFileList = new ArrayList<>();

            for (File file : listOfFiles) {
                if (file.getName().endsWith(TypeXls)) {
                    ValidFile validFile = new ValidFile();
                    validFile.type = TypeXls;
                    validFile.file = file;
                    validFileList.add(validFile);
                } else if (file.getName().endsWith(TypeXlsx)) {
                    ValidFile validFile = new ValidFile();
                    validFile.type = TypeXlsx;
                    validFile.file = file;
                    validFileList.add(validFile);
                } else if (file.getName().endsWith(TypeCsv)) {
                    ValidFile validFile = new ValidFile();
                    validFile.type = TypeCsv;
                    validFile.file = file;
                    validFileList.add(validFile);
                }
            }
            System.out.println("Found " + validFileList.size() + " valid files in target folder");
            System.out.println("Process started");
            int count = 1;
            for (ValidFile vf : validFileList) {
                if (vf.type.equals(TypeXls)) {
                    System.out.print("Handling " + (count++) + "/" + validFileList.size() + " file...   ");
                    ExcelConverter.rewriteXls(vf.file, outputFolderPath);
                } else if (vf.type.equals(TypeXlsx)) {
                    System.out.print("Handling " + (count++) + "/" + validFileList.size() + " file...   ");
                    ExcelConverter.rewriteXlsx(vf.file, outputFolderPath);
                } else if (vf.type.equals(TypeCsv)) {
                    File moduldaten = null;
                    File modultexte = null;
                    if (validFileList.get(0).file.getName().startsWith("moduldaten")) {
                        moduldaten = validFileList.get(0).file;
                        modultexte = validFileList.get(1).file;
                    } else {
                        moduldaten = validFileList.get(1).file;
                        modultexte = validFileList.get(0).file;
                    }
                    CsvConverter.rewriteCsv(moduldaten, modultexte, outputFolderPath);
                }
            }
            System.out.println("Process completed");
        } else {
            System.out.println("Invalid input folder path");
        }
    }
}