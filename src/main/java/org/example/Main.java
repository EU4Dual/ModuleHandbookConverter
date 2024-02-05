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

        String inputFolderPath = "src/main/resources/excel";
        String outputFolderPath = "src/main/outputs/excel";

        File inputFolder = new File(inputFolderPath);
        File outputFolder = new File(outputFolderPath);
        File[] listOfFiles = inputFolder.listFiles();

        if (!outputFolder.exists()) {
            if (outputFolder.mkdirs()) {
                System.out.println("Output folder created: " + outputFolder.getAbsolutePath());
            } else {
                System.out.println("Failed to create output folder");
                return; // Exit the program if the folder creation fails
            }
        }

        // loop all files in input folder
        if (listOfFiles != null) {
            ArrayList<ValidFile> validFileList = new ArrayList<>();
            int validFilesCount = 0;

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
                System.out.print("Handling " + (count++) + "/" + validFileList.size() + " file...   ");
                if (vf.type.equals(TypeXls)) {
                    ExcelConverter.rewriteXls(vf, outputFolderPath);
                } else if (vf.type.equals(TypeXlsx)) {
                    ExcelConverter.rewriteXlsx(vf, outputFolderPath);
                }
            }
            System.out.println("Process completed");
        } else {
            System.out.println("Invalid input folder path");
        }
    }
}