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
            File[] validFiles = new File[listOfFiles.length];
            int validFilesCount = 0;
            for (File file : listOfFiles) {
                if (file.getName().endsWith(".xls")) {
                    validFiles[validFilesCount++] = file;
                }
            }
            System.out.println("Found " + validFilesCount + " valid files in target folder");
            System.out.println("Process started");
            for (int i = 0; i < validFilesCount; i++) {
                if (validFiles[i].isFile() && validFiles[i].getName().endsWith(".xls")) {
                    System.out.print("Handling " + (i+1) + "/" + validFilesCount + " file...   ");
                    ExcelConverter.rewriteExcel(validFiles[i], outputFolderPath);
                }
            }
            System.out.println("Process completed");
        } else {
            System.out.println("Invalid input folder path");
        }
    }
}