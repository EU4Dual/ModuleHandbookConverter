package org.example;

import java.io.*;
import java.util.ArrayList;
import java.util.Scanner;

class ValidFile {
    String type;
    File file;
}

public class Main {

    public static final String TypeXls = ".xls";
    public static final String TypeXlsx = ".xlsx";
    public static final String TypeCsv = ".csv";

    public static final String TypePdf = ".pdf";

    public static void main(String[] args) throws IOException {

        String inputFolderPath, outputFolderPath;
        String fileType = TypeXls;

        System.out.print("Please enter file type [1]Excel or [2]CSV or [3]PDF: ");
        Scanner scanner = new Scanner(System.in);
        switch (scanner.nextLine()) {
            case "1":
                System.out.println("[1]Excel is entered");
                inputFolderPath = "src/main/resources/excel";
                outputFolderPath = "src/main/outputs/excel";
                break;
            case "2":
                System.out.println("[2]CSV is entered");
                fileType = TypeCsv;
                inputFolderPath = "src/main/resources/csv";
                outputFolderPath = "src/main/outputs/csv";
                break;
            case "3":
                System.out.println("[3]PDF is entered");
                fileType = TypePdf;
                inputFolderPath = "src/main/resources/pdf";
                outputFolderPath = "src/main/outputs/pdf";
                break;
            default:
                System.out.println("Invalid input. Process ended.");
                return;
        }

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
            if (!fileType.equals(TypePdf)) {
                System.out.println("Found " + validFileList.size() + " valid files in target folder");
            }
            System.out.println("Process started");
            int count = 1;
            // input files are Excel
            if (fileType.equals(TypeXls)) {
                for (ValidFile vf : validFileList) {
                    if (vf.type.equals(TypeXls)) {
                        System.out.print("Handling " + (count++) + "/" + validFileList.size() + " file...   ");
                        ExcelConverter.rewriteXls(vf.file, outputFolderPath);
                    } else if (vf.type.equals(TypeXlsx)) {
                        System.out.print("Handling " + (count++) + "/" + validFileList.size() + " file...   ");
                        ExcelConverter.rewriteXlsx(vf.file, outputFolderPath);
                    }
                }
                // input files are CSV
            } else if (fileType.equals(TypeCsv)) {
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
                // input files are PDF
            } else if (fileType.equals(TypePdf)) {
                PdfConverter.crawlPdfLinks();
                for (ValidFile vf : validFileList) {
                    if (vf.type.equals(TypePdf)) {
                        System.out.print("Handling " + (count++) + "/" + validFileList.size() + " file...   ");
                        // PdfConverter.rewritePdf(vf.file, outputFolderPath);

                    }
                }
            }
            System.out.println("Process completed");
        } else {
            System.out.println("Invalid input folder path");
        }
    }
}