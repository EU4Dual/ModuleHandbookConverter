package org.example;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVPrinter;
import org.apache.commons.csv.CSVRecord;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.*;
import java.util.stream.Collectors;

import static org.apache.poi.ss.usermodel.Row.MissingCellPolicy.CREATE_NULL_AS_BLANK;

public class CsvConverter {

    public static void rewriteCsv(File moduldaten, File moduletexte, String outputFolderPath) throws IOException {

        try {
            // get module list
            String moduldatenLocation = moduldaten.getAbsolutePath();
            String modultexteLocation = moduletexte.getAbsolutePath();

            System.out.print("Reading moduldaten...   ");
            ArrayList<Map<String, String>> moduleList = readModuldaten(moduldatenLocation);
            System.out.println("Done");

            System.out.print("Reading modultexte...   ");
            updateModuleList(moduleList, modultexteLocation);
            System.out.println("Done");

            // write to output file
            writeModuleListToCSV(moduleList, outputFolderPath + File.separator + "module-reformatted.csv");

        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    public static ArrayList<Map<String, String>> readModuldaten(String fileLocation) throws IOException, NullPointerException {

        ArrayList<Map<String, String>> moduleList = new ArrayList<>();

        try {
            // read moduledaten
            Reader reader = Files.newBufferedReader(Paths.get(fileLocation));
            CSVParser csvParser = new CSVParser(reader, CSVFormat.newFormat(',').withHeader().withQuote('"'));
            List<String> header = csvParser.getHeaderNames();
            int size = header.size();
            for (CSVRecord csvRecord : csvParser) {
                Map<String, String> module = new HashMap<>();
                for (int i = 0; i < size; i++) {
                    module.put(header.get(i), csvRecord.get(i));
                }
                moduleList.add(module);
            }

        } catch (IOException e) {
            throw new RuntimeException(e);
        }

        return moduleList;
    }

    public static void updateModuleList(ArrayList<Map<String, String>> moduleList, String fileLocation) throws IOException {
        try {
            Reader reader = Files.newBufferedReader(Paths.get(fileLocation));
            CSVParser csvParser = new CSVParser(reader, CSVFormat.newFormat(',').withHeader().withQuote('"'));

            for (CSVRecord csvRecord : csvParser) {
                String id = csvRecord.get("ID");
                String category = csvRecord.get("KATEGORIE");
                String text = csvRecord.get("TEXT");

                // Find modules with matching MODULID or UNITID
                List<Map<String, String>> matchingModules = moduleList.stream()
                        .filter(module -> id.equals(module.get("MODULID")) || id.equals(module.get("UNITID")))
                        .collect(Collectors.toList());

                // Update matching modules with the new category and text
                for (Map<String, String> module : matchingModules) {
                    module.put(category, text);
                }
            }

        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }


    public static void writeModuleListToCSV(ArrayList<Map<String, String>> moduleList, String outputPath) throws IOException {
        // Extract all unique keys from all modules
        Set<String> allKeys = moduleList.stream()
                .flatMap(module -> module.keySet().stream())
                .distinct()
                .collect(Collectors.toSet());

        // Convert the set of keys to an array for CSV header
        String[] headers = allKeys.toArray(new String[0]);

        try (CSVPrinter csvPrinter = new CSVPrinter(new FileWriter(outputPath), CSVFormat.DEFAULT.withHeader(headers))) {
            for (Map<String, String> module : moduleList) {
                // Create a record array with values for each header
                String[] record = allKeys.stream().map(module::get).toArray(String[]::new);
                csvPrinter.printRecord(record);
            }
        }
        System.out.println("module-reformatted.csv created");
    }
}