package org.example;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class HtmlConverter {

    public static void outputXls() throws IOException {

        try {

            File inputFolder = new File("src/main/resources/html");
            File[] folderList = inputFolder.listFiles();
            ArrayList<Map<String, String>> moduleList = new ArrayList<>();
            for (File folder : folderList) {
                File[] files = folder.listFiles();
                String folderName = folder.getName();

                System.out.println("START processing folder [" + folderName + "]");
                for (File file : files) {
                    System.out.print("   reading " + file.getName() + " .... ");
                    // get module list
                    moduleList.addAll(extractHtml(file));
                    System.out.println("DONE");
                }

                // create new workbook and sheet
                Workbook workbook = new HSSFWorkbook();
                Sheet sheet = workbook.createSheet("sheet1");

                // write the header row
                Row headerRow = sheet.createRow(0);
                int cellIndex = 0;
                for (String key : moduleList.getFirst().keySet()) {
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
                String outputFileName = folderName + "_Modulhandbuch_Output.xls";

                // write to output file
                try (FileOutputStream fileOut = new FileOutputStream("src/main/outputs/html" + File.separator + outputFileName)) {
                    workbook.write(fileOut);
                    System.out.println(outputFileName + " created");
                } catch (IOException e) {
                    e.printStackTrace();
                } finally {
                    workbook.close();
                }

                moduleList.clear();

                System.out.println("END processing folder [" + folderName + "]");

            }


        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    public static ArrayList<Map<String, String>> extractHtml (File file) {

        ArrayList<Map<String, String>> moduleList = new ArrayList<>();

        try {
            // Load HTML file into Jsoup Document
            Document doc = Jsoup.parse(file, "UTF-8");


            // Select all <div> elements
            Elements divs = doc.select("div");

            String key = "";
            String value = "";

            boolean readInfo = false;

            Map<String, String> info = new HashMap<>();

            // Process each <div>
            for (int divIndex = 0; divIndex < divs.size(); divIndex++) {

                Element div = divs.get(divIndex);
                String className = div.className();
                String text = div.text();



                // read Studienakademie, Studienrichtung, Studiengang and Studienbereich
                if (className.startsWith("pc") && !readInfo) {

                    ArrayList<String> infoList = new ArrayList<>();

                    // read info div by div
                    while (!text.equals("Modulhandbuch")) {

                        if (className.startsWith("c")) {
                            infoList.add(text);
                        }

                        div = divs.get(divIndex++);
                        className = div.className();
                        text = div.text();

                    }

                    // process info

                    if (!infoList.getFirst().equals("Studienakademie")) {
                        info.put("STUDIENBEREICH_EN", infoList.getFirst());
                    }

                    if (!infoList.getLast().equals("Studiengang")) {
                        info.put("STUDIENBEREICH_DE", infoList.getLast());
                    }

                    for (int i = 0; i < infoList.size(); i++) {
                        if (infoList.get(i).equals("Studienakademie") && i > 0) {
                            info.put("STUDIENAKADEMIE", infoList.get(i-1));
                        }
                        if (infoList.get(i).equals("Studienrichtung") && i > 1) {
                            info.put("STUDIENRICHTUNG_DE", infoList.get(i - 1));
                            if (infoList.get(i-3).equals("Studienakademie")) {
                                info.put("STUDIENRICHTUNG_EN", infoList.get(i - 2));
                            }
                        }
                        if (infoList.get(i).equals("Studiengang") && i > 2) {
                            info.put("STUDIENGANG_DE", infoList.get(i-1));
                            if (infoList.get(i-3).equals("Studienrichtung")) {
                                info.put("STUDIENGANG_EN", infoList.get(i - 2));
                            }
                        }
                    }
                    readInfo = true;

                }

                // found German module title
                if (className.matches("c\\sx3\\sy([0-9a-f]+)\\sw4\\sh([0-9a-f]+)") && text.matches(".+\\s\\(.+\\)") && !text.equals("Curriculum (Pflicht und Wahlmodule)")) {

                    // Create a map to store key-value pairs
                    Map<String, String> module = new HashMap<>(info);
                    Map<String, String> tempModuleKey = new HashMap<>();
                    Map<String, String> tempModuleValue = new HashMap<>();

                    // german module title
                    div = divs.get(divIndex);
                    key = "MODULBEZEICHNUNG_DE";
                    value = div.text().split(" \\(")[0];
                    module.put(key, value);
                    divIndex += 2;

                    while (!divs.get(divIndex).className().startsWith("c")) {
                        divIndex++;
                    }

                    // english module title
                    div = divs.get(divIndex);
                    key = "MODULBEZEICHNUNG_EN";
                    value = div.text();
                    module.put(key, value);
                    divIndex += 2;

                    // set all other keys with blank values
                    String[] allKeys = new AllKeys().allKeys;
                    List<String> keyList = Arrays.asList(allKeys);
                    for (String eachKey : keyList) {
                        if (!module.containsKey(eachKey)) {
                            module.put(eachKey, "");
                        }
                    }

                    String[] sectionHeader = new AllKeys().sectionHeader;
                    List<String> sectionList = Arrays.asList(sectionHeader);

                    boolean meetGermanTitle = false;
                    boolean meetSectionStart = false;
                    boolean meetSectionEnd = false;
                    boolean withinPage = true;

                    // scan details for each module
                    while (divIndex < divs.size() && !meetGermanTitle) {

                        div = divs.get(divIndex);
                        className = div.className();
                        text = div.text();

                        if (!className.startsWith("c")) {
                            divIndex++;
                            continue;
                        }

                        // encounter next German title
                        if (className.matches("c\\sx3\\sy([0-9a-f]+)\\sw4\\sh([0-9a-f]+)") && text.matches(".+\\s\\(.+\\)")) {
                            divIndex--;
                            meetGermanTitle = true;
                            withinPage = true;
                            continue;
                        }

                        // encounter page footer
                        if (text.matches("Stand\\svom\\s\\d{2}.\\d{2}.\\d{4}") || text.matches("\\S+\\s\\/\\/\\sSeite\\s\\d+")) {
                                withinPage = false;
                                meetSectionEnd = true;
                                divIndex++;
                        }

                        // encounter section
                        else if (className.matches("c\\sx([0-9a-f]+)\\sy([0-9a-f]+)\\sw([0-9a-f]+)\\sh([0-9a-f]+)") && sectionList.contains(text)) {
                            if (!meetSectionStart) {
                                meetSectionStart = true;
                                meetSectionEnd = false;
                                if (text.equals("BESONDERHEITEN") || text.equals("VORAUSSETZUNGEN") || text.equals("LITERATUR")
                                        || text.equals("FACHKOMPETENZ") || text.equals("METHODENKOMPETENZ") || text.equals("PERSONALE UND SOZIALE KOMPETENZ")
                                        || text.equals("ÜBERGREIFENDE HANDLUNGSKOMPETENZ") || text.equals("EINGESETZTE LEHRFORMEN")) {
                                    tempModuleKey.put("0 0", text);
                                };
                                if (text.equals("EINGESETZTE LEHR/LERNMETHODEN") || text.equals("EINGESETZTE LEHRFORMEN")) {
                                    tempModuleKey.put("0 0", "EINGESETZTE LEHR/LERNMETHODEN");
                                };
                            } else {
                                meetSectionStart = false;
                                meetSectionEnd = true;
                                divIndex--;
                            }
                            withinPage = true;
                        }

                        // encounter regular cells
                        else if (className.matches("c\\sx([0-9a-f]+)\\sy([0-9a-f]+)\\sw([0-9a-f]+)\\sh([0-9a-f]+)")) {

                            // Extract x, y, w and h values from the class name
                            String[] values = div.className().replaceAll("c\\sx([0-9a-f]+)\\sy([0-9a-f]+)\\sw([0-9a-f]+)\\sh([0-9a-f]+)", "$1 $2 $3 $4").split(" ");
                            int x = Integer.parseInt(values[0], 16);
                            int y = Integer.parseInt(values[1], 16);
                            String tempKey = x + " " + y;

                            if (withinPage) {
                                if (keyList.contains(text)) {
                                    tempModuleKey.put(tempKey, text);
                                } else {
                                    tempModuleValue.put(tempKey, text);
                                }
                            }
                        }

                        // write to module after scanning a whole section
                        if (meetSectionEnd) {
                            for (Map.Entry<String, String> keyEntry : tempModuleKey.entrySet()) {
                                String[] keyCoords = keyEntry.getKey().split(" ");
                                int keyX = Integer.parseInt(keyCoords[0]);
                                String keyKey = keyEntry.getValue();

                                StringBuilder concatenatedValues = new StringBuilder();

                                // get current value if the key is already existed
                                if (module.containsKey(keyKey) && !module.get(keyKey).isEmpty()) {
                                    concatenatedValues.append(module.get(keyKey));
                                }

                                for (Map.Entry<String, String> valueEntry : tempModuleValue.entrySet()) {
                                    String[] valueCoords = valueEntry.getKey().split(" ");
                                    int valueX = Integer.parseInt(valueCoords[0]);

                                    if (valueX == keyX || keyX == 0) {
                                        if (!concatenatedValues.isEmpty()) {
                                            concatenatedValues.append(";; ");
                                        }
                                        concatenatedValues.append(valueEntry.getValue());
                                    }
                                }

                                module.put(keyEntry.getValue(), concatenatedValues.toString());

                            }

                            tempModuleKey.clear();
                            tempModuleValue.clear();

                            meetSectionEnd = false;
                        }
                        divIndex++;
                    }
                    moduleList.add(module);
                    divIndex--;
                }
            }
            // System.out.println(moduleList);
        } catch (IOException e) {
            e.printStackTrace();
        }

        return moduleList;

    }

}
