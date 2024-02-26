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

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URL;
import java.net.URLConnection;
import java.util.*;

public class HtmlConverter {

    public static void outputXls() throws IOException {

        try {
            // get module list
            ArrayList<Map<String, String>> moduleList = extractHtml();

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
            String outputFileName = "output2.xls";

            // write to output file
            try (FileOutputStream fileOut = new FileOutputStream("src/main/outputs/html" + File.separator + outputFileName)) {
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

    public static ArrayList<Map<String, String>> extractHtml () {

        ArrayList<Map<String, String>> moduleList = new ArrayList<>();

        try {
            // Load HTML file into Jsoup Document
            Document doc = Jsoup.parse(new File("src/main/resources/html/FN_Elektrotechnik_Automation.html"), "UTF-8");

            // Select all <div> elements
            Elements divs = doc.select("div");

            String key = "";
            String value = "";

            // Process each <div>
            for (int divIndex = 0; divIndex < divs.size(); divIndex++) {

                Element div = divs.get(divIndex);
                String className = div.className();
                String text = div.text();

                if (!className.startsWith("c")) {
                    continue;
                }

                // found German module title
                if (className.matches("c\\sx3\\sy([0-9a-f]+)\\sw4\\shf") && text.matches(".+\\s\\(.+\\)")) {

                    // Create a map to store key-value pairs
                    Map<String, String> module = new HashMap<>();
                    Map<String, String> tempModuleKey = new HashMap<>();
                    Map<String, String> tempModuleValue = new HashMap<>();

//                    // Studienakademie
//                    key = "STUDIENAKADEMIE";
//                    value = div.text();
//                    module.put(key, value);
//                    divIndex += 2;
//
//                    // Studienrichtung
//                    div = divs.get(divIndex);
//                    key = "STUDIENRICHTUNG_DE";
//                    value = div.text().split(" // ")[0];
//                    module.put(key, value);
//                    key = "STUDIENRICHTUNG_EN";
//                    value = div.text().split(" // ")[1];
//                    module.put(key, value);
//                    divIndex += 2;
//
//                    // Studiengang
//                    div = divs.get(divIndex);
//                    key = "STUDIENGANG_DE";
//                    value = div.text().split(" // ")[0];
//                    module.put(key, value);
//                    key = "STUDIENGANG_EN";
//                    value = div.text().split(" // ")[1];
//                    module.put(key, value);
//                    divIndex += 2;
//
//                    // Studienbereich
//                    div = divs.get(divIndex);
//                    key = "STUDIENBEREICH_DE";
//                    value = div.text().split(" // ")[0];
//                    module.put(key, value);
//                    key = "STUDIENBEREICH_EN";
//                    value = div.text().split(" // ")[1];
//                    module.put(key, value);
//                    divIndex += 2;

                    // german module title
                    div = divs.get(divIndex);
                    key = "MODULBEZEICHNUNG (DEUTSCH)";
                    value = div.text().split(" \\(")[0];
                    module.put(key, value);
                    divIndex += 2;

                    // english module title
                    div = divs.get(divIndex);
                    key = "MODULBEZEICHNUNG (ENLISCH)";
                    value = div.text();
                    module.put(key, value);
                    divIndex += 2;

                    // set all keys with blank values
                    String[] allKeys = new AllKeys().allKeys;
                    List<String> keyList = Arrays.asList(allKeys);
                    for (String eachKey : keyList) {
                        module.put(eachKey, "");
                    }

                    String[] sectionHeader = new AllKeys().sectionHeader;
                    List<String> sectionList = Arrays.asList(sectionHeader);

                    boolean meetGermanTitle = false;
                    boolean meetSectionStart = false;
                    boolean meetSectionEnd = false;

                    // scan module details
                    while (divIndex < divs.size() && !meetGermanTitle) {

                        div = divs.get(divIndex);
                        className = div.className();
                        text = div.text();

                        if (divIndex == divs.size()-1) {
                            for (Map.Entry<String, String> keyEntry : tempModuleKey.entrySet()) {
                                String[] keyCoords = keyEntry.getKey().split(" ");
                                int keyX = Integer.parseInt(keyCoords[0]);

                                StringBuilder concatenatedValues = new StringBuilder();

                                for (Map.Entry<String, String> valueEntry : tempModuleValue.entrySet()) {
                                    String[] valueCoords = valueEntry.getKey().split(" ");
                                    int valueX = Integer.parseInt(valueCoords[0]);

                                    if (valueX == keyX || keyX == 0) {
                                        if (concatenatedValues.length() > 0) {
                                            concatenatedValues.append("; ");
                                        }
                                        concatenatedValues.append(valueEntry.getValue());
                                    }
                                }

                                module.put(keyEntry.getValue(), concatenatedValues.toString());

                            }
                        }

                        if (text.matches("Stand\\svom\\s\\d{2}.\\d{2}.\\d{4}") || text.matches("\\S+\\s\\/\\/\\sSeite\\s\\d+")) {
                            divIndex++;
                            continue;
                        }

                        if (!className.startsWith("c")) {
                            divIndex++;
                            continue;
                        }

                        if (text.equals("FACHKOMPETENZ")) {
                            System.out.println("here");
                        }
                        // encounter next German title
                        if (className.matches("c\\sx3\\sy([0-9a-f]+)\\sw4\\shf") && text.matches(".+\\s\\(.+\\)")) {
                            divIndex--;
                            meetGermanTitle = true;
                            meetSectionEnd = true;
                        } else {
                            if (className.matches("c\\sx([0-9a-f]+)\\sy([0-9a-f]+)\\sw([0-9a-f]+)\\sh([0-9a-f]+)")) {

                                // encounter section
                                if (sectionList.contains(text)) {
                                    if (!meetSectionStart) {
                                        meetSectionStart = true;
                                    } else {
                                        meetSectionEnd = true;
                                    }
                                } else {
                                    // encounter regular cells
                                    // Extract x, y, w and h values from the class name
                                    String[] values = div.className().replaceAll("c\\sx([0-9a-f]+)\\sy([0-9a-f]+)\\sw([0-9a-f]+)\\sh([0-9a-f]+)", "$1 $2 $3 $4").split(" ");
                                    int x = Integer.parseInt(values[0], 16);
                                    int y = Integer.parseInt(values[1], 16);
                                    int w = Integer.parseInt(values[2], 16);
                                    int h = Integer.parseInt(values[3], 16);
                                    String tempKey = x + " " + y;

                                    if (keyList.contains(text)) {
                                        tempModuleKey.put(tempKey, text);
                                    } else {
                                        tempModuleValue.put(tempKey, text);
                                    }
                                }
                            }
                        }
                        divIndex++;

                        if (className.equals("c x10 y113 wa h1")) {
                            System.out.println("here");
                        }
                        if (meetSectionEnd) {
                            for (Map.Entry<String, String> keyEntry : tempModuleKey.entrySet()) {
                                String[] keyCoords = keyEntry.getKey().split(" ");
                                int keyX = Integer.parseInt(keyCoords[0]);
                                String keyKey = keyEntry.getValue();

                                StringBuilder concatenatedValues = new StringBuilder();

                                if (module.keySet().contains(keyKey) && !module.get(keyKey).toString().isEmpty()) {
                                    concatenatedValues.append(module.get(keyKey));
                                }

                                for (Map.Entry<String, String> valueEntry : tempModuleValue.entrySet()) {
                                    String[] valueCoords = valueEntry.getKey().split(" ");
                                    int valueX = Integer.parseInt(valueCoords[0]);

                                    if (valueX == keyX || keyX == 0) {
                                        if (concatenatedValues.length() > 0) {
                                            concatenatedValues.append("; ");
                                        }
                                        concatenatedValues.append(valueEntry.getValue());
                                    }
                                }

                                module.put(keyEntry.getValue(), concatenatedValues.toString());

                            }

                            tempModuleKey.clear();


                            if (meetSectionEnd && !meetGermanTitle) {
                                if (text.equals("BESONDERHEITEN") || text.equals("VORAUSSETZUNGEN") || text.equals("LITERATUR")) {
                                    tempModuleKey.put("0 0", text);
                                };
                            }

                            tempModuleValue.clear();

                            meetSectionEnd = false;
                        }

                    }
                    moduleList.add(module);
                    divIndex--;


                }

            }
            System.out.println(moduleList);


        } catch (IOException e) {
            e.printStackTrace();
        }

        return moduleList;

    }

}
