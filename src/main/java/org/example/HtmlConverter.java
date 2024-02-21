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
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

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
            String outputFileName = "output.xls";

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
            Document doc = Jsoup.parse(new File("src/main/resources/html/Automation.html"), "UTF-8");

            // Create a map to store key-value pairs
            Map<String, String> tempMap = new HashMap<>();

            // Select all <div> elements
            Elements divs = doc.select("div");

            String key = "";
            String value = "";

            boolean titleFound = false;

            // Process each <div>
            for (int divIndex = 0; divIndex < divs.size(); divIndex++) {

                Element div = divs.get(divIndex);
                String className = div.className();

                if (!className.startsWith("c")) {
                    continue;
                }

                // found German module title
                if (className.equals("c x3 y48 w4 hf")) {

                    titleFound = true;

                    Map<String, String> module = new HashMap<>();

                    // german module title
                    key = "MODULBEZEICHNUNG (Deutsch)";
                    value = div.text().split(" \\(")[0];
                    module.put(key, value);
                    divIndex += 2;

                    // english module title
                    div = divs.get(divIndex);
                    key = "MODULBEZEICHNUNG (Englisch)";
                    value = div.text();
                    module.put(key, value);
                    divIndex += 2;

                    // scan module details
                    while (divIndex < divs.size()) {

                        div = divs.get(divIndex);
                        className = div.className();

                        if (!className.startsWith("c")) {
                            divIndex++;
                            continue;
                        }

                        // found new German title
                        if (className.equals("c x3 y48 w4 hf")) {
                            divIndex--;
                            break;
                        } else {
                            // Check if the class name matches the expected pattern
                            if (className.matches("c\\sx([0-9a-f]+)\\sy([0-9a-f]+)\\sw([0-9a-f]+)\\sh([0-9a-f]+)")) {
                                // Extract x, y, w and h values from the class name
                                String[] values = div.className().replaceAll("c\\sx([0-9a-f]+)\\sy([0-9a-f]+)\\sw([0-9a-f]+)\\sh([0-9a-f]+)", "$1 $2 $3 $4").split(" ");
                                int x = Integer.parseInt(values[0], 16);
                                int y = Integer.parseInt(values[1], 16);
                                String w = values[2];
                                String h = values[3];

                                // Get the text content
                                String text = div.text();

                                // Form the key for the map
                                String tempKey = x + " " + y + " " + w;

                                // Check if there's a pair above
                                String aboveKey = x + " " + (y - 1) + " " + w;
                                if (tempMap.containsKey(aboveKey)) {
                                    // Form the key for the final map
                                    String finalKey = tempMap.get(aboveKey);

                                    // Add the pair to the final data map
                                    module.put(finalKey, text);
                                } else {
                                    // Add the current <div> as a potential key to the map
                                    tempMap.put(tempKey, text);
                                }
                            }
                            divIndex++;
                        }

                    }
                    moduleList.add(module);
                }

            }
            System.out.println(moduleList);


        } catch (IOException e) {
            e.printStackTrace();
        }

        return moduleList;

    }

}
