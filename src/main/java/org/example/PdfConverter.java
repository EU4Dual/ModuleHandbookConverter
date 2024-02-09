package org.example;

import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.io.*;
import java.net.URL;
import java.net.URLConnection;

public class PdfConverter {

    public static void crawlPdfLinks () {

        // Specify the URL of the HTML document
        String[] urlArray = {"https://www.dhbw.de/fileadmin/user/public/SP/Studienbereich_Technik.htm",
                            "https://www.dhbw.de/fileadmin/user/public/SP/Studienbereich_Wirtschaft.htm",
                            "https://www.dhbw.de/fileadmin/user/public/SP/Studienbereich_Gesundheit.htm",
                            "https://www.dhbw.de/fileadmin/user/public/SP/Studienbereich_Sozialwesen.htm"
                            };

        // String url = "https://www.dhbw.de/fileadmin/user/public/SP/Studienbereich_Technik.htm";

        try {
            int urlCount = 1;
            int pdfCount = 1;
            // Fetch the HTML content from the URL
            for (String url : urlArray) {

                System.out.println("Crawling the " + urlCount++ + " URL...");

                Document document = Jsoup.connect(url).get();

                // Select all anchor elements with an href attribute containing ".pdf"
                Elements links = document.select("a[href$=.pdf]");

                // Download each PDF file
                for (Element link : links) {
                    String pdfLink = link.attr("href");
                    pdfLink = "https://www.dhbw.de/fileadmin/user/public/SP/" + pdfLink;
                    System.out.print("[" + pdfCount++ + "] ");
                    downloadFile(pdfLink);
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void downloadFile(String fileUrl) {
        try {
            URL url = new URL(fileUrl);
            URLConnection connection = url.openConnection();

            // Get the file name from the URL
            String fileName = fileUrl.substring(fileUrl.indexOf("/SP/") + 4);
            fileName = fileName.replace('/', '_');

            try (BufferedInputStream in = new BufferedInputStream(connection.getInputStream());
                 FileOutputStream fileOutputStream = new FileOutputStream("src/main/outputs/pdf/" + fileName)) {

                byte[] dataBuffer = new byte[1024];
                int bytesRead;

                while ((bytesRead = in.read(dataBuffer, 0, 1024)) != -1) {
                    fileOutputStream.write(dataBuffer, 0, bytesRead);
                }

                System.out.println("File downloaded: " + fileName);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    }
