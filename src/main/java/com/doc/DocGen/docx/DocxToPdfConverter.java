package com.doc.DocGen.docx;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

public class DocxToPdfConverter {

    private static final int THREAD_POOL_SIZE = 4; // Adjust as necessary

    public static void main(String[] args) {
        ExecutorService executor = Executors.newFixedThreadPool(THREAD_POOL_SIZE);

        for (int i = 0; i < 10; i++) { // Example loop to simulate multiple templates
            String docxFilePath = "/home/mgblow/gitlab/DocGen/DocGen/src/main/resources/template.docx"; // Change to actual file names
            String modifiedDocxPath = "modified_template_" + i + ".docx";
            String pdfFilePath = "output_" + i + ".pdf";

            executor.submit(() -> {
                try {
                    processTemplate(docxFilePath, modifiedDocxPath, pdfFilePath);
                } catch (IOException e) {
                    e.printStackTrace();
                }
            });
        }

        executor.shutdown();
    }

    private static void processTemplate(String docxPath, String modifiedDocxPath, String pdfPath) throws IOException {
        // Load the DOCX file
        try (FileInputStream fis = new FileInputStream(docxPath);
             XWPFDocument document = new XWPFDocument(fis)) {

            // Manipulate text
            for (XWPFParagraph paragraph : document.getParagraphs()) {
                for (XWPFRun run : paragraph.getRuns()) {
                    String text = run.getText(0);
                    if (text != null && text.contains("LOREM_IPSUM")) {
                        text = text.replace("PLACEHOLDER_TEXT", "New Text Here");
                        run.setText(text, 0);
                    }
                }
            }

            // Save the modified DOCX
            try (FileOutputStream out = new FileOutputStream(modifiedDocxPath)) {
                document.write(out);
            }

            // Convert the modified DOCX to PDF
            convertDocxToPdf(modifiedDocxPath, pdfPath);
        }
    }

    private static void convertDocxToPdf(String docxPath, String pdfPath) {
        // Simple conversion logic (may need to handle complex formats)
        try (PDDocument pdfDoc = new PDDocument()) {
            PDPage page = new PDPage(PDRectangle.A4);
            pdfDoc.addPage(page);
            PDPageContentStream contentStream = new PDPageContentStream(pdfDoc, page);

            // Here you'd read from the DOCX and write content to PDF.
            // Placeholder for PDFBox content writing logic.
            // For example, you might need to extract text and draw it:
            // contentStream.beginText();
            // contentStream.showText("Sample Text from DOCX"); // Replace with actual text extraction
            // contentStream.endText();

            contentStream.close();
            pdfDoc.save(pdfPath);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
