package com.doc.DocGen.docx;

import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPictureData;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.*;

public class SimpleDocxToPdfConverter {

    public static void main(String[] args) {
        try {
            // Load the DOCX file
            FileInputStream fis = new FileInputStream("/home/mgblow/gitlab/DocGen/DocGen/src/main/resources/template.docx");
            XWPFDocument document = new XWPFDocument(fis);

            // Manipulate text
            for (XWPFParagraph paragraph : document.getParagraphs()) {
                for (XWPFRun run : paragraph.getRuns()) {
                    String text = run.getText(0);
                    if (text != null && text.contains("LOREM_IPSUM")) {
                        text = text.replace("LOREM_IPSUM", "مجتبی اسدی");
                        run.setText(text, 0);
                    }
                }
            }

            // Manipulate image (if you need to replace it)
            for (XWPFPictureData pictureData : document.getAllPictures()) {
                // For example, to remove or replace, you'd manipulate accordingly
                // This is a placeholder for your image manipulation logic
                FileInputStream is = new FileInputStream("/home/mgblow/gitlab/DocGen/DocGen/src/main/resources/image.jpg");
                byte[] bytes = IOUtils.toByteArray(is);
                replacePictureData(pictureData, bytes);
            }

            // Save the modified DOCX
            FileOutputStream out = new FileOutputStream("/home/mgblow/gitlab/DocGen/DocGen/src/main/resources/modified_template.docx");
            document.write(out);
            out.close();
            document.close();

            // Convert the modified DOCX to PDF
            convertToPDF("/home/mgblow/gitlab/DocGen/DocGen/src/main/resources/modified_template.docx", "/home/mgblow/gitlab/DocGen/DocGen/src/main/resources/output.pdf");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    static void replacePictureData(XWPFPictureData source, byte[] data) {
        try (ByteArrayInputStream in = new ByteArrayInputStream(data);
             OutputStream out = source.getPackagePart().getOutputStream();
        ) {
            byte[] buffer = new byte[2048];
            int length;
            while ((length = in.read(buffer)) > 0) {
                out.write(buffer, 0, length);
            }
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    public static void convertToPDF(String docPath, String pdfPath) {
        try {
            //taking input from docx file
            InputStream doc = new FileInputStream(new File(docPath));
            //process for creating pdf started
            XWPFDocument document = new XWPFDocument(doc);
            PdfOptions options = PdfOptions.create();
            OutputStream out = new FileOutputStream(new File(pdfPath));
            PdfConverter.getInstance().convert(document, out, options);
        } catch (IOException ex) {
            System.out.println(ex.getMessage());
        }
    }

}
