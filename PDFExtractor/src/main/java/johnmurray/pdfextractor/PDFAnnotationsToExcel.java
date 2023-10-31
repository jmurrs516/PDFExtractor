/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package johnmurray.pdfextractor;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.interactive.annotation.PDAnnotation;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class PDFAnnotationsToExcel {
    public void Search(String inputPath, String exportPath) {
        File folder = new File(inputPath); // Replace with the path to your folder of PDFs
        File[] pdfFiles = folder.listFiles((dir, name) -> name.toLowerCase().endsWith(".pdf"));

        if (pdfFiles == null) {
            System.err.println("No PDF files found in the specified folder.");
            return;
        }

        try {
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("Annotations");
            int rowIndex = 0;

            // Create headers for the columns
            Row headerRow = sheet.createRow(rowIndex++);
            headerRow.createCell(0).setCellValue("File Name");
            headerRow.createCell(1).setCellValue("Comment Count");
            headerRow.createCell(2).setCellValue("Comment");

            for (File pdfFile : pdfFiles) {
                PDDocument document = PDDocument.load(pdfFile);
                int annotationCount = 0;

                for (PDPage page : document.getPages()) {
                    List<PDAnnotation> annotations = page.getAnnotations();
                    for (PDAnnotation annotation : annotations) {
                        String annotationText = annotation.getContents();
                        if (annotationText != null && !annotationText.isEmpty()) {
                            Row row = sheet.createRow(rowIndex++);
                            row.createCell(0).setCellValue(pdfFile.getName()); // File Name
                            row.createCell(1).setCellValue(annotationCount + 1); // Comment Count
                            row.createCell(2).setCellValue(annotationText); // Comment
                            annotationCount++;
                        }
                    }
                }

                // Create a cell for the total annotation count per PDF file
                Row countRow = sheet.createRow(rowIndex++);
                countRow.createCell(0).setCellValue(pdfFile.getName());
                countRow.createCell(1).setCellValue("Total Comments");
                countRow.createCell(2).setCellValue(annotationCount);

                document.close();
            }

            // Specify the path to your desired folder and save the Excel file
            String outputFolderPath = exportPath; // Change to your specific folder path
            String excelFilePath = outputFolderPath + File.separator + "annotations.xlsx";

            FileOutputStream fileOut = new FileOutputStream(excelFilePath);
            workbook.write(fileOut);
            fileOut.close();
            
            System.out.println("Excel file saved to: " + excelFilePath);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
