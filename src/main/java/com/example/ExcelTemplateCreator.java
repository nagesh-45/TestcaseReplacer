package com.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelTemplateCreator {
    public static void main(String[] args) {
        // Get the project root directory
        String projectRoot = System.getProperty("user.dir");
        String templatePath = projectRoot + "/src/main/resources/template.xlsx";
        
        // Create resources directory if it doesn't exist
        File resourcesDir = new File(projectRoot + "/src/main/resources");
        if (!resourcesDir.exists()) {
            resourcesDir.mkdirs();
        }

        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Data");
            
            // Create header row
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("Field");
            headerRow.createCell(1).setCellValue("Value");
            
            // Create data rows with placeholders
            String[][] data = {
                {"Name", "{name}"},
                {"Company", "{company}"},
                {"Date", "{date}"},
                {"Amount", "{amount}"},
                {"Description", "{description}"}
            };
            
            for (int i = 0; i < data.length; i++) {
                Row row = sheet.createRow(i + 1);
                row.createCell(0).setCellValue(data[i][0]);
                row.createCell(1).setCellValue(data[i][1]);
            }
            
            // Auto-size columns
            sheet.autoSizeColumn(0);
            sheet.autoSizeColumn(1);
            
            // Save the template
            File templateFile = new File(templatePath);
            try (FileOutputStream fileOut = new FileOutputStream(templateFile)) {
                workbook.write(fileOut);
                System.out.println("Excel template created successfully at: " + templateFile.getAbsolutePath());
            } catch (IOException e) {
                System.err.println("Error writing to file: " + e.getMessage());
                e.printStackTrace();
            }
            
        } catch (IOException e) {
            System.err.println("Error creating workbook: " + e.getMessage());
            e.printStackTrace();
        }
    }
} 