package com.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Properties;

public class ExcelReplacer {
    public static void main(String[] args) {
        // Get the project root directory
        String projectRoot = System.getProperty("user.dir");
        String propertiesPath = projectRoot + File.separator + "src" + File.separator + "main" + File.separator + "resources" + File.separator + "config.properties";
        String outputFolderPath = projectRoot + File.separator + "output";
        
        FileInputStream fis = null;
        FileOutputStream fos = null;
        Workbook workbook = null;

        try {
            // Load properties file
            Properties properties = new Properties();
            File propertiesFile = new File(propertiesPath);
            if (!propertiesFile.exists()) {
                System.err.println("Properties file not found at: " + propertiesPath);
                return;
            }
            System.out.println("Loading properties file...");
            fis = new FileInputStream(propertiesFile);
            properties.load(fis);
            
            // Get Excel filename from properties
            String excelFileName = properties.getProperty("excel.filename");
            if (excelFileName == null || excelFileName.trim().isEmpty()) {
                System.err.println("Excel filename not specified in config.properties. Please add 'excel.filename=yourfile.xlsm'");
                return;
            }
            
            String excelFilePath = projectRoot + File.separator + "src" + File.separator + "main" + File.separator + "resources" + File.separator + excelFileName;
            
            // Create timestamp for filename
            SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMdd_HHmmss");
            String timestamp = dateFormat.format(new Date());
            String outputFilePath = outputFolderPath + File.separator + excelFileName.replace(".xlsm", "") + "_" + timestamp + ".xlsm";

            System.out.println("Starting Excel replacement process...");
            System.out.println("Properties file path: " + propertiesPath);
            System.out.println("Excel file path: " + excelFilePath);
            System.out.println("Output folder path: " + outputFolderPath);
            System.out.println("Output file will be saved as: " + new File(outputFilePath).getName());

            // Create output folder if it doesn't exist
            File outputFolder = new File(outputFolderPath);
            if (!outputFolder.exists()) {
                if (outputFolder.mkdirs()) {
                    System.out.println("Created output folder at: " + outputFolderPath);
                } else {
                    System.err.println("Failed to create output folder at: " + outputFolderPath);
                    return;
                }
            }

            System.out.println("Properties loaded successfully. Found " + (properties.size() - 1) + " replacement properties.");

            // Load Excel file
            File excelFile = new File(excelFilePath);
            if (!excelFile.exists()) {
                System.err.println("Excel file not found at: " + excelFilePath);
                return;
            }
            System.out.println("Loading Excel file...");
            
            // Read the existing Excel file
            fis = new FileInputStream(excelFile);
            workbook = new XSSFWorkbook(fis);
            System.out.println("Excel file loaded successfully. Number of sheets: " + workbook.getNumberOfSheets());

            // Process each sheet
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);
                System.out.println("Processing sheet: " + sheet.getSheetName());
                
                // Process each row
                for (Row row : sheet) {
                    // Process each cell
                    for (Cell cell : row) {
                        if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
                            String cellValue = cell.getStringCellValue();
                            boolean wasModified = false;
                            
                            // Replace placeholders with property values
                            for (String key : properties.stringPropertyNames()) {
                                // Skip the excel.filename property
                                if (key.equals("excel.filename")) continue;
                                
                                String placeholder = "{" + key + "}";
                                String oldValue = cellValue;
                                cellValue = cellValue.replace(placeholder, properties.getProperty(key));
                                
                                // Check if any replacements were made
                                if (!oldValue.equals(cellValue)) {
                                    wasModified = true;
                                    System.out.println("Replaced placeholder " + placeholder + " with value: " + properties.getProperty(key));
                                }
                            }
                            
                            // Only update the cell if it was modified
                            if (wasModified) {
                                cell.setCellValue(cellValue);
                            }
                        }
                    }
                }
            }

            // Save the modified Excel file to the output folder
            System.out.println("Saving modified Excel file...");
            File outputFile = new File(outputFilePath);
            fos = new FileOutputStream(outputFile);
            workbook.write(fos);
            System.out.println("Excel file has been updated successfully at: " + outputFile.getAbsolutePath());

        } catch (IOException e) {
            System.err.println("Error processing files: " + e.getMessage());
            e.printStackTrace();
        } finally {
            // Close resources
            try {
                if (fis != null) fis.close();
                if (fos != null) fos.close();
                if (workbook != null) workbook.close();
            } catch (IOException e) {
                System.err.println("Error closing resources: " + e.getMessage());
            }
        }
    }
} 