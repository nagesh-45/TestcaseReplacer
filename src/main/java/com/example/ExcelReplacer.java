package com.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Properties;

public class ExcelReplacer {
    public static void main(String[] args) {
        // Get the project root directory
        String projectRoot = System.getProperty("user.dir");
        String propertiesPath = projectRoot + File.separator + "Files" + File.separator + "config.properties";
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
            fis = new FileInputStream(propertiesFile);
            properties.load(fis);
            
            // Get Excel filename from properties
            String excelFileName = properties.getProperty("excel.filename");
            if (excelFileName == null || excelFileName.trim().isEmpty()) {
                System.err.println("Excel filename not specified in config.properties");
                return;
            }
            
            String excelFilePath = projectRoot + File.separator + "Files" + File.separator + excelFileName;
            
            // Create timestamp for filename
            SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMdd_HHmmss");
            String timestamp = dateFormat.format(new Date());
            String outputFilePath = outputFolderPath + File.separator + excelFileName.replace(".xlsm", "") + "_" + timestamp + ".xlsm";

            // Create output folder if it doesn't exist
            File outputFolder = new File(outputFolderPath);
            if (!outputFolder.exists()) {
                outputFolder.mkdirs();
            }

            // Load Excel file
            File excelFile = new File(excelFilePath);
            if (!excelFile.exists()) {
                System.err.println("Excel file not found at: " + excelFilePath);
                return;
            }
            
            // Read the existing Excel file
            fis = new FileInputStream(excelFile);
            workbook = new XSSFWorkbook(fis);

            int totalReplacements = 0;
            int sheetReplacements = 0;

            System.out.println("\nStarting Excel replacement process...");
            System.out.println("Processing file: " + excelFileName);
            System.out.println("Properties loaded: " + (properties.size() - 1) + " replacement values");

            // Process each sheet
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);
                sheetReplacements = 0;
                System.out.println("\nProcessing sheet: " + sheet.getSheetName());

                // Process each row
                for (Row row : sheet) {
                    // Process each cell
                    for (Cell cell : row) {
                        if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
                            String cellValue = cell.getStringCellValue();
                            String originalValue = cellValue;
                            
                            // Replace placeholders with property values
                            for (String key : properties.stringPropertyNames()) {
                                // Skip the excel.filename property
                                if (key.equals("excel.filename")) continue;
                                
                                String placeholder = "{" + key + "}";
                                String oldValue = cellValue;
                                cellValue = cellValue.replace(placeholder, properties.getProperty(key));
                                
                                // Check if any replacements were made
                                if (!oldValue.equals(cellValue)) {
                                    cell.setCellValue(cellValue);
                                    totalReplacements++;
                                    sheetReplacements++;
                                    
                                    // Print details of the replacement
                                    System.out.println("Replacement #" + totalReplacements + ":");
                                    System.out.println("  Location: Sheet '" + sheet.getSheetName() + 
                                                     "', Row " + (row.getRowNum() + 1) + 
                                                     ", Column " + (cell.getColumnIndex() + 1));
                                    System.out.println("  Placeholder: " + placeholder);
                                    System.out.println("  Value: " + properties.getProperty(key));
                                    System.out.println("  Original: " + originalValue);
                                    System.out.println("  Updated: " + cellValue);
                                }
                            }
                        }
                    }
                }
                System.out.println("\nSheet '" + sheet.getSheetName() + "' completed: " + sheetReplacements + " replacements");
            }

            // Save the modified Excel file
            File outputFile = new File(outputFilePath);
            fos = new FileOutputStream(outputFile);
            workbook.write(fos);
            
            System.out.println("\nSummary:");
            System.out.println("Total replacements made: " + totalReplacements);
            System.out.println("Output file saved as: " + outputFile.getAbsolutePath());

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