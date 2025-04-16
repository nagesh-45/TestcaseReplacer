package com.example;

import java.io.BufferedReader;
import java.io.File;
import java.io.InputStreamReader;
import java.util.Scanner;

public class ExcelLauncher {
    public static void main(String[] args) {
        Scanner scanner = new Scanner(System.in);
        boolean running = true;

        while (running) {
            System.out.println("\nExcel Replacer - Menu");
            System.out.println("====================");
            System.out.println("1. Run Excel Replacer");
            System.out.println("2. Rebuild Project");
            System.out.println("3. Clean and Rebuild Project");
            System.out.println("4. Exit");
            System.out.print("\nEnter your choice (1-4): ");

            String choice = scanner.nextLine().trim();

            try {
                switch (choice) {
                    case "1":
                        runExcelReplacer();
                        break;
                    case "2":
                        runMavenCommand("install");
                        break;
                    case "3":
                        runMavenCommand("clean install");
                        break;
                    case "4":
                        running = false;
                        System.out.println("\nThank you for using Excel Replacer!");
                        break;
                    default:
                        System.out.println("\nInvalid choice. Please try again.");
                }
            } catch (Exception e) {
                System.out.println("\nError: " + e.getMessage());
            }

            if (running && !choice.equals("4")) {
                System.out.println("\nPress Enter to continue...");
                scanner.nextLine();
            }
        }
        scanner.close();
    }

    private static void runExcelReplacer() throws Exception {
        File jarFile = new File("target/testcasereplacer-1.0-SNAPSHOT.jar");
        if (!jarFile.exists()) {
            System.out.println("\nJAR file not found. Building project first...");
            runMavenCommand("install");
        }

        File configFile = new File("src/main/resources/config.properties");
        if (!configFile.exists()) {
            throw new Exception("config.properties not found in src/main/resources");
        }

        System.out.println("\nStarting Excel Replacer...");
        Process process = Runtime.getRuntime().exec("java -jar target/testcasereplacer-1.0-SNAPSHOT.jar");
        
        // Read and print the output
        BufferedReader reader = new BufferedReader(new InputStreamReader(process.getInputStream()));
        String line;
        while ((line = reader.readLine()) != null) {
            System.out.println(line);
        }
        
        int exitCode = process.waitFor();
        if (exitCode != 0) {
            throw new Exception("Program exited with code " + exitCode);
        }
    }

    private static void runMavenCommand(String command) throws Exception {
        System.out.println("\nRunning Maven command: " + command);
        Process process = Runtime.getRuntime().exec("mvn " + command);
        
        // Read and print the output
        BufferedReader reader = new BufferedReader(new InputStreamReader(process.getInputStream()));
        String line;
        while ((line = reader.readLine()) != null) {
            System.out.println(line);
        }
        
        int exitCode = process.waitFor();
        if (exitCode != 0) {
            throw new Exception("Maven command failed with exit code " + exitCode);
        }
    }
} 