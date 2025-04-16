#!/bin/bash

# Set paths
PROJECT_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
FILES_DIR="$PROJECT_ROOT/Files"
OUTPUT_DIR="$PROJECT_ROOT/Output"
JAR_FILE="$PROJECT_ROOT/target/testcasereplacer-1.0-SNAPSHOT.jar"

echo "Checking setup..."
echo "================"
echo

# Check current directory
echo "Current Directory:"
echo "$PROJECT_ROOT"
echo

# Check Files folder
echo "Checking Files folder..."
if [ -d "$FILES_DIR" ]; then
    echo "Files folder exists at: $FILES_DIR"
    echo "Contents of Files folder:"
    ls -la "$FILES_DIR"
else
    echo "ERROR: Files folder not found at: $FILES_DIR"
fi
echo

# Check Excel file
echo "Checking Excel file..."
EXCEL_FOUND=$(find "$FILES_DIR" -name "*.xlsm" | head -n 1)
if [ -n "$EXCEL_FOUND" ]; then
    echo "Found Excel file: $(basename "$EXCEL_FOUND")"
else
    echo "ERROR: No Excel file (.xlsm) found in Files folder"
fi
echo

# Check config file
echo "Checking config file..."
CONFIG_FOUND=$(find "$FILES_DIR" -name "*.properties" | head -n 1)
if [ -n "$CONFIG_FOUND" ]; then
    echo "Found config file: $(basename "$CONFIG_FOUND")"
else
    echo "ERROR: No config file (.properties) found in Files folder"
fi
echo

# Check JAR file
echo "Checking JAR file..."
if [ -f "$JAR_FILE" ]; then
    echo "JAR file exists at: $JAR_FILE"
else
    echo "ERROR: JAR file not found at: $JAR_FILE"
fi
echo

# Check Java installation
echo "Checking Java installation..."
if command -v java &> /dev/null; then
    echo "Java is installed and available"
    java -version
else
    echo "ERROR: Java is not installed or not found in PATH"
fi
echo

# Check Output folder
echo "Checking Output folder..."
if [ -d "$OUTPUT_DIR" ]; then
    echo "Output folder exists at: $OUTPUT_DIR"
    echo "Current contents:"
    ls -la "$OUTPUT_DIR"
else
    echo "Output folder will be created when running the program"
fi
echo

echo "================"
echo "Setup Check Complete"
echo
echo "If you see any ERROR messages above, please fix those issues first."
echo 