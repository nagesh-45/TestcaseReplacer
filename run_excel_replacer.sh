#!/bin/bash

# Set paths
FILES_DIR="Files"
OUTPUT_DIR="Output"
JAR_FILE="target/testcasereplacer-1.0-SNAPSHOT-jar-with-dependencies.jar"

# Create required directories
mkdir -p "$OUTPUT_DIR"

# Check for required files
echo "Checking files..."
ls -la "$FILES_DIR"

# Set Excel file
EXCEL_FILE="$FILES_DIR/IDPE Onus Tcs.xlsm"
CONFIG_FILE="$FILES_DIR/config.properties"

# Check if files exist
if [ ! -f "$EXCEL_FILE" ]; then
    echo "ERROR: Excel file not found: $EXCEL_FILE"
    exit 1
fi

if [ ! -f "$CONFIG_FILE" ]; then
    echo "ERROR: Config file not found: $CONFIG_FILE"
    exit 1
fi

# Create timestamp (works on both Linux and Mac)
if [[ "$OSTYPE" == "darwin"* ]]; then
    # Mac specific timestamp
    TIMESTAMP=$(date +"%Y%m%d_%H%M%S")
else
    # Linux timestamp
    TIMESTAMP=$(date +"%Y%m%d_%H%M%S")
fi

OUTPUT_FILE="$OUTPUT_DIR/IDPE Onus Tcs_${TIMESTAMP}.xlsm"

# Create temp directory (works on both Linux and Mac)
if [[ "$OSTYPE" == "darwin"* ]]; then
    # Mac specific temp directory
    TEMP_DIR=$(mktemp -d -t excel_replacer)
else
    # Linux temp directory
    TEMP_DIR=$(mktemp -d)
fi
echo "Created temporary directory: $TEMP_DIR"

# Copy files to temp directory
cp "$EXCEL_FILE" "$TEMP_DIR/input.xlsm"
cp "$CONFIG_FILE" "$TEMP_DIR/config.properties"

# Run Java
echo "Running Excel Replacer..."
echo "Java version:"
java -version
echo "Java home:"
echo "$JAVA_HOME"

java -jar "$JAR_FILE" "$TEMP_DIR/input.xlsm" "$TEMP_DIR/config.properties" "$OUTPUT_FILE"
if [ $? -ne 0 ]; then
    echo "ERROR: Java program failed"
    echo "Common issues:"
    echo "1. Check if the Excel file is not corrupted"
    echo "2. Verify the config.properties format"
    echo "3. Check Java version compatibility"
    echo "4. Ensure all required dependencies are in the JAR"
    rm -rf "$TEMP_DIR"
    exit 1
fi

# Clean up
rm -rf "$TEMP_DIR"

echo "Done!" 