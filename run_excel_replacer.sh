#!/bin/bash

# Set paths
PROJECT_ROOT="$(cd "$(dirname "$0")" && pwd)"
FILES_DIR="$PROJECT_ROOT/Files"
OUTPUT_DIR="$PROJECT_ROOT/Output"
JAR_FILE="$PROJECT_ROOT/target/testcasereplacer-1.0-SNAPSHOT-jar-with-dependencies.jar"

# Debug information
echo "Current directory: $(pwd)"
echo "Project root: $PROJECT_ROOT"
echo "Files directory: $FILES_DIR"
echo "Output directory: $OUTPUT_DIR"
echo "JAR file: $JAR_FILE"

# Create output directory if it doesn't exist
mkdir -p "$OUTPUT_DIR"

# Check for required files
echo "Checking for required files in $FILES_DIR..."
if [ ! -d "$FILES_DIR" ]; then
    echo "ERROR: Files directory not found at $FILES_DIR"
    exit 1
fi

# List all files in the Files directory
echo "Available files in $FILES_DIR:"
ls -la "$FILES_DIR"

# Check for Excel file
EXCEL_FILE=$(find "$FILES_DIR" -name "*.xlsm" | head -n 1)
if [ -z "$EXCEL_FILE" ]; then
    echo "ERROR: No Excel file (.xlsm) found in $FILES_DIR"
    exit 1
fi
echo "Found Excel file: $(basename "$EXCEL_FILE")"
echo "Full path: $EXCEL_FILE"

# Check for config file
CONFIG_FILE=$(find "$FILES_DIR" -name "*.properties" | head -n 1)
if [ -z "$CONFIG_FILE" ]; then
    echo "ERROR: No config file (.properties) found in $FILES_DIR"
    exit 1
fi
echo "Found config file: $(basename "$CONFIG_FILE")"
echo "Full path: $CONFIG_FILE"

# Check for Java
if ! command -v java &> /dev/null; then
    echo "ERROR: Java is not installed or not found in PATH"
    exit 1
fi

# Display Java version
echo "Java version:"
java -version

# Display Java home
echo "Java home:"
echo "$JAVA_HOME"

# Check for JAR file
if [ ! -f "$JAR_FILE" ]; then
    echo "ERROR: JAR file not found at $JAR_FILE"
    echo "Please run 'mvn clean package' to build the JAR file"
    exit 1
fi
echo "JAR file exists at: $JAR_FILE"
echo "JAR file size: $(du -h "$JAR_FILE" | cut -f1)"

# Create timestamp for output file
TIMESTAMP=$(date +"%Y%m%d_%H%M%S")
OUTPUT_FILE="$OUTPUT_DIR/$(basename "$EXCEL_FILE" .xlsm)_${TIMESTAMP}.xlsm"

echo "Running Excel Replacer..."
echo "Input Excel: $EXCEL_FILE"
echo "Config File: $CONFIG_FILE"
echo "Output File: $OUTPUT_FILE"

# Create temporary directory for processing
TEMP_DIR=$(mktemp -d)
echo "Created temporary directory: $TEMP_DIR"

# Copy files to temporary directory
cp "$EXCEL_FILE" "$TEMP_DIR/input.xlsm"
cp "$CONFIG_FILE" "$TEMP_DIR/config.properties"

# Run Java program with debug options
echo "Running Java program with debug options..."
echo "Java command: java -jar \"$JAR_FILE\" \"$TEMP_DIR/input.xlsm\" \"$TEMP_DIR/config.properties\" \"$OUTPUT_FILE\""
if ! java -jar "$JAR_FILE" "$TEMP_DIR/input.xlsm" "$TEMP_DIR/config.properties" "$OUTPUT_FILE"; then
    echo "ERROR: Java program failed with exit code $?"
    echo "Please check the Java output above for details."
    echo "Common issues:"
    echo "1. Check if the Excel file is not corrupted"
    echo "2. Verify the config.properties format"
    echo "3. Check Java version compatibility"
    echo "4. Ensure all required dependencies are in the JAR"
    rm -rf "$TEMP_DIR"
    exit 1
fi

# Clean up temporary directory
rm -rf "$TEMP_DIR"

echo "Excel Replacer completed successfully!"
echo "Output file created at: $OUTPUT_FILE"

# Remove existing JARs
rm -f target/*.jar

# Rebuild with Maven
mvn clean package 