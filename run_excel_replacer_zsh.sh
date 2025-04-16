#!/usr/bin/zsh

# Zsh-specific Excel Replacer Script

# Check if running in Zsh
if [ -z "$ZSH_VERSION" ]; then
    echo "This script must be run with Zsh"
    echo "Please run with: /usr/bin/zsh $0"
    exit 1
fi

# Enable error handling
set -e
set -o pipefail

# Set paths using Zsh array syntax
typeset -A paths=(
    FILES_DIR "Files"
    OUTPUT_DIR "Output"
    JAR_FILE "target/testcasereplacer-1.0-SNAPSHOT-jar-with-dependencies.jar"
)

# Create required directories
mkdir -p "${paths[OUTPUT_DIR]}"

# Check for required files using Zsh globbing
echo "Checking files in ${paths[FILES_DIR]}..."
ls -la "${paths[FILES_DIR]}"

# Set Excel file
typeset -A files=(
    EXCEL_FILE "${paths[FILES_DIR]}/IDPE Onus Tcs.xlsm"
    CONFIG_FILE "${paths[FILES_DIR]}/config.properties"
)

# Check if files exist using Zsh test
for file in "${(@k)files}"; do
    if [[ ! -f "${files[$file]}" ]]; then
        print -P "%F{red}ERROR: File not found: ${files[$file]}%f"
        exit 1
    fi
done

# Create timestamp using Zsh date format
TIMESTAMP=$(strftime "%Y%m%d_%H%M%S")
OUTPUT_FILE="${paths[OUTPUT_DIR]}/IDPE Onus Tcs_${TIMESTAMP}.xlsm"

# Create temp directory using Zsh temp file handling
TEMP_DIR=$(mktemp -d)
print -P "%F{green}Created temporary directory: ${TEMP_DIR}%f"

# Copy files to temp directory using Zsh array
typeset -A temp_files=(
    input.xlsm "${files[EXCEL_FILE]}"
    config.properties "${files[CONFIG_FILE]}"
)

for dest src in "${(@kv)temp_files}"; do
    cp "$src" "${TEMP_DIR}/${dest}"
done

# Check Java installation using Zsh command substitution
if ! (( $+commands[java] )); then
    print -P "%F{red}ERROR: Java is not installed%f"
    print -P "%F{yellow}Please install Java using one of these methods:%f"
    print -P "  %F{blue}Ubuntu/Debian:%f sudo apt install default-jdk"
    print -P "  %F{blue}CentOS/RHEL:%f sudo yum install java-11-openjdk"
    print -P "  %F{blue}Fedora:%f sudo dnf install java-11-openjdk"
    print -P "  %F{blue}Arch Linux:%f sudo pacman -S jdk-openjdk"
    exit 1
fi

# Display Java information with Zsh colors
print -P "%F{blue}Java version:%f"
java -version
print -P "%F{blue}Java home:%f"
print -P "%F{green}${JAVA_HOME}%f"

# Run Java with Zsh error handling
print -P "%F{blue}Running Excel Replacer...%f"
if ! java -jar "${paths[JAR_FILE]}" "${TEMP_DIR}/input.xlsm" "${TEMP_DIR}/config.properties" "$OUTPUT_FILE"; then
    print -P "%F{red}ERROR: Java program failed%f"
    print -P "%F{yellow}Common issues on Linux with Zsh:%f"
    print -P "1. Check if the Excel file is not corrupted"
    print -P "2. Verify the config.properties format"
    print -P "3. Check Java version compatibility"
    print -P "4. Ensure all required dependencies are in the JAR"
    print -P "5. Check file permissions"
    print -P "6. Verify Zsh environment variables"
    rm -rf "$TEMP_DIR"
    exit 1
fi

# Clean up
rm -rf "$TEMP_DIR"

print -P "%F{green}Excel Replacer completed successfully!%f"
print -P "%F{blue}Output file created at: ${OUTPUT_FILE}%f" 