#!/bin/bash

# Get the directory where the script is located
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" &> /dev/null && pwd )"

# Check if Maven is installed
if ! command -v mvn &> /dev/null; then
    echo "Maven is not installed. Please install Maven first."
    exit 1
fi

# Check if Java is installed
if ! command -v java &> /dev/null; then
    echo "Java is not installed. Please install Java first."
    exit 1
fi

# Create a desktop shortcut if it doesn't exist
if [ ! -f ~/Desktop/Run\ Excel\ Replacer.command ]; then
    echo "#!/bin/bash" > ~/Desktop/Run\ Excel\ Replacer.command
    echo "cd \"$SCRIPT_DIR\"" >> ~/Desktop/Run\ Excel\ Replacer.command
    echo "./run.sh" >> ~/Desktop/Run\ Excel\ Replacer.command
    chmod +x ~/Desktop/Run\ Excel\ Replacer.command
    echo "Created desktop shortcut"
fi

# Run the program
echo "Starting Excel Replacer..."
cd "$SCRIPT_DIR"
mvn exec:java -Dexec.mainClass="com.example.ExcelReplacer"

# Check if the program ran successfully
if [ $? -eq 0 ]; then
    echo "Program completed successfully!"
    # Open the output folder
    open "$SCRIPT_DIR/output"
else
    echo "Program failed to run. Please check the error messages above."
fi 