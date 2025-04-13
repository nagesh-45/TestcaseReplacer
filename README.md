# Excel Placeholder Replacer

A Java tool that replaces placeholders in Excel files with values from a properties file.

## Features

- Replaces multiple occurrences of placeholders in Excel cells
- Preserves the original Excel file
- Creates timestamped output files
- Supports .xlsm (macro-enabled) Excel files
- Detailed logging of replacements
- Works on both Windows and macOS

## Prerequisites

- Java 24 or higher
- Maven

### Windows Setup
1. Install Java from: https://www.oracle.com/java/technologies/downloads/
2. Install Maven from: https://maven.apache.org/download.cgi
3. Add Java and Maven to your system PATH

### macOS Setup
1. Install Java using Homebrew: `brew install openjdk`
2. Install Maven using Homebrew: `brew install maven`

## Configuration

1. Place your Excel file in `src/main/resources/`
2. Create a `config.properties` file in `src/main/resources/` with your key-value pairs:

```properties
# Example properties
name=John Doe
company=Acme Inc
date=2024-03-20
amount=1000.50
description=Sample project
```

## Usage

### Windows
1. Double-click `run.bat` in the project folder
   - Or use the shortcut created on your desktop
2. The program will run and open the output folder when done

### macOS
1. Double-click `Run Excel Replacer.command` on your desktop
   - Or run `./run.sh` from Terminal
2. The program will run and open the output folder when done

## Example

If your Excel cell contains:
```
Hello {name}, welcome to {company}. {name} is our valued customer.
```

And your properties file has:
```
name=John
company=ABC
```

The output will be:
```
Hello John, welcome to ABC. John is our valued customer.
```

## Notes

- The original Excel file remains unchanged
- Each run creates a new timestamped file
- The program logs all replacements to the console
- Make sure the Excel file is not open in Excel when running the program

## Troubleshooting

If you encounter issues:
1. Check that the Excel file is in the correct location
2. Verify the properties file format
3. Ensure the Excel file is not open in Excel
4. Check the console output for error messages
5. Verify Java and Maven are installed and in your PATH 