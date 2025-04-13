# Excel Placeholder Replacer

A Java tool that replaces placeholders in Excel files with values from a properties file.

## Features

- Replaces multiple occurrences of placeholders in Excel cells
- Preserves the original Excel file
- Creates timestamped output files
- Supports .xlsm (macro-enabled) Excel files
- Detailed logging of replacements

## Prerequisites

- Java 24 or higher
- Maven

## Setup

1. Clone the repository
2. Install dependencies:
   ```bash
   mvn clean install
   ```

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

1. Prepare your Excel file:
   - Use placeholders in the format `{key}` where `key` matches a property name
   - Example: `Hello {name}, welcome to {company}`

2. Run the program:
   ```bash
   mvn exec:java -Dexec.mainClass="com.example.ExcelReplacer"
   ```

3. Find the output:
   - The modified file will be created in the `output` folder
   - Filename format: `IDPE Onus Tcs_YYYYMMDD_HHMMSS.xlsm`

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