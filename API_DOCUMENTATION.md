# ExcelExporter API Documentation

## Overview

ExcelExporter is a .NET 8.0 command-line tool for extracting and exporting Excel workbook contents to various formats. It supports exporting worksheet data, VBA code, custom ribbons, and other Excel components to CSV and JSON formats for version control and analysis.

## Table of Contents

1. [Command Line Interface](#command-line-interface)
2. [Core Classes](#core-classes)
3. [Data Models](#data-models)
4. [Utility Classes](#utility-classes)
5. [Usage Examples](#usage-examples)
6. [Error Handling](#error-handling)
7. [Configuration](#configuration)

## Command Line Interface

### Basic Usage

```bash
ExcelExporter.exe <input_file> [arguments]
```

### Arguments

| Argument | Short | Description | Example |
|----------|-------|-------------|---------|
| `--help` | `-h` | Display help message and usage instructions | `ExcelExporter.exe --help` |
| `--version` | `-v` | Display program version | `ExcelExporter.exe --version` |
| `--outputDir` | `-od` | Specify output directory for extracted files | `ExcelExporter.exe input.xlsx --outputDir ./output` |
| `--outputName` | `-on` | Specify custom name for output folder | `ExcelExporter.exe input.xlsx --outputName my_export` |
| `--fileType` | `-ft` | Specify how to handle the workbook (xlsx, xlsm, etc.) | `ExcelExporter.exe input.xlsx --fileType xlsm` |
| `--saveType` | `-st` | Specify output format (csv, json) | `ExcelExporter.exe input.xlsx --saveType json` |

### Supported File Types

- **xlsx** - Standard Excel workbook
- **xltx** - Excel template
- **xlsm** - Macro-enabled Excel workbook
- **xltm** - Macro-enabled Excel template
- **xlam** - Excel add-in

### Supported Save Types

- **csv** - Comma-separated values format
- **json** - JSON format (default)

## Core Classes

### ArgumentsHandler

The main class responsible for parsing and validating command-line arguments.

#### Public Methods

##### `HandleArguments(string[] args, out ParsedArguments parsedArgs, out string message)`

Parses command-line arguments and populates a `ParsedArguments` object.

**Parameters:**
- `args` - Array of command-line arguments
- `parsedArgs` - Output parameter containing parsed arguments
- `message` - Output parameter containing any error messages

**Returns:** `bool` - `true` if parsing succeeded, `false` otherwise

**Example:**
```csharp
string[] args = { "workbook.xlsx", "--outputDir", "./output", "--fileType", "xlsx" };
if (ArgumentsHandler.HandleArguments(args, out ParsedArguments parsed, out string message))
{
    // Use parsed arguments
    Console.WriteLine($"Input: {parsed.InputFilePath}");
    Console.WriteLine($"Output: {parsed.OutputDir}");
}
else
{
    Console.WriteLine($"Error: {message}");
}
```

### CSVHandler

Handles CSV file generation from 2D arrays.

#### Public Methods

##### `WriteToCSV(object[,] array, string outputDir, string fileName)`

Writes a 2D array to a CSV file with proper escaping for special characters.

**Parameters:**
- `array` - 2D object array containing the data to export
- `outputDir` - Directory where the CSV file will be saved
- `fileName` - Name of the file (without extension)

**Example:**
```csharp
object[,] data = new object[3, 2] 
{
    { "Name", "Age" },
    { "John", 25 },
    { "Jane", 30 }
};

CSVHandler.WriteToCSV(data, "./output", "employees");
// Creates: ./output/employees.csv
```

### GenerateRequiredDirectories

Creates the necessary directory structure for exported files.

#### Public Methods

##### `Run(string inputFilePath, string fileType, string savePath)`

Generates the required directory structure based on file type.

**Parameters:**
- `inputFilePath` - Path to the input Excel file
- `fileType` - Type of Excel file (determines which directories to create)
- `savePath` - Base path where directories will be created

**Directory Structure Created:**
```
output_folder/
├── Sheets/           # Worksheet data (CSV files)
├── VBA/              # VBA code (for macro-enabled files)
│   ├── Microsoft Excel Objects/
│   ├── Modules/
│   ├── Classes/
│   └── Forms/
└── RibbonX/          # Custom ribbon XML (for macro-enabled files)
    └── Icons/
```

**Example:**
```csharp
GenerateRequiredDirectories.Run("workbook.xlsm", "xlsm", "./exports/workbook_xlsm");
```

### VBA_Handling

Handles extraction of VBA code from Excel workbooks.

#### Public Methods

##### `SetFileType(string fileType)`

Sets the file extension for exported VBA files.

**Parameters:**
- `fileType` - File extension (e.g., "vba", "bas")

##### `ExportWorksheetVBA(Excel.Worksheet worksheet, string outputDir)`

Exports VBA code from a specific worksheet.

**Parameters:**
- `worksheet` - Excel worksheet object
- `outputDir` - Directory to save VBA files

**Example:**
```csharp
VBA_Handling vbaHandler = new VBA_Handling();
vbaHandler.SetFileType("vba");
vbaHandler.ExportWorksheetVBA(worksheet, "./output/VBA/Microsoft Excel Objects");
```

##### `ExportThisWorkbookVBA(Excel.Workbook workbook, string outputDir)`

Exports VBA code from the ThisWorkbook module.

**Parameters:**
- `workbook` - Excel workbook object
- `outputDir` - Directory to save VBA files

##### `ExportModulesVBA(Excel.Workbook workbook, string outputDir)`

Exports all standard VBA modules.

**Parameters:**
- `workbook` - Excel workbook object
- `outputDir` - Directory to save VBA files

##### `ExportClassesVBA(Excel.Workbook workbook, string outputDir)`

Exports all VBA class modules.

**Parameters:**
- `workbook` - Excel workbook object
- `outputDir` - Directory to save VBA files

##### `ExportFormsVBA(Excel.Workbook workbook, string outputDir)`

Exports VBA user forms (creates directories but doesn't export .frx files yet).

**Parameters:**
- `workbook` - Excel workbook object
- `outputDir` - Directory to save VBA files

### Ribbon_Handling

Handles extraction of custom ribbon XML from Excel workbooks.

#### Public Methods

##### `ExportRibbonXML(string tempWorkbookPath, string outputDir)`

Extracts custom ribbon XML files from an Excel workbook.

**Parameters:**
- `tempWorkbookPath` - Path to a temporary copy of the workbook
- `outputDir` - Directory to save ribbon XML files

**Example:**
```csharp
Ribbon_Handling ribbonHandler = new Ribbon_Handling();
ribbonHandler.ExportRibbonXML("./temp/workbook.xlsm", "./output/RibbonX");
```

## Data Models

### ParsedArguments

Contains all parsed command-line arguments.

#### Properties

| Property | Type | Description |
|----------|------|-------------|
| `InputFilePath` | `string?` | Path to the input Excel file |
| `FileType` | `string?` | Type of Excel file (xlsx, xlsm, etc.) |
| `OutputDir` | `string?` | Output directory for extracted files |
| `OutputName` | `string?` | Custom name for output folder |
| `SaveType` | `string?` | Output format (csv, json) |

**Example:**
```csharp
ParsedArguments args = new ParsedArguments
{
    InputFilePath = "workbook.xlsx",
    FileType = "xlsx",
    OutputDir = "./exports",
    OutputName = "my_export",
    SaveType = "csv"
};
```

### Command

Represents a command-line command with its handler.

#### Properties

| Property | Type | Description |
|----------|------|-------------|
| `Key` | `string` | Command identifier |
| `Aliases` | `string` | Command aliases (e.g., "--help, -h") |
| `Description` | `string` | Command description |
| `Handler` | `Function` | Delegate for command handling |

#### Function Delegate

```csharp
public delegate bool Function(
    string[] args,
    ParsedArguments parsedArguments,
    out string handlerOutput
);
```

## Utility Classes

### GlobalDictionaries

Contains global configuration and validation data.

#### Static Properties

##### `AllowedFileTypes`

HashSet containing supported Excel file extensions.

```csharp
public static HashSet<string> AllowedFileTypes = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
{
    "xlsx", "xltx", "xlsm", "xltm", "xlam"
};
```

##### `AllowedSaveTypes`

HashSet containing supported output formats.

```csharp
public static HashSet<string> AllowedSaveTypes = new HashSet<string>
{
    "csv", "json"
};
```

### JSON_Handler

Handles JSON file generation (currently incomplete implementation).

#### Public Methods

##### `Write(object[,] array, string outputDir, string filename)`

Writes a 2D array to a JSON file (implementation in progress).

**Parameters:**
- `array` - 2D object array containing the data
- `outputDir` - Directory to save the JSON file
- `filename` - Name of the file (without extension)

## Usage Examples

### Basic Excel Export

```bash
# Export a standard Excel workbook
ExcelExporter.exe workbook.xlsx

# Export with custom output directory
ExcelExporter.exe workbook.xlsx --outputDir ./exports

# Export macro-enabled workbook with VBA extraction
ExcelExporter.exe workbook.xlsm --fileType xlsm --outputDir ./exports
```

### Advanced Usage

```bash
# Export with custom folder name and JSON format
ExcelExporter.exe data.xlsx --outputName "quarterly_report" --saveType json --outputDir ./reports

# Export macro-enabled template
ExcelExporter.exe template.xltm --fileType xltm --outputDir ./templates
```

### Programmatic Usage

```csharp
using ExcelExporter.Classes;

// Parse command line arguments
string[] args = { "workbook.xlsx", "--outputDir", "./output", "--fileType", "xlsx" };
if (ArgumentsHandler.HandleArguments(args, out ParsedArguments parsed, out string message))
{
    // Create directory structure
    GenerateRequiredDirectories.Run(parsed.InputFilePath, parsed.FileType, "./output");
    
    // Export CSV data
    object[,] worksheetData = GetWorksheetData(); // Your method to get data
    CSVHandler.WriteToCSV(worksheetData, "./output/Sheets", "Sheet1_display");
    
    // Export VBA if macro-enabled
    if (parsed.FileType.Contains("m"))
    {
        VBA_Handling vbaHandler = new VBA_Handling();
        // ... VBA export logic
    }
}
```

## Error Handling

The application includes comprehensive error handling:

### Argument Validation

- Validates file types against `GlobalDictionaries.AllowedFileTypes`
- Validates save types against `GlobalDictionaries.AllowedSaveTypes`
- Provides clear error messages for invalid arguments

### Excel COM Object Management

- Proper cleanup of Excel COM objects using `Marshal.FinalReleaseComObject`
- Garbage collection after cleanup
- Exception handling for Excel operations

### File Operations

- Directory creation with error handling
- File writing with proper encoding
- CSV escaping for special characters

## Configuration

### Dependencies

The project requires the following NuGet packages:

- `DocumentFormat.OpenXml` (3.3.0)
- `Microsoft.Office.Interop.Excel` (15.0.4795.1001)
- `MicrosoftOfficeCore` (15.0.0)
- `ThammimTech.Microsoft.Vbe.Interop` (15.0.0)

### System Requirements

- .NET 8.0 Runtime
- Microsoft Excel (for COM interop)
- Windows OS (due to Excel COM requirements)

### Project Structure

```
ExcelExporter/
├── Classes/
│   ├── ArgumentsHandler.cs      # Command-line argument parsing
│   ├── Commands.cs              # Command definitions
│   ├── CSV Handler.cs           # CSV export functionality
│   ├── Generate Required Directories.cs  # Directory creation
│   ├── JSON Handler.cs          # JSON export (incomplete)
│   ├── ParsedArguments.cs      # Argument data model
│   ├── Ribbon_Handling.cs      # Ribbon XML extraction
│   └── VBA_Handling.cs         # VBA code extraction
├── Commands/
│   └── Help.cs                 # Help command implementation
├── Resources/
│   └── GlobalDictionaries.cs   # Global configuration
├── Program.cs                  # Main entry point
└── ExcelExporter.csproj       # Project file
```

## Notes

- The application does not execute VBA code during extraction for security reasons
- Custom ribbon extraction requires Office RibbonX Editor compatibility
- JSON export functionality is currently incomplete
- VBA form export (.frx files) is not yet implemented
- The tool is designed for version control and analysis of Excel workbooks

## Roadmap

Future enhancements may include:
- Complete JSON export implementation
- VBA form file extraction
- Power Query export
- Table export
- Chart/image export
- Additional output formats