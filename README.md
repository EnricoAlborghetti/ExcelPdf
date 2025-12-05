# ExcelPdf

A robust C# library for converting Excel files (XLSX) to PDF documents. This library leverages **NPOI** for reading Excel files and **QuestPDF** for high-quality PDF generation. It supports advanced features like merged cells, image rendering, complex styling, and dynamic content injection.

## Features

- **Excel to PDF Conversion**: Convert entire sheets or specific cell ranges to PDF.
- **Advanced Styling**: Preserves fonts, colors, borders, alignments, and text rotation.
- **Merged Cells**: Correctly handles merged cells for both text and images.
- **Image Support**: Renders images embedded in Excel and allows injecting new images into specific cells.
- **Dynamic Content**: Inject or override cell values and images at runtime without modifying the source Excel file.
- **Excel Manipulation**: Helper utilities to duplicate sheets, duplicate rows, and modify cell content programmatically.
- **Border Logic**: Implements "effective border" logic to correctly render borders defined on adjacent cells.

## Dependencies

- [NPOI](https://github.com/nissl-lab/npoi) - For reading and manipulating Excel files.
- [QuestPDF](https://github.com/QuestPDF/QuestPDF) - For generating PDF documents.

## Installation

Ensure your project references the `ExcelPdf` library. You will also need to install the NuGet packages for NPOI and QuestPDF.

```xml
<PackageReference Include="NPOI" Version="..." />
<PackageReference Inc`lude="QuestPDF" Version="..." />
```

## Usage

### Converting Excel to PDF

The `ExcelToPdfConverter` class is the main entry point for conversion.

```csharp
using ExcelPdf;
using QuestPDF.Fluent;

// 1. Initialize the converter
var converter = new ExcelToPdfConverter();

// 2. Set the input Excel file
converter.SetInputFile("path/to/input.xlsx");

// 3. (Optional) Set specific ranges to print
// Format: "SheetName!A1:D10" (specific sheet) or "A1:D10" (all sheets) or "SheetName" (entire sheet)
converter.SetPrintRange(new List<string> { "Sheet1!A1:F20", "Sheet2" });

// 4. (Optional) Inject dynamic values
// NOTE: Updates are applied by (Row, Col) coordinates to ALL printed sheets.
// The sheet name in the key is currently ignored by the injection logic.
var cellUpdates = new Dictionary<string, string>
{
    { "B2", "New Value" }, // Applies to cell B2 on every printed sheet
    { "C5", "123.45" }
};
converter.SetValues(cellUpdates);

// 5. (Optional) Inject dynamic images
// NOTE: Like values, image injections are coordinate-based and apply to all sheets.
var imageUpdates = new Dictionary<string, string>
{
    { "A1", "path/to/logo.png" }
};
converter.SetImages(imageUpdates);

// 6. Generate the PDF Document
var document = converter.Convert();

// 7. Save or Process the PDF
document.GeneratePdf("path/to/output.pdf");
```

### Manipulating Excel Files

The `ExcelHelper` class provides utilities for modifying Excel files before conversion or for other purposes.

```csharp
using ExcelPdf;

// Open an Excel file
using (var helper = new ExcelHelper("path/to/template.xlsx"))
{
    // Set cell value
    helper.SetCellValue("Sheet1", "A1", "Hello World");

    // Insert an image into a cell
    helper.SetCellImage("Sheet1", "B2", "path/to/image.png");

    // Duplicate a sheet
    helper.DuplicateSheet("TemplateSheet", "NewSheet");

    // Duplicate a row (copies styles and merged regions)
    helper.DuplicateRow("Sheet1", sourceRowIndex: 5, destRowIndex: 10);

    // Save changes
    helper.Save("path/to/modified.xlsx");
}
```

## API Reference

### `ExcelToPdfConverter`

| Method                                          | Description                                                          |
| ----------------------------------------------- | -------------------------------------------------------------------- |
| `SetInputFile(string path)`                     | Sets the path to the source Excel file.                              |
| `SetPrintRange(IEnumerable<string> ranges)`     | Specifies which sheets or ranges to include in the PDF.              |
| `SetValues(Dictionary<string, string> updates)` | Overrides cell values. Key format: `Sheet!Cell` (e.g., `Sheet1!A1`). |
| `SetImages(Dictionary<string, string> updates)` | Overrides/Inserts images. Key: `Sheet!Cell`, Value: Image Path.      |
| `SetDebug(bool debug)`                          | Enables console debug output for cell rendering details.             |
| `Convert()`                                     | Returns a QuestPDF `Document` object ready for generation.           |

### `ExcelHelper`

| Method                                                         | Description                                                       |
| -------------------------------------------------------------- | ----------------------------------------------------------------- |
| `SetCellValue(string sheet, string address, string value)`     | Sets the text value of a specific cell.                           |
| `SetCellImage(string sheet, string address, string imagePath)` | Inserts an image into a cell, handling anchors.                   |
| `DuplicateSheet(string source, string newName)`                | Creates a copy of an existing sheet.                              |
| `DuplicateRow(string sheet, int sourceRow, int destRow)`       | Copies a row's content, style, and merged regions to a new index. |
| `Save(string? outputPath)`                                     | Saves the workbook. If path is null, overwrites the original.     |
| `GetSheetNames()`                                              | Returns a list of all sheet names in the workbook.                |

### `ExcelToPdfExt`

Internal extension methods used for styling.

- **`GetEffectiveBorderWidth`**: Calculates the visible border by checking the cell and its neighbors (Top/Bottom/Left/Right), ensuring Excel grid borders are rendered correctly.
- **`ApplyTint`**: Adjusts RGB colors based on Excel's tint property.

## Limitations

- **Dynamic Injection Scope**: The `SetValues` and `SetImages` methods apply updates based on cell coordinates (Row, Column) across **all** sheets being processed. They do not currently support targeting a specific sheet by name if multiple sheets are being converted in the same pass.
- **Formula Evaluation**: While NPOI's formula evaluator is used, complex Excel formulas or those relying on unsupported functions may not evaluate correctly.
- **Chart Rendering**: Native Excel charts are not currently rendered to PDF.

## License

This project uses QuestPDF, which may require a license for commercial use. Please check [QuestPDF Licensing](https://www.questpdf.com/license/) for details.
