using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using NPOI.SS.Formula.Functions;

namespace ExcelPdf
{
    /// <summary>
    /// Handles the conversion of Excel files to PDF documents using NPOI and QuestPDF.
    /// </summary>
    public class ExcelToPdfConverter
    {
        private bool _debug = false;
        private string? _inputPath = null;
        private IEnumerable<string>? _printRanges = null;
        private Dictionary<string, string>? _cellUpdates = null;
        private Dictionary<string, string>? _imageUpdates = null;

        public ExcelToPdfConverter()
        {

        }

        /// <summary>
        /// Enables or disables debug output to the console.
        /// </summary>
        /// <param name="debug">If set to true, prints detailed cell rendering information.</param>
        public void SetDebug(bool debug)
        {
            _debug = debug;
        }

        /// <summary>
        /// Sets the path to the input Excel file.
        /// </summary>
        /// <param name="inputPath">The absolute or relative path to the .xlsx file.</param>
        public void SetInputFile(string inputPath)
        {
            _inputPath = inputPath;
        }
        /// <summary>
        /// Specifies which sheets or cell ranges to include in the PDF.
        /// </summary>
        /// <param name="printRanges">
        /// A collection of strings representing ranges (e.g., "Sheet1!A1:B2") or sheet names (e.g., "Sheet1").
        /// If null or empty, all visible sheets are converted.
        /// </param>
        public void SetPrintRange(IEnumerable<string> printRanges)
        {
            _printRanges = printRanges;
        }

        /// <summary>
        /// Sets dynamic image replacements or insertions.
        /// </summary>
        /// <param name="imageUpdates">
        /// A dictionary where the key is the cell address (e.g., "A1") and the value is the path to the image file.
        /// Note: The update applies to the specified cell coordinates on ALL processed sheets.
        /// </param>
        public void SetImages(Dictionary<string, string> imageUpdates)
        {
            _imageUpdates = imageUpdates;
        }

        /// <summary>
        /// Sets dynamic cell value updates.
        /// </summary>
        /// <param name="cellUpdates">
        /// A dictionary where the key is the cell address (e.g., "B2") and the value is the new text content.
        /// Note: The update applies to the specified cell coordinates on ALL processed sheets.
        /// </param>
        public void SetValues(Dictionary<string, string> cellUpdates)
        {
            _cellUpdates = cellUpdates;
        }

        /// <summary>
        /// Generates the PDF document based on the configured input and settings.
        /// </summary>
        /// <returns>A QuestPDF Document instance that can be saved or further processed.</returns>
        /// <exception cref="FileNotFoundException">Thrown if the input file does not exist.</exception>
        public Document Convert()
        {
            if (!File.Exists(_inputPath))
            {
                throw new FileNotFoundException($"Input file not found: {_inputPath}");
            }

            using var stream = new FileStream(_inputPath, FileMode.Open, FileAccess.Read);
            var workbook = WorkbookFactory.Create(stream);
            var evaluator = workbook.GetCreationHelper().CreateFormulaEvaluator();

            // Parse cell updates
            Dictionary<(int Row, int Col), string>? parsedUpdates = null;
            if (_cellUpdates != null)
            {
                parsedUpdates = new Dictionary<(int Row, int Col), string>();
                foreach (var kvp in _cellUpdates)
                {
                    var cellRef = new NPOI.SS.Util.CellReference(kvp.Key);
                    parsedUpdates[(cellRef.Row, cellRef.Col)] = kvp.Value;
                }
            }

            // Parse image updates
            Dictionary<(int Row, int Col), byte[]>? parsedImageUpdates = null;
            if (_imageUpdates != null)
            {
                parsedImageUpdates = new Dictionary<(int Row, int Col), byte[]>();
                foreach (var kvp in _imageUpdates)
                {
                    var cellRef = new NPOI.SS.Util.CellReference(kvp.Key);
                    if (File.Exists(kvp.Value))
                    {
                        parsedImageUpdates[(cellRef.Row, cellRef.Col)] = File.ReadAllBytes(kvp.Value);
                    }
                    else
                    {
                        Console.WriteLine($"Warning: Image file not found for cell {kvp.Key}: {kvp.Value}");
                    }
                }
            }

            return Document.Create(container =>
            {
                if (_printRanges != null && _printRanges.Any())
                {
                    foreach (var rangeString in _printRanges)
                    {
                        string? targetSheetName = null;
                        NPOI.SS.Util.CellRangeAddress? targetRange = null;

                        var parts = rangeString.Split('!');
                        if (parts.Length == 2)
                        {
                            targetSheetName = parts[0];
                            targetRange = NPOI.SS.Util.CellRangeAddress.ValueOf(parts[1]);
                        }
                        else
                        {
                            // Check if the string matches a sheet name (case-insensitive)
                            string? matchedSheetName = null;
                            for (int i = 0; i < workbook.NumberOfSheets; i++)
                            {
                                if (workbook.GetSheetName(i).Equals(parts[0], StringComparison.OrdinalIgnoreCase))
                                {
                                    matchedSheetName = workbook.GetSheetName(i);
                                    break;
                                }
                            }

                            if (matchedSheetName != null)
                            {
                                targetSheetName = matchedSheetName;
                                targetRange = null; // Full sheet
                            }
                            else
                            {
                                targetRange = NPOI.SS.Util.CellRangeAddress.ValueOf(parts[0]);
                            }
                        }

                        foreach (int sheetIndex in Enumerable.Range(0, workbook.NumberOfSheets))
                        {
                            var sheet = workbook.GetSheetAt(sheetIndex);

                            // If a specific sheet is requested, skip others
                            if (targetSheetName != null && !sheet.SheetName.Equals(targetSheetName, StringComparison.OrdinalIgnoreCase))
                            {
                                continue;
                            }

                            if (workbook.IsSheetHidden(sheetIndex)) continue;

                            AddPage(container, sheet, evaluator, targetRange, parsedUpdates, parsedImageUpdates);
                        }
                    }
                }
                else
                {
                    // Default behavior: all sheets, full content
                    foreach (int sheetIndex in Enumerable.Range(0, workbook.NumberOfSheets))
                    {
                        var sheet = workbook.GetSheetAt(sheetIndex);
                        if (workbook.IsSheetHidden(sheetIndex)) continue;

                        AddPage(container, sheet, evaluator, null, parsedUpdates, parsedImageUpdates);
                    }
                }
            });
        }



        private void AddPage(IDocumentContainer container, ISheet sheet, IFormulaEvaluator evaluator, NPOI.SS.Util.CellRangeAddress? targetRange, Dictionary<(int Row, int Col), string>? parsedUpdates, Dictionary<(int Row, int Col), byte[]>? parsedImageUpdates)
        {
            container.Page(page =>
            {
                page.Margin(0.5f, Unit.Centimetre);
                page.Size(PageSizes.A4);
                page.PageColor(Colors.White);
                page.DefaultTextStyle(x => x.FontSize(10).FontFamily(Fonts.Arial));

                page.Content().Element(content =>
                {
                    ComposeSheet(content, sheet, evaluator, targetRange, parsedUpdates, parsedImageUpdates);
                });
            });
        }

        private void ComposeSheet(IContainer container, ISheet sheet, IFormulaEvaluator evaluator, NPOI.SS.Util.CellRangeAddress? printRange = null, Dictionary<(int Row, int Col), string>? parsedUpdates = null, Dictionary<(int Row, int Col), byte[]>? parsedImageUpdates = null)
        {
            int minRow = printRange?.FirstRow ?? sheet.FirstRowNum;
            int maxRow = printRange?.LastRow ?? sheet.LastRowNum;
            int minCol = printRange?.FirstColumn ?? 0;
            int maxCol = printRange != null ? printRange.LastColumn + 1 : 0;

            if (printRange == null)
            {
                for (int i = minRow; i <= maxRow; i++)
                {
                    var row = sheet.GetRow(i);
                    if (row != null)
                    {
                        if (row.LastCellNum > maxCol) maxCol = row.LastCellNum;
                    }
                }
            }

            var images = GetSheetImages(sheet);

            float totalRelativeWidth = 0;
            var columnWidths = new Dictionary<int, float>();
            for (int c = minCol; c < maxCol; c++)
            {
                var width = sheet.GetColumnWidth(c);
                float widthInPoints = (float)((width / 256f) * 6.0f);
                columnWidths[c] = widthInPoints;
                totalRelativeWidth += widthInPoints;
            }

            // A4 width (595.28) - Margins (2cm = 56.7pts)
            float printablePageWidth = 595.28f - 56.7f;

            container.Table(table =>
            {
                table.ColumnsDefinition(columns =>
                {
                    for (int c = minCol; c < maxCol; c++)
                    {
                        columns.RelativeColumn(columnWidths[c]);
                    }
                });

                for (int r = minRow; r <= maxRow; r++)
                {
                    var row = sheet.GetRow(r);
                    float rowHeight = row?.HeightInPoints ?? sheet.DefaultRowHeightInPoints;
                    if (rowHeight <= 0) rowHeight = 15;

                    for (int c = 0; c < maxCol; c++)
                    {
                        var cell = row?.GetCell(c);
                        if (cell == null) continue;


                        float mergedRowHeight = rowHeight;

                        if (IsMergedCell(sheet, r, c, out var heightMergedRange) && heightMergedRange != null)
                        {
                            mergedRowHeight = 0;
                            for (int mr = heightMergedRange.FirstRow; mr <= heightMergedRange.LastRow; mr++)
                            {
                                var mRow = sheet.GetRow(mr);
                                float mRowHeight = mRow?.HeightInPoints ?? sheet.DefaultRowHeightInPoints;
                                if (mRowHeight <= 0) mRowHeight = 15;
                                mergedRowHeight += mRowHeight;
                            }
                        }

                        if (_debug)
                        {
                            Console.WriteLine($"\n--- CELL {cell!.Address} ---");
                            Console.WriteLine($"Value: {GetCellValue(cell, evaluator)}");

                            // Dimensions
                            float debugCellWidth = (columnWidths[c] / totalRelativeWidth) * printablePageWidth;
                            float debugCellHeight = rowHeight;
                            Console.WriteLine($"Dimensions: W={debugCellWidth:F2}pt, H={debugCellHeight:F2}pt");

                            // Style info
                            var style = cell.CellStyle;
                            if (style != null)
                            {
                                Console.WriteLine($"Rotation: {style.Rotation}");
                                Console.WriteLine($"Alignment: H={style.Alignment}, V={style.VerticalAlignment}");

                                var font = sheet.Workbook.GetFontAt(style.FontIndex);
                                Console.WriteLine($"Font: {font.FontName}, {font.FontHeightInPoints}pt");
                                Console.WriteLine($"Style: {(font.IsBold ? "Bold " : "")}{(font.IsItalic ? "Italic " : "")}{(font.Underline != FontUnderlineType.None ? "Underline" : "")}");

                                var bgColor = ExcelToPdfExt.GetBackgroundColor(cell);
                                if (bgColor != null) Console.WriteLine($"Background: {bgColor}");
                            }
                            Console.WriteLine("------------------------");
                        }

                        if (IsMergedCell(sheet, r, c, out var mergedRange) && mergedRange != null && (mergedRange.FirstRow != r || mergedRange.FirstColumn != c))
                        {
                            continue;
                        }

                        var tableCell = table.Cell();

                        if (mergedRange != null && (mergedRange.FirstRow == r && mergedRange.FirstColumn == c))
                        {
                            int spanRow = Math.Min(mergedRange.LastRow, maxRow) - r + 1;
                            int spanCol = Math.Min(mergedRange.LastColumn + 1, maxCol) - c;

                            if (spanRow > 1) tableCell.RowSpan((uint)spanRow);
                            if (spanCol > 1) tableCell.ColumnSpan((uint)spanCol);
                        }

                        tableCell.Row((uint)(r - minRow + 1));
                        tableCell.Column((uint)(c - minCol + 1));

                        IContainer cellContainer = tableCell;
                        float effectiveHeight = cell.IsMergedCell ? mergedRowHeight : rowHeight;
                        cellContainer = cellContainer.MinHeight(effectiveHeight).MaxHeight(effectiveHeight);
                        // Pass the cell position and merged range to ApplyCellStyle
                        cellContainer = ExcelToPdfExt.ApplyCellStyle(cellContainer, cell, r, c, mergedRange);

                        // Apply alignment to the cell container (for non-layer content)
                        // cellContainer = ApplyAlignment(cellContainer, cell);

                        bool hasImage = images.TryGetValue((r, c), out var pictureData);

                        // For merged cells, also check for images within the merged region
                        if (!hasImage && mergedRange != null)
                        {
                            for (int mr = mergedRange.FirstRow; mr <= mergedRange.LastRow && !hasImage; mr++)
                            {
                                for (int mc = mergedRange.FirstColumn; mc <= mergedRange.LastColumn && !hasImage; mc++)
                                {
                                    if (images.TryGetValue((mr, mc), out pictureData))
                                    {
                                        hasImage = true;
                                    }
                                }
                            }
                        }

                        // Check for injected image
                        if (parsedImageUpdates != null && parsedImageUpdates.TryGetValue((r, c), out var injectedImage))
                        {
                            hasImage = true;
                            pictureData = injectedImage;
                        }
                        
                        // For merged cells, also check for injected images within the merged region
                        if (!hasImage && mergedRange != null && parsedImageUpdates != null)
                        {
                            for (int mr = mergedRange.FirstRow; mr <= mergedRange.LastRow && !hasImage; mr++)
                            {
                                for (int mc = mergedRange.FirstColumn; mc <= mergedRange.LastColumn && !hasImage; mc++)
                                {
                                    if (parsedImageUpdates.TryGetValue((mr, mc), out var injectedRegionImage))
                                    {
                                        hasImage = true;
                                        pictureData = injectedRegionImage;
                                    }
                                }
                            }
                        }

                        // Check for injected value
                        string cellValue;
                        if (parsedUpdates != null && parsedUpdates.TryGetValue((r, c), out var injectedValue))
                        {
                            cellValue = injectedValue;
                        }
                        else
                        {
                            cellValue = cell != null ? GetCellValue(cell, evaluator) : "";
                        }

                        // Calculate cell dimensions for RenderText
                        // Use exact width based on page layout
                        float cellWidth = (columnWidths[c] / totalRelativeWidth) * printablePageWidth;
                        float cellHeight = rowHeight;

                        if (mergedRange != null && (mergedRange.FirstRow == r && mergedRange.FirstColumn == c))
                        {
                            cellWidth = 0;
                            int lastCol = Math.Min(mergedRange.LastColumn, maxCol - 1);
                            for (int mc = mergedRange.FirstColumn; mc <= lastCol; mc++)
                            {
                                if (columnWidths.ContainsKey(mc))
                                {
                                    cellWidth += (columnWidths[mc] / totalRelativeWidth) * printablePageWidth;
                                }
                            }

                            cellHeight = 0;
                            for (int mr = mergedRange.FirstRow; mr <= mergedRange.LastRow; mr++)
                            {
                                var mRow = sheet.GetRow(mr);
                                float mRowHeight = mRow?.HeightInPoints ?? sheet.DefaultRowHeightInPoints;
                                if (mRowHeight <= 0) mRowHeight = 15;
                                cellHeight += mRowHeight;
                            }
                        }

                        bool hasText = !string.IsNullOrEmpty(cellValue);

                        if (hasImage || hasText)
                        {
                            cellContainer.Element(e =>
                            {
                                if (hasImage && hasText)
                                {
                                    e.Layers(layers =>
                                    {
                                        layers.PrimaryLayer().Element(l =>
                                        {
                                            // Apply alignment to the text layer explicitly
                                            // l = ApplyAlignment(l, cell);
                                            RenderText(l, cell, cellValue, cellWidth, cellHeight);
                                        });
                                        layers.Layer().Element(l =>
                                        {
                                            // Center images in merged cells
                                            l.AlignCenter().AlignMiddle()
                                                .MaxHeight(mergedRowHeight)
                                                .MaxWidth(cellWidth)
                                                .Image(pictureData!)
                                                .FitArea();
                                        });
                                    });
                                }
                                else if (hasImage)
                                {
                                    // Center images in merged cells
                                    e.AlignCenter().AlignMiddle()
                                        .MaxHeight(mergedRowHeight)
                                        .MaxWidth(cellWidth)
                                        .Image(pictureData!)
                                        .FitArea();
                                }
                                else
                                {
                                    RenderText(e, cell, cellValue, cellWidth, cellHeight);
                                }
                            });
                        }
                    }
                }
            });
        }

        private void RenderText(IContainer container, ICell? cell, string cellValue, float availableWidth, float availableHeight)
        {
            if (availableWidth < 1 || availableHeight < 1) return;

            var style = cell?.CellStyle;
            bool isRotated90 = false;
            float questPdfRotation = 0;

            if (style != null && style.Rotation != 0 && style.Rotation != 255)
            {
                float rotation = style.Rotation;
                if (rotation <= 90) questPdfRotation = -rotation;
                else if (rotation > 90) questPdfRotation = (rotation - 90);

                // Check if rotation is close to 90 or 270 degrees (vertical text)
                if (Math.Abs(Math.Abs(questPdfRotation) - 90) < 1)
                {
                    isRotated90 = true;
                }
            }

            if (isRotated90)
            {
                // Apply Outer Alignment (Horizontal Position in Cell) based on Excel's VerticalAlignment
                var outerContainer = container.ScaleToFit();

                if (style != null)
                {
                    // Map VerticalAlignment to Horizontal Alignment
                    outerContainer = style.VerticalAlignment switch
                    {
                        NPOI.SS.UserModel.VerticalAlignment.Top => outerContainer.AlignRight(), // Assuming Top = Right side of cell
                        NPOI.SS.UserModel.VerticalAlignment.Center => outerContainer.AlignCenter(),
                        NPOI.SS.UserModel.VerticalAlignment.Bottom => outerContainer.AlignLeft(), // Assuming Bottom = Left side of cell
                        _ => outerContainer.AlignCenter()
                    };

                    // Map HorizontalAlignment to Vertical Alignment (Outer)
                    // In Excel, Horizontal Alignment controls the position along the text axis (which is vertical in the PDF)
                    var hAlign = style.Alignment;
                    if (hAlign == NPOI.SS.UserModel.HorizontalAlignment.General)
                    {
                        // Default for text is Left (Bottom), for numbers is Right (Top)
                        // But for rotated text, usually we want Center or Bottom. Let's stick to mapping logic.
                        hAlign = NPOI.SS.UserModel.HorizontalAlignment.Justify;
                    }

                    outerContainer = hAlign switch
                    {
                        NPOI.SS.UserModel.HorizontalAlignment.Left => outerContainer.AlignBottom(),
                        NPOI.SS.UserModel.HorizontalAlignment.Center => outerContainer.AlignMiddle(),
                        NPOI.SS.UserModel.HorizontalAlignment.Right => outerContainer.AlignTop(),
                        NPOI.SS.UserModel.HorizontalAlignment.Justify => outerContainer.AlignBottom(),
                        _ => outerContainer.AlignBottom()
                    };
                }
                else
                {
                    outerContainer = outerContainer.AlignCenter().AlignMiddle();
                }

                outerContainer.Element(e =>
                {
                    // Use Unconstrained to ensure text doesn't wrap.
                    // Use Shrink (formerly MinimalBox) to ensure the container reports the actual text size (not infinity) to ScaleToFit.
                    var rotated = questPdfRotation < 0 ? e.RotateLeft() : e.RotateRight();

                    rotated.Text(text =>
                    {
                        var t = text.Span(cellValue);
                        ExcelToPdfExt.ApplyFontStyles(t, cell);
                    });
                });
            }
            else
            {
                // Standard behavior for non-90-degree rotation
                // Apply standard alignment
                container = ExcelToPdfExt.ApplyAlignment(container, cell, isText: true);

                container.ScaleToFit().Element(e =>
                {
                    if (questPdfRotation != 0)
                    {
                        e = e.Rotate(questPdfRotation);
                    }

                    e.ApplyAlignment(cell, isText: true).Text(text =>
                    {
                        var t = text.Span(cellValue);
                        ExcelToPdfExt.ApplyFontStyles(t, cell);
                    });
                });
            }
        }

        private Dictionary<(int Row, int Col), byte[]> GetSheetImages(ISheet sheet)
        {
            var images = new Dictionary<(int Row, int Col), byte[]>();
            var drawing = sheet.CreateDrawingPatriarch() as XSSFDrawing;
            if (drawing != null)
            {
                foreach (var shape in drawing.GetShapes())
                {
                    if (shape is XSSFPicture picture)
                    {
                        var anchor = picture.ClientAnchor;
                        if (anchor != null)
                        {
                            int row = anchor.Row1;
                            int col = anchor.Col1;
                            images[(row, col)] = picture.PictureData.Data;
                        }
                    }
                }
            }
            return images;
        }

        private bool IsMergedCell(ISheet sheet, int row, int col, out NPOI.SS.Util.CellRangeAddress? range)
        {
            range = null;
            for (int i = 0; i < sheet.NumMergedRegions; i++)
            {
                var region = sheet.GetMergedRegion(i);
                if (region.IsInRange(row, col))
                {
                    range = region;
                    return true;
                }
            }
            return false;
        }

        private string GetCellValue(ICell? cell, IFormulaEvaluator evaluator)
        {
            if (cell == null) return "";
            var formatter = new DataFormatter();
            var value = formatter.FormatCellValue(cell, evaluator);
            if (string.IsNullOrWhiteSpace(value) && _debug)
            {
                return cell.Address.ToString();
            }
            return value.Trim(' ', '\n');
        }
    }
}
