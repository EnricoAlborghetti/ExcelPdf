using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using QuestPDF.Fluent;
using QuestPDF.Infrastructure;

namespace ExcelPdf
{
    /// <summary>
    /// Extension methods and helper functions for styling and border calculation during PDF conversion.
    /// </summary>
    public static class ExcelToPdfExt
    {
        /// <summary>
        /// Applies background color and borders to a QuestPDF container based on the cell's style.
        /// </summary>
        /// <param name="container">The QuestPDF container.</param>
        /// <param name="cell">The NPOI cell source.</param>
        /// <param name="row">The current row index.</param>
        /// <param name="col">The current column index.</param>
        /// <param name="mergedRange">The merged range this cell belongs to, if any.</param>
        /// <returns>The modified container.</returns>
        public static IContainer ApplyCellStyle(IContainer container, ICell? cell, int row, int col, NPOI.SS.Util.CellRangeAddress? mergedRange)
        {
            if (cell == null) return container;
            var style = cell.CellStyle;

            // Background color
            string? hexColor = GetBackgroundColor(cell);
            if (!string.IsNullOrEmpty(hexColor))
            {
                container = container.Background(hexColor);
            }

            // Borders
            // We need to check the effective border for each side, considering adjacent cells.
            var sheet = cell.Sheet;

            float top = GetEffectiveBorderWidth(sheet, row, col, BorderSide.Top, mergedRange);
            float bottom = GetEffectiveBorderWidth(sheet, row, col, BorderSide.Bottom, mergedRange);
            float left = GetEffectiveBorderWidth(sheet, row, col, BorderSide.Left, mergedRange);
            float right = GetEffectiveBorderWidth(sheet, row, col, BorderSide.Right, mergedRange);

            if (top > 0) container = container.BorderTop(top);
            if (bottom > 0) container = container.BorderBottom(bottom);
            if (left > 0) container = container.BorderLeft(left);
            if (right > 0) container = container.BorderRight(right);

            return container;
        }

        public enum BorderSide { Top, Bottom, Left, Right }

        /// <summary>
        /// Calculates the effective border width for a specific side of a cell or merged range,
        /// considering the borders of adjacent cells.
        /// </summary>
        /// <param name="sheet">The Excel sheet.</param>
        /// <param name="row">The row index.</param>
        /// <param name="col">The column index.</param>
        /// <param name="side">The side of the border to check.</param>
        /// <param name="mergedRange">The merged range, if applicable.</param>
        /// <returns>The maximum border width found for that edge.</returns>
        public static float GetEffectiveBorderWidth(this ISheet sheet, int row, int col, BorderSide side, NPOI.SS.Util.CellRangeAddress? mergedRange)
        {
            // Determine the range of cells to check for this side.
            // If it's a merged cell, we check the entire edge of the merged region.
            // If not, it's just the single cell (row, col).

            int startRow = row;
            int endRow = row;
            int startCol = col;
            int endCol = col;

            if (mergedRange != null)
            {
                startRow = mergedRange.FirstRow;
                endRow = mergedRange.LastRow;
                startCol = mergedRange.FirstColumn;
                endCol = mergedRange.LastColumn;
            }

            float maxBorderWidth = 0;

            // Helper to get border width from a specific cell and side
            float GetWidth(int r, int c, BorderSide s)
            {
                var cell = sheet.GetRow(r)?.GetCell(c);
                if (cell == null) return 0;
                var style = cell.CellStyle;
                if (style == null) return 0;

                BorderStyle bs = BorderStyle.None;
                switch (s)
                {
                    case BorderSide.Top: bs = style.BorderTop; break;
                    case BorderSide.Bottom: bs = style.BorderBottom; break;
                    case BorderSide.Left: bs = style.BorderLeft; break;
                    case BorderSide.Right: bs = style.BorderRight; break;
                }
                return bs == BorderStyle.None ? 0 : GetBorderWidth(bs);
            }

            switch (side)
            {
                case BorderSide.Top:
                    // Check Top border of cells in the top row of the range
                    // AND Bottom border of cells in the row above
                    for (int c = startCol; c <= endCol; c++)
                    {
                        maxBorderWidth = Math.Max(maxBorderWidth, GetWidth(startRow, c, BorderSide.Top));
                        if (startRow > 0)
                            maxBorderWidth = Math.Max(maxBorderWidth, GetWidth(startRow - 1, c, BorderSide.Bottom));
                    }
                    break;

                case BorderSide.Bottom:
                    // Check Bottom border of cells in the bottom row of the range
                    // AND Top border of cells in the row below
                    for (int c = startCol; c <= endCol; c++)
                    {
                        maxBorderWidth = Math.Max(maxBorderWidth, GetWidth(endRow, c, BorderSide.Bottom));
                        // Note: We don't strictly check bounds for row + 1 as GetRow will just return null if out of bounds
                        maxBorderWidth = Math.Max(maxBorderWidth, GetWidth(endRow + 1, c, BorderSide.Top));
                    }
                    break;

                case BorderSide.Left:
                    // Check Left border of cells in the left column of the range
                    // AND Right border of cells in the column to the left
                    for (int r = startRow; r <= endRow; r++)
                    {
                        maxBorderWidth = Math.Max(maxBorderWidth, GetWidth(r, startCol, BorderSide.Left));
                        if (startCol > 0)
                            maxBorderWidth = Math.Max(maxBorderWidth, GetWidth(r, startCol - 1, BorderSide.Right));
                    }
                    break;

                case BorderSide.Right:
                    // Check Right border of cells in the right column of the range
                    // AND Left border of cells in the column to the right
                    for (int r = startRow; r <= endRow; r++)
                    {
                        maxBorderWidth = Math.Max(maxBorderWidth, GetWidth(r, endCol, BorderSide.Right));
                        maxBorderWidth = Math.Max(maxBorderWidth, GetWidth(r, endCol + 1, BorderSide.Left));
                    }
                    break;
            }

            return maxBorderWidth;
        }

        /// <summary>
        /// Applies horizontal and vertical alignment to a container.
        /// </summary>
        /// <param name="container">The QuestPDF container.</param>
        /// <param name="cell">The NPOI cell source.</param>
        /// <param name="isText">If true, applies specific logic for text indentation.</param>
        /// <returns>The aligned container.</returns>
        public static IContainer ApplyAlignment(this IContainer container, ICell? cell, bool isText = false)
        {
            if (cell == null) return container.AlignLeft().AlignMiddle();

            var style = cell.CellStyle;

            // Vertical Alignment
            container = style.VerticalAlignment switch
            {
                NPOI.SS.UserModel.VerticalAlignment.Top => container.AlignTop(),
                NPOI.SS.UserModel.VerticalAlignment.Center => container.AlignMiddle(),
                NPOI.SS.UserModel.VerticalAlignment.Bottom => container.AlignBottom(),
                NPOI.SS.UserModel.VerticalAlignment.Justify => container.AlignMiddle(),
                _ => container.AlignMiddle()
            };

            // Horizontal Alignment
            var alignment = style.Alignment;
            if (alignment == NPOI.SS.UserModel.HorizontalAlignment.General)
            {
                // Determine default based on cell type
                bool isNumeric = cell.CellType == CellType.Numeric ||
                                (cell.CellType == CellType.Formula && cell.CachedFormulaResultType == CellType.Numeric);

                if (isNumeric) alignment = NPOI.SS.UserModel.HorizontalAlignment.Right;
                else alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
            }

            if (isText && alignment == NPOI.SS.UserModel.HorizontalAlignment.Left && style.Indention == 0)
            {
                container = container.PaddingLeft(2);
            }

            container = alignment switch
            {
                NPOI.SS.UserModel.HorizontalAlignment.Left => container.AlignLeft(),
                NPOI.SS.UserModel.HorizontalAlignment.Center => container.AlignCenter(),
                NPOI.SS.UserModel.HorizontalAlignment.Right => container.AlignRight(),
                NPOI.SS.UserModel.HorizontalAlignment.Justify => container.AlignLeft(),
                _ => container.AlignLeft()
            };

            return container;
        }

        /// <summary>
        /// Applies font styles (bold, italic, underline, size, family) to a text span.
        /// </summary>
        /// <param name="text">The QuestPDF text span descriptor.</param>
        /// <param name="cell">The NPOI cell source.</param>
        public static void ApplyFontStyles(TextSpanDescriptor text, ICell? cell)
        {
            if (cell == null) return;
            var font = cell.Sheet.Workbook.GetFontAt(cell.CellStyle.FontIndex);

            text.FontSize((float)font.FontHeightInPoints);

            if (font.IsBold) text.Bold();
            if (font.IsItalic) text.Italic();
            if (font.Underline != FontUnderlineType.None) text.Underline();

            if (!string.IsNullOrEmpty(font.FontName))
            {
                text.FontFamily(font.FontName);
            }
        }

        /// <summary>
        /// Retrieves the background color of a cell as a hex string, handling tinting.
        /// </summary>
        /// <param name="cell">The NPOI cell.</param>
        /// <returns>A hex color string (e.g., "#FF0000") or null if no fill.</returns>
        public static string? GetBackgroundColor(this ICell cell)
        {
            var style = cell.CellStyle;
            if (style.FillPattern == FillPattern.NoFill) return null;

            if (style is XSSFCellStyle xssfStyle)
            {
                var color = xssfStyle.FillForegroundXSSFColor;
                if (color == null) return null;

                    byte[] rgb = color.RGB;
                    if (rgb != null && rgb.Length == 3)
                    {
                        if (color.Tint != 0)
                        {
                            return ApplyTint(rgb, color.Tint);
                        }
                        return $"#{rgb[0]:X2}{rgb[1]:X2}{rgb[2]:X2}";
                    }
            }
            else if (style is HSSFCellStyle hssfStyle)
            {
                var workbook = cell.Sheet.Workbook as HSSFWorkbook;
                var palette = workbook?.GetCustomPalette();
                var color = palette?.GetColor(hssfStyle.FillForegroundColor);

                if (color != null)
                {
                    byte[] rgb = color.RGB;
                    if (rgb != null && rgb.Length == 3)
                    {
                        return $"#{rgb[0]:X2}{rgb[1]:X2}{rgb[2]:X2}";
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// Converts NPOI BorderStyle to a width in points.
        /// </summary>
        /// <param name="style">The NPOI border style.</param>
        /// <returns>The width in points.</returns>
        public static float GetBorderWidth(BorderStyle style)
        {
            return style switch
            {
                BorderStyle.Thin => 1f,
                BorderStyle.Medium => 1.5f,
                BorderStyle.Thick => 2.5f,
                _ => 1f
            };
        }

        private static string ApplyTint(byte[] rgb, double tint)
        {
            double r = rgb[0];
            double g = rgb[1];
            double b = rgb[2];

            if (tint < 0)
            {
                // Darken
                // New = Old * (1 + tint)
                // Since tint is negative, (1 + tint) is < 1
                r = r * (1.0 + tint);
                g = g * (1.0 + tint);
                b = b * (1.0 + tint);
            }
            else
            {
                // Lighten
                // New = Old * (1 - tint) + (255 * tint)
                r = r * (1.0 - tint) + (255 * tint);
                g = g * (1.0 - tint) + (255 * tint);
                b = b * (1.0 - tint) + (255 * tint);
            }

            // Clamp and convert back to byte
            byte newR = (byte)Math.Clamp(r, 0, 255);
            byte newG = (byte)Math.Clamp(g, 0, 255);
            byte newB = (byte)Math.Clamp(b, 0, 255);

            return $"#{newR:X2}{newG:X2}{newB:X2}";
        }

    }
}
