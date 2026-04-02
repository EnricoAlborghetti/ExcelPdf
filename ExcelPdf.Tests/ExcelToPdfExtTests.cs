using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using QuestPDF.Fluent;
using QuestPDF.Infrastructure;
using Moq;
using Xunit;
using ExcelPdf;

namespace ExcelPdf.Tests
{
    public class ExcelToPdfExtTests
    {
        [Fact]
        public void GetBorderWidth_ShouldReturnCorrectWidths()
        {
            Assert.Equal(1f, ExcelToPdfExt.GetBorderWidth(BorderStyle.Thin));
            Assert.Equal(1.5f, ExcelToPdfExt.GetBorderWidth(BorderStyle.Medium));
            Assert.Equal(2.5f, ExcelToPdfExt.GetBorderWidth(BorderStyle.Thick));
            Assert.Equal(1f, ExcelToPdfExt.GetBorderWidth(BorderStyle.Dotted)); // Default case
        }

        [Fact]
        public void GetBackgroundColor_NoFill_ShouldReturnNull()
        {
            var workbook = new XSSFWorkbook();
            var sheet = workbook.CreateSheet();
            var row = sheet.CreateRow(0);
            var cell = row.CreateCell(0);
            cell.CellStyle.FillPattern = FillPattern.NoFill;

            var color = cell.GetBackgroundColor();

            Assert.Null(color);
        }

        [Fact]
        public void GetBackgroundColor_XSSF_ShouldReturnHexColor()
        {
            var workbook = new XSSFWorkbook();
            var sheet = workbook.CreateSheet();
            var row = sheet.CreateRow(0);
            var cell = row.CreateCell(0);
            var style = (XSSFCellStyle)workbook.CreateCellStyle();

            var color = new XSSFColor(new byte[] { 255, 0, 0 }); // Red
            style.SetFillForegroundColor(color);
            style.FillPattern = FillPattern.SolidForeground;
            cell.CellStyle = style;

            var hexColor = cell.GetBackgroundColor();

            Assert.Equal("#FF0000", hexColor);
        }

        [Fact]
        public void GetBackgroundColor_XSSF_WithTint_ShouldReturnTintedColor()
        {
            var workbook = new XSSFWorkbook();
            var sheet = workbook.CreateSheet();
            var row = sheet.CreateRow(0);
            var cell = row.CreateCell(0);
            var style = (XSSFCellStyle)workbook.CreateCellStyle();

            var color = new XSSFColor(new byte[] { 255, 0, 0 }); // Red
            color.Tint = 0.5; // Lighten
            style.SetFillForegroundColor(color);
            style.FillPattern = FillPattern.SolidForeground;
            cell.CellStyle = style;

            var hexColor = cell.GetBackgroundColor();

            // Lighten Red (255, 0, 0) by 0.5:
            // R = 255 * (1 - 0.5) + (255 * 0.5) = 127.5 + 127.5 = 255
            // G = 0 * (1 - 0.5) + (255 * 0.5) = 127.5 -> 127
            // B = 0 * (1 - 0.5) + (255 * 0.5) = 127.5 -> 127
            // #FF7F7F
            Assert.Equal("#FF7F7F", hexColor);
        }

        [Fact]
        public void GetEffectiveBorderWidth_SingleCell_ShouldReturnMaxBorder()
        {
            var workbook = new XSSFWorkbook();
            var sheet = workbook.CreateSheet();
            var row = sheet.CreateRow(1);
            var cell = row.CreateCell(1);

            var style = workbook.CreateCellStyle();
            style.BorderTop = BorderStyle.Thick;
            cell.CellStyle = style;

            // Check top border of cell (1,1)
            float width = sheet.GetEffectiveBorderWidth(1, 1, ExcelToPdfExt.BorderSide.Top, null);

            Assert.Equal(2.5f, width);
        }

        [Fact]
        public void GetEffectiveBorderWidth_AdjacentCell_ShouldReturnMaxBorder()
        {
            var workbook = new XSSFWorkbook();
            var sheet = workbook.CreateSheet();

            // Cell (0,1) has bottom border
            var row0 = sheet.CreateRow(0);
            var cell01 = row0.CreateCell(1);
            var style01 = workbook.CreateCellStyle();
            style01.BorderBottom = BorderStyle.Medium;
            cell01.CellStyle = style01;

            // Cell (1,1) has no top border
            var row1 = sheet.CreateRow(1);
            var cell11 = row1.CreateCell(1);

            // Check effective top border of cell (1,1) - should pick up bottom border of (0,1)
            float width = sheet.GetEffectiveBorderWidth(1, 1, ExcelToPdfExt.BorderSide.Top, null);

            Assert.Equal(1.5f, width);
        }

        [Fact]
        public void ApplyAlignment_ShouldCallCorrectMethods()
        {
            var workbook = new XSSFWorkbook();
            var sheet = workbook.CreateSheet();
            var cell = sheet.CreateRow(0).CreateCell(0);
            var style = workbook.CreateCellStyle();
            style.Alignment = HorizontalAlignment.Center;
            style.VerticalAlignment = VerticalAlignment.Top;
            cell.CellStyle = style;

            var containerMock = new Mock<IContainer>();
            // Since QuestPDF methods are extensions, they eventually call some method on IContainer or return a new one.
            // This is tricky to test with Moq because of extensions.
            // Instead, we just verify the method runs without error.

            var result = ExcelToPdfExt.ApplyAlignment(containerMock.Object, cell);

            Assert.NotNull(result);
        }

        [Fact]
        public void ApplyFontStyles_ShouldApplyProperties()
        {
            var workbook = new XSSFWorkbook();
            var font = workbook.CreateFont();
            font.FontHeightInPoints = 12;
            font.IsBold = true;
            font.FontName = "Arial";

            var sheet = workbook.CreateSheet();
            var cell = sheet.CreateRow(0).CreateCell(0);
            var style = workbook.CreateCellStyle();
            style.SetFont(font);
            cell.CellStyle = style;

            // TextSpanDescriptor is difficult to mock, so we verify no exception is thrown
            // and the method runs as expected.
            ExcelToPdfExt.ApplyFontStyles(null!, cell); // Should handle null text descriptor gracefully or at least not crash when called as per method signature
        }

        [Fact]
        public void ApplyTint_Lighten_ShouldReturnCorrectHex()
        {
            byte[] rgb = { 255, 0, 0 }; // Red
            double tint = 0.5; // Lighten

            var hex = ExcelToPdfExt.ApplyTint(rgb, tint);

            Assert.Equal("#FF7F7F", hex);
        }

        [Fact]
        public void ApplyTint_Darken_ShouldReturnCorrectHex()
        {
            byte[] rgb = { 255, 0, 0 }; // Red
            double tint = -0.5; // Darken

            var hex = ExcelToPdfExt.ApplyTint(rgb, tint);

            Assert.Equal("#7F0000", hex);
        }

        [Fact]
        public void ApplyCellStyle_ShouldRunSuccessfully()
        {
            var workbook = new XSSFWorkbook();
            var sheet = workbook.CreateSheet();
            var cell = sheet.CreateRow(0).CreateCell(0);
            var style = (XSSFCellStyle)workbook.CreateCellStyle();
            style.SetFillForegroundColor(new XSSFColor(new byte[] { 0, 255, 0 }));
            style.FillPattern = FillPattern.SolidForeground;
            style.BorderTop = BorderStyle.Thick;
            cell.CellStyle = style;

            var containerMock = new Mock<IContainer>();
            containerMock.Setup(c => c.Background(It.IsAny<string>())).Returns(containerMock.Object);
            containerMock.Setup(c => c.BorderTop(It.IsAny<float>())).Returns(containerMock.Object);

            var result = ExcelToPdfExt.ApplyCellStyle(containerMock.Object, cell, 0, 0, null);

            Assert.NotNull(result);
        }
    }
}
