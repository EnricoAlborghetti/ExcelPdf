using System;
using System.Diagnostics;
using System.IO;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;
using NPOI.SS.UserModel;
using ExcelPdf;
using System.Reflection;

namespace ExcelPdf.Benchmarks
{
    public class Program
    {
        public static void Main(string[] args)
        {
            string tempFile = "benchmark.xlsx";
            int numMergedRegions = 200;
            int numLookups = 500;

            Console.WriteLine("Creating workbook...");
            using (var fs = new FileStream(tempFile, FileMode.Create, FileAccess.Write))
            {
                var workbook = new XSSFWorkbook();
                var sheet = workbook.CreateSheet("Sheet1");
                for (int i = 0; i < numMergedRegions; i++)
                {
                    sheet.AddMergedRegion(new CellRangeAddress(i * 2, i * 2 + 1, 0, 1));
                }
                workbook.Write(fs);
            }

            try
            {
                using (var helper = new ExcelHelper(tempFile))
                {
                    var sheetField = typeof(ExcelHelper).GetField("_workbook", BindingFlags.NonPublic | BindingFlags.Instance);
                    var workbook = sheetField.GetValue(helper) as IWorkbook;
                    var sheet = workbook.GetSheetAt(0);

                    var method = typeof(ExcelHelper).GetMethod("GetMergedRegion", BindingFlags.NonPublic | BindingFlags.Instance);

                    Console.WriteLine("Starting benchmark...");
                    var sw = Stopwatch.StartNew();
                    for (int i = 0; i < numLookups; i++)
                    {
                        method.Invoke(helper, new object[] { sheet, (i % numMergedRegions) * 2, 0 });
                    }
                    sw.Stop();
                    Console.WriteLine($"Result: {numLookups} lookups with {numMergedRegions} merged regions took {sw.ElapsedMilliseconds}ms");
                }
            }
            finally
            {
                if (File.Exists(tempFile)) File.Delete(tempFile);
            }
        }
    }
}
