using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Sparkline;
using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Xml;

namespace EPPlusTest
{
    [TestClass]
    public class SparklineTests
    {
        [TestMethod]
        public void ReadSparklinesFromWorkbook()
        {
            string workbooksDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\workbooks");
            using (var package = new ExcelPackage(new FileInfo(workbooksDir + @"\Sparkline Demos.xlsx")))
            {
                var sparklineGroups = package.Workbook.Worksheets.First().SparklineGroups;
                Assert.IsNotNull(sparklineGroups);
                Assert.IsNotNull(sparklineGroups.SparklineGroups);
                Assert.AreEqual(3, sparklineGroups.SparklineGroups.Count);
                var group1 = sparklineGroups.SparklineGroups[0];
                var group2 = sparklineGroups.SparklineGroups[1];
                var group3 = sparklineGroups.SparklineGroups[2];
                Assert.AreEqual(Color.FromArgb(unchecked((int)0xFF376092)), group1.ColorSeries);
                Assert.AreEqual(Color.FromArgb(unchecked((int)0xFF376092)), group2.ColorSeries);
                Assert.AreEqual(Color.FromArgb(unchecked((int)0xFF323232)), group3.ColorSeries);
                Assert.AreEqual(SparklineType.Column, group1.Type);
                Assert.AreEqual(SparklineType.Line, group2.Type);
                Assert.AreEqual(SparklineType.Stacked, group3.Type);
                Assert.AreEqual("Sheet1!D6:F6", group1.Sparklines[0].Formula.Address);
                Assert.AreEqual("Sheet1!D7:F7", group2.Sparklines[0].Formula.Address);
                Assert.AreEqual("Sheet1!D8:F8", group3.Sparklines[0].Formula.Address);
                Assert.AreEqual("G6", group1.Sparklines[0].HostCell.Address);
                Assert.AreEqual("G7", group2.Sparklines[0].HostCell.Address);
                Assert.AreEqual("G8", group3.Sparklines[0].HostCell.Address);
            }
        }

        [TestMethod]
        public void ReadThemedSparklinesFromWorkbook()
        {
            var newFile = new FileInfo(Path.GetTempFileName());
            if (newFile.Exists)
                newFile.Delete();
            string workbooksDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\workbooks");
            using (var package = new ExcelPackage(new FileInfo(workbooksDir + @"\ThemedSparkline.xlsx")))
            {
                var sparklineGroups = package.Workbook.Worksheets.First().SparklineGroups;
                Assert.IsNotNull(sparklineGroups);
                Assert.IsNotNull(sparklineGroups.SparklineGroups);
                Assert.AreEqual(1, sparklineGroups.SparklineGroups.Count);
                var group1 = sparklineGroups.SparklineGroups[0];
                Assert.AreEqual(Color.Empty, group1.ColorSeries);
                var colorSeriesNode = group1.TopNode.SelectSingleNode("x14:colorSeries", group1.NameSpaceManager);
                Assert.IsNotNull(colorSeriesNode);
                Assert.AreEqual("3", colorSeriesNode.Attributes["theme"].InnerText);
                Assert.AreEqual(SparklineType.Column, group1.Type);
                package.SaveAs(newFile);
            }
            using (var package = new ExcelPackage(newFile))
            {
                var sparklineGroups = package.Workbook.Worksheets.First().SparklineGroups;
                Assert.IsNotNull(sparklineGroups);
                Assert.IsNotNull(sparklineGroups.SparklineGroups);
                Assert.AreEqual(1, sparklineGroups.SparklineGroups.Count);
                var group1 = sparklineGroups.SparklineGroups[0];
                Assert.AreEqual(Color.Empty, group1.ColorSeries);
                var colorSeriesNode = group1.TopNode.SelectSingleNode("x14:colorSeries", group1.NameSpaceManager);
                Assert.IsNotNull(colorSeriesNode);
                Assert.AreEqual("3", colorSeriesNode.Attributes["theme"].InnerText);
                Assert.AreEqual(SparklineType.Column, group1.Type);
            }
        }

        [TestMethod]
        public void InsertCellsUpdatesSparklineRanges()
        {
            string workbooksDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\workbooks");
            using (var package = new ExcelPackage(new FileInfo(workbooksDir + @"\Sparkline Demos.xlsx")))
            {
                var sheet = package.Workbook.Worksheets.First();
                var sparklineGroups = sheet.SparklineGroups;
                var group1 = sparklineGroups.SparklineGroups[0];
                var newLine = new ExcelSparkline(group1, group1.NameSpaceManager) { Formula = new ExcelAddress("Sheet1!D9:F9"), HostCell = new ExcelAddress("G9") };
                group1.Sparklines.Add(newLine);

                var copied = package.Workbook.Worksheets.Add("Copied", sheet);
                Assert.IsNotNull(sparklineGroups);
                Assert.IsNotNull(sparklineGroups.SparklineGroups);
                Assert.AreEqual(3, sparklineGroups.SparklineGroups.Count);
                var group2 = sparklineGroups.SparklineGroups[1];
                var group3 = sparklineGroups.SparklineGroups[2];
                Assert.AreEqual(Color.FromArgb(unchecked((int)0xFF376092)), group1.ColorSeries);
                Assert.AreEqual(Color.FromArgb(unchecked((int)0xFF376092)), group2.ColorSeries);
                Assert.AreEqual(Color.FromArgb(unchecked((int)0xFF323232)), group3.ColorSeries);
                Assert.AreEqual(SparklineType.Column, group1.Type);
                Assert.AreEqual(SparklineType.Line, group2.Type);
                Assert.AreEqual(SparklineType.Stacked, group3.Type);
                Assert.AreEqual("Sheet1!D6:F6", group1.Sparklines[0].Formula.Address);
                Assert.AreEqual("Sheet1!D7:F7", group2.Sparklines[0].Formula.Address);
                Assert.AreEqual("Sheet1!D8:F8", group3.Sparklines[0].Formula.Address);
                Assert.AreEqual("G6", group1.Sparklines[0].HostCell.Address);
                Assert.AreEqual("G7", group2.Sparklines[0].HostCell.Address);
                Assert.AreEqual("G8", group3.Sparklines[0].HostCell.Address);

                sheet.InsertRow(2, 3);
                sheet.InsertColumn(5, 3);
                copied.InsertRow(2, 4);
                copied.InsertColumn(5, 4);
                copied.Name = "Alejandro";
                Assert.AreEqual("'Sheet1'!D9:I9", group1.Sparklines[0].Formula.Address);
                Assert.AreEqual("'Sheet1'!D12:I12", group1.Sparklines[1].Formula.Address);
                Assert.AreEqual("'Sheet1'!D10:I10", group2.Sparklines[0].Formula.Address);
                Assert.AreEqual("'Sheet1'!D11:I11", group3.Sparklines[0].Formula.Address);
                Assert.AreEqual("J9", group1.Sparklines[0].HostCell.Address);
                Assert.AreEqual("J12", group1.Sparklines[1].HostCell.Address);
                Assert.AreEqual("J10", group2.Sparklines[0].HostCell.Address);
                Assert.AreEqual("J11", group3.Sparklines[0].HostCell.Address);
                var copiedGroup1 = copied.SparklineGroups.SparklineGroups[0];
                var copiedGroup2 = copied.SparklineGroups.SparklineGroups[1];
                var copiedGroup3 = copied.SparklineGroups.SparklineGroups[2];
                Assert.AreEqual("'Alejandro'!D10:J10", copiedGroup1.Sparklines[0].Formula.Address);
                Assert.AreEqual("'Alejandro'!D13:J13", copiedGroup1.Sparklines[1].Formula.Address);
                Assert.AreEqual("'Alejandro'!D11:J11", copiedGroup2.Sparklines[0].Formula.Address);
                Assert.AreEqual("'Alejandro'!D12:J12", copiedGroup3.Sparklines[0].Formula.Address);
                Assert.AreEqual("K10", copiedGroup1.Sparklines[0].HostCell.Address);
                Assert.AreEqual("K13", copiedGroup1.Sparklines[1].HostCell.Address);
                Assert.AreEqual("K11", copiedGroup2.Sparklines[0].HostCell.Address);
                Assert.AreEqual("K12", copiedGroup3.Sparklines[0].HostCell.Address);
            }
        }

        [TestMethod]
        public void CopyCellsAddsNewSparklines()
        {
            var newFile = new FileInfo(Path.GetTempFileName());
            if (newFile.Exists)
                newFile.Delete();
            try
            {
                string workbooksDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\workbooks");
                using (var package = new ExcelPackage(new FileInfo(workbooksDir + @"\Sparkline Demos.xlsx")))
                {
                    var sheet1 = package.Workbook.Worksheets.First();
                    var sheet2 = package.Workbook.Worksheets.Add("Sheet 2", sheet1);
                    var sheet3 = package.Workbook.Worksheets.Add("Sheet 3", sheet1);
                    sheet2.Name = "Sheet2";
                    sheet3.Name = "Sheet3";
                    foreach (var sheet in package.Workbook.Worksheets)
                    {
                        var sparklineGroups = sheet.SparklineGroups;
                        var group1 = sparklineGroups.SparklineGroups[0];
                        Assert.IsNotNull(sparklineGroups);
                        Assert.IsNotNull(sparklineGroups.SparklineGroups);
                        Assert.AreEqual(3, sparklineGroups.SparklineGroups.Count);
                        var group2 = sparklineGroups.SparklineGroups[1];
                        var group3 = sparklineGroups.SparklineGroups[2];
                        Assert.AreEqual(1, group1.Sparklines.Count);
                        Assert.AreEqual(Color.FromArgb(unchecked((int)0xFF376092)), group1.ColorSeries);
                        Assert.AreEqual(Color.FromArgb(unchecked((int)0xFF376092)), group2.ColorSeries);
                        Assert.AreEqual(Color.FromArgb(unchecked((int)0xFF323232)), group3.ColorSeries);
                        Assert.AreEqual(SparklineType.Column, group1.Type);
                        Assert.AreEqual(SparklineType.Line, group2.Type);
                        Assert.AreEqual(SparklineType.Stacked, group3.Type);
                        var sheetName = sheet.Name == "Sheet1" ? sheet.Name : $"'{sheet.Name}'";
                        Assert.AreEqual($"{sheetName}!D6:F6", group1.Sparklines[0].Formula.Address);
                        Assert.AreEqual($"{sheetName}!D7:F7", group2.Sparklines[0].Formula.Address);
                        Assert.AreEqual($"{sheetName}!D8:F8", group3.Sparklines[0].Formula.Address);
                        Assert.AreEqual($"G6", group1.Sparklines[0].HostCell.Address);
                        Assert.AreEqual($"G7", group2.Sparklines[0].HostCell.Address);
                        Assert.AreEqual($"G8", group3.Sparklines[0].HostCell.Address);

                        sheet.Cells["6:6"].Copy(sheet.Cells["9:9"]);
                        Assert.AreEqual(2, group1.Sparklines.Count);
                        var newLine = group1.Sparklines[1];
                        Assert.AreEqual($"'{sheet.Name}'!D9:F9", newLine.Formula.Address);
                        Assert.AreEqual("G9", newLine.HostCell.Address);

                        Assert.AreEqual($"{sheetName}!D6:F6", group1.Sparklines[0].Formula.Address);
                        Assert.AreEqual("G6", group1.Sparklines[0].HostCell.Address);
                    }
                    package.SaveAs(newFile);
                }
                using (var package = new ExcelPackage(newFile))
                {
                    foreach (var sheet in package.Workbook.Worksheets)
                    {
                        var sparklineGroups = sheet.SparklineGroups;
                        var group1 = sparklineGroups.SparklineGroups[0];
                        Assert.IsNotNull(sparklineGroups);
                        Assert.IsNotNull(sparklineGroups.SparklineGroups);
                        Assert.AreEqual(3, sparklineGroups.SparklineGroups.Count);
                        var group2 = sparklineGroups.SparklineGroups[1];
                        var group3 = sparklineGroups.SparklineGroups[2];
                        Assert.AreEqual(Color.FromArgb(unchecked((int)0xFF376092)), group1.ColorSeries);
                        Assert.AreEqual(Color.FromArgb(unchecked((int)0xFF376092)), group2.ColorSeries);
                        Assert.AreEqual(Color.FromArgb(unchecked((int)0xFF323232)), group3.ColorSeries);
                        Assert.AreEqual(SparklineType.Column, group1.Type);
                        Assert.AreEqual(SparklineType.Line, group2.Type);
                        Assert.AreEqual(SparklineType.Stacked, group3.Type);
                        var sheetName = sheet.Name == "Sheet1" ? sheet.Name : $"'{sheet.Name}'";
                        Assert.AreEqual($"{sheetName}!D6:F6", group1.Sparklines[0].Formula.Address);
                        Assert.AreEqual($"{sheetName}!D7:F7", group2.Sparklines[0].Formula.Address);
                        Assert.AreEqual($"{sheetName}!D8:F8", group3.Sparklines[0].Formula.Address);
                        Assert.AreEqual("G6", group1.Sparklines[0].HostCell.Address);
                        Assert.AreEqual("G7", group2.Sparklines[0].HostCell.Address);
                        Assert.AreEqual("G8", group3.Sparklines[0].HostCell.Address);
                        Assert.AreEqual(2, group1.Sparklines.Count);
                        var newLine = group1.Sparklines[1];
                        Assert.AreEqual($"'{sheet.Name}'!D9:F9", newLine.Formula.Address);
                        Assert.AreEqual("G9", newLine.HostCell.Address);

                        Assert.AreEqual($"{sheetName}!D6:F6", group1.Sparklines[0].Formula.Address);
                        Assert.AreEqual("G6", group1.Sparklines[0].HostCell.Address);
                    }
                }
            }
            finally
            {
                if (newFile.Exists)
                    newFile.Delete();
            }
        }

        [TestMethod]
        public void UpdateSparklinesAndSaveWorkbook()
        {
            var newFile = new FileInfo(Path.GetTempFileName());
            if (newFile.Exists)
                newFile.Delete();
            string workbooksDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\workbooks");
            using (var package = new ExcelPackage(new FileInfo(workbooksDir + @"\Sparkline Demos.xlsx")))
            {
                var sparklineGroups = package.Workbook.Worksheets.First().SparklineGroups;
                Assert.IsNotNull(sparklineGroups);
                Assert.IsNotNull(sparklineGroups.SparklineGroups);
                Assert.AreEqual(3, sparklineGroups.SparklineGroups.Count);
                var group1 = sparklineGroups.SparklineGroups[0];
                var group2 = sparklineGroups.SparklineGroups[1];
                var group3 = sparklineGroups.SparklineGroups[2];
                Assert.AreEqual(Color.FromArgb(unchecked((int)0xFF376092)), group1.ColorSeries);
                Assert.AreEqual(Color.FromArgb(unchecked((int)0xFF376092)), group2.ColorSeries);
                Assert.AreEqual(Color.FromArgb(unchecked((int)0xFF323232)), group3.ColorSeries);
                Assert.AreEqual(SparklineType.Column, group1.Type);
                Assert.AreEqual(SparklineType.Line, group2.Type);
                Assert.AreEqual(SparklineType.Stacked, group3.Type);
                Assert.AreEqual("Sheet1!D6:F6", group1.Sparklines[0].Formula.Address);
                Assert.AreEqual("Sheet1!D7:F7", group2.Sparklines[0].Formula.Address);
                Assert.AreEqual("Sheet1!D8:F8", group3.Sparklines[0].Formula.Address);
                Assert.AreEqual("G6", group1.Sparklines[0].HostCell.Address);
                Assert.AreEqual("G7", group2.Sparklines[0].HostCell.Address);
                Assert.AreEqual("G8", group3.Sparklines[0].HostCell.Address);
                var sparkline = new ExcelSparkline(group1, group1.NameSpaceManager) { Formula = new ExcelAddress("A1:B2"), HostCell = new ExcelAddress("G9") };
                group1.Sparklines.Add(sparkline);
                group1.Type = SparklineType.Stacked;
                package.SaveAs(newFile);
            }
            try
            {
                using (var package = new ExcelPackage(newFile))
                {
                    var sparklineGroups = package.Workbook.Worksheets.First().SparklineGroups;
                    Assert.IsNotNull(sparklineGroups);
                    Assert.IsNotNull(sparklineGroups.SparklineGroups);
                    Assert.AreEqual(3, sparklineGroups.SparklineGroups.Count);
                    var group1 = sparklineGroups.SparklineGroups[0];
                    var group2 = sparklineGroups.SparklineGroups[1];
                    var group3 = sparklineGroups.SparklineGroups[2];
                    Assert.AreEqual(Color.FromArgb(unchecked((int)0xFF376092)), group1.ColorSeries);
                    Assert.AreEqual(Color.FromArgb(unchecked((int)0xFF376092)), group2.ColorSeries);
                    Assert.AreEqual(Color.FromArgb(unchecked((int)0xFF323232)), group3.ColorSeries);
                    Assert.AreEqual(SparklineType.Stacked, group1.Type);
                    Assert.AreEqual(SparklineType.Line, group2.Type);
                    Assert.AreEqual(SparklineType.Stacked, group3.Type);
                    Assert.AreEqual("Sheet1!D6:F6", group1.Sparklines[0].Formula.Address);
                    Assert.AreEqual("A1:B2", group1.Sparklines[1].Formula.Address);
                    Assert.AreEqual("Sheet1!D7:F7", group2.Sparklines[0].Formula.Address);
                    Assert.AreEqual("Sheet1!D8:F8", group3.Sparklines[0].Formula.Address);
                    Assert.AreEqual("G6", group1.Sparklines[0].HostCell.Address);
                    Assert.AreEqual("G9", group1.Sparklines[1].HostCell.Address);
                    Assert.AreEqual("G7", group2.Sparklines[0].HostCell.Address);
                    Assert.AreEqual("G8", group3.Sparklines[0].HostCell.Address);
                }
            }
            finally
            {
                if (newFile.Exists)
                    newFile.Delete();
            }
        }

        // Saving sparklines from scratch is currently not supported.
        [ExpectedException(typeof(NotImplementedException))]
        [TestMethod]
        public void SaveSparklines()
        {
            FileInfo newFile = new FileInfo(Path.GetTempFileName());
            if (newFile.Exists)
                newFile.Delete();
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");
                var group = new ExcelSparklineGroup(sheet, sheet.NameSpaceManager);
                var sparkline = new ExcelSparkline(group, sheet.NameSpaceManager) { Formula = new ExcelAddress("G1:G20"), HostCell = new ExcelAddress("B2") };
                group.Sparklines.Add(sparkline);
                sheet.SparklineGroups.SparklineGroups.Add(group);

                package.SaveAs(newFile);
            }
            using (var package = new ExcelPackage(newFile))
            {
                var sparklineGroup = package.Workbook.Worksheets.First().SparklineGroups.SparklineGroups.First();
                var sparkline = sparklineGroup.Sparklines.First();
                Assert.AreEqual("B2", sparkline.HostCell.Address);
                Assert.AreEqual("G1:G20", sparkline.Formula.Address);
            }
        }

        [TestMethod]
        public void UpdateSparklinesWithDefaultAttributesAndSaveWorkbook()
        {
            var newFile = new FileInfo(Path.GetTempFileName());
            if (newFile.Exists)
                newFile.Delete();
            string workbooksDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\workbooks");
            using (var package = new ExcelPackage(new FileInfo(workbooksDir + @"\Sparkline Demos.xlsx")))
            {
                var sparklineGroups = package.Workbook.Worksheets.First().SparklineGroups;
                Assert.IsNotNull(sparklineGroups);
                Assert.IsNotNull(sparklineGroups.SparklineGroups);
                Assert.AreEqual(3, sparklineGroups.SparklineGroups.Count);
                var group1 = sparklineGroups.SparklineGroups[0];
                var group2 = sparklineGroups.SparklineGroups[1];
                var group3 = sparklineGroups.SparklineGroups[2];
                Assert.AreEqual(Color.FromArgb(unchecked((int)0xFF376092)), group1.ColorSeries);
                Assert.AreEqual(Color.FromArgb(unchecked((int)0xFF376092)), group2.ColorSeries);
                Assert.AreEqual(Color.FromArgb(unchecked((int)0xFF323232)), group3.ColorSeries);
                Assert.AreEqual(SparklineType.Column, group1.Type);
                Assert.AreEqual(SparklineType.Line, group2.Type);
                Assert.AreEqual(SparklineType.Stacked, group3.Type);
                Assert.AreEqual("Sheet1!D6:F6", group1.Sparklines[0].Formula.Address);
                Assert.AreEqual("Sheet1!D7:F7", group2.Sparklines[0].Formula.Address);
                Assert.AreEqual("Sheet1!D8:F8", group3.Sparklines[0].Formula.Address);
                Assert.AreEqual("G6", group1.Sparklines[0].HostCell.Address);
                Assert.AreEqual("G7", group2.Sparklines[0].HostCell.Address);
                Assert.AreEqual("G8", group3.Sparklines[0].HostCell.Address);
                Assert.AreEqual(DispBlanksAs.Gap, group1.DisplayEmptyCellsAs);
                var sparkline = new ExcelSparkline(group1, group1.NameSpaceManager) { Formula = new ExcelAddress("A1:B2"), HostCell = new ExcelAddress("G9") };
                group1.Sparklines.Add(sparkline);
                group1.Type = SparklineType.Stacked;
                group1.ManualMax = 1000.23;
                group1.ManualMin = 15.5;
                group1.LineWeight = 1.5;
                group1.DateAxis = true;
                group1.DisplayEmptyCellsAs = DispBlanksAs.Span;
                group1.Markers = true;
                group1.High = true;
                group1.Low = true;
                group1.First = true;
                group1.Last = true;
                group1.Negative = true;
                group1.DisplayXAxis = true;
                group1.DisplayHidden = true;
                group1.MinAxisType = SparklineAxisMinMax.Custom;
                group1.MaxAxisType = SparklineAxisMinMax.Group;
                group1.RightToLeft = true;
                group1.ColorSeries = Color.FromArgb(unchecked((int)0xAA375052));
                group1.ColorNegative = Color.FromArgb(unchecked((int)0xAB375052));
                group1.ColorAxis = Color.FromArgb(unchecked((int)0xAC375052));
                group1.ColorMarkers = Color.FromArgb(unchecked((int)0xAD375052));
                group1.ColorFirst = Color.FromArgb(unchecked((int)0xAE375052));
                group1.ColorLast = Color.FromArgb(unchecked((int)0xAF375052));
                group1.ColorHigh = Color.FromArgb(unchecked((int)0xBA375052));
                group1.ColorLow = Color.FromArgb(unchecked((int)0xBB375052));


                package.SaveAs(newFile);
            }
            try
            {
                newFile.Refresh();
                long allPropertiesDefined = newFile.Length;
                using (var package = new ExcelPackage(newFile))
                {
                    var sheet1 = package.Workbook.Worksheets.First();
                    var groups = sheet1.SparklineGroups;
                    var group1 = groups.SparklineGroups[0];
                    this.ValidateGroup(group1);
                    var sheet2 = package.Workbook.Worksheets.Add("Copied", package.Workbook.Worksheets.First());
                    var copiedGroup1 = sheet2.SparklineGroups.SparklineGroups[0];
                    var copiedGroup2 = sheet2.SparklineGroups.SparklineGroups[1];
                    var copiedGroup3 = sheet2.SparklineGroups.SparklineGroups[2];
                    Assert.AreEqual("A1:B2", copiedGroup1.Sparklines[1].Formula.Address);
                    Assert.AreEqual("G9", copiedGroup1.Sparklines[1].HostCell.Address);
                    Assert.AreEqual("'Copied'!D7:F7", copiedGroup2.Sparklines[0].Formula.Address);
                    Assert.AreEqual("'Copied'!D8:F8", copiedGroup3.Sparklines[0].Formula.Address);

                    sheet1.InsertRow(2, 2);
                    // Ensure that unchanged Group2 and Group3 properties remain the same.
                    var group2 = groups.SparklineGroups[1];
                    var group3 = groups.SparklineGroups[2];
                    Assert.AreEqual(Color.FromArgb(unchecked((int)0xFF376092)), group2.ColorSeries);
                    Assert.AreEqual(Color.FromArgb(unchecked((int)0xFF323232)), group3.ColorSeries);
                    Assert.AreEqual(SparklineType.Line, group2.Type);
                    Assert.AreEqual(SparklineType.Stacked, group3.Type);
                    Assert.AreEqual("'Sheet1'!D9:F9", group2.Sparklines[0].Formula.Address);
                    Assert.AreEqual("'Sheet1'!D10:F10", group3.Sparklines[0].Formula.Address);
                    Assert.AreEqual("G9", group2.Sparklines[0].HostCell.Address);
                    Assert.AreEqual("G10", group3.Sparklines[0].HostCell.Address);

                    // Reset all fields on group1 to their default values.
                    group1.Type = SparklineType.Line;
                    group1.ManualMax = null;
                    group1.ManualMin = null;
                    group1.LineWeight = 0.75;
                    group1.DateAxis = false;
                    group1.DisplayEmptyCellsAs = DispBlanksAs.Zero;
                    group1.Markers = false;
                    group1.High = false;
                    group1.Low = false;
                    group1.First = false;
                    group1.Last = false;
                    group1.Negative = false;
                    group1.DisplayXAxis = false;
                    group1.DisplayHidden = false;
                    group1.MinAxisType = SparklineAxisMinMax.Individual;
                    group1.MaxAxisType = SparklineAxisMinMax.Individual;
                    group1.RightToLeft = false;
                    group1.ColorSeries = Color.FromArgb(unchecked((int)0xFFFFFFFF));
                    group1.ColorNegative = Color.FromArgb(unchecked((int)0xFFFFFFFF));
                    group1.ColorAxis = Color.FromArgb(unchecked((int)0xFFFFFFFF));
                    group1.ColorMarkers = Color.FromArgb(unchecked((int)0xFFFFFFFF));
                    group1.ColorFirst = Color.FromArgb(unchecked((int)0xFFFFFFFF));
                    group1.ColorLast = Color.FromArgb(unchecked((int)0xFFFFFFFF));
                    group1.ColorHigh = Color.FromArgb(unchecked((int)0xFFFFFFFF));
                    group1.ColorLow = Color.FromArgb(unchecked((int)0xFFFFFFFF));
                    this.ValidateGroup(sheet2.SparklineGroups.SparklineGroups.First());

                    package.Save();
                }

                using (var package = new ExcelPackage(newFile))
                {
                    var groups = package.Workbook.Worksheets.First().SparklineGroups;
                    var group1 = groups.SparklineGroups[0];
                    Assert.AreEqual(SparklineType.Line, group1.Type);
                    Assert.IsNull(group1.ManualMax);
                    Assert.IsNull(group1.ManualMin);
                    Assert.IsNull(group1.LineWeight);
                    Assert.AreEqual(DispBlanksAs.Zero, group1.DisplayEmptyCellsAs);
                    Assert.IsFalse(group1.DateAxis);
                    Assert.IsFalse(group1.Markers);
                    Assert.IsFalse(group1.High);
                    Assert.IsFalse(group1.Low);
                    Assert.IsFalse(group1.First);
                    Assert.IsFalse(group1.Last);
                    Assert.IsFalse(group1.Negative);
                    Assert.IsFalse(group1.DisplayXAxis);
                    Assert.IsFalse(group1.DisplayHidden);
                    Assert.AreEqual(SparklineAxisMinMax.Individual, group1.MinAxisType);
                    Assert.AreEqual(SparklineAxisMinMax.Individual, group1.MaxAxisType);
                    Assert.IsFalse(group1.RightToLeft);
                    var expectedColor = Color.FromArgb(unchecked((int)0xFFFFFFFF));
                    Assert.AreEqual(expectedColor, group1.ColorSeries);
                    Assert.AreEqual(expectedColor, group1.ColorNegative);
                    Assert.AreEqual(expectedColor, group1.ColorAxis);
                    Assert.AreEqual(expectedColor, group1.ColorMarkers);
                    Assert.AreEqual(expectedColor, group1.ColorFirst);
                    Assert.AreEqual(expectedColor, group1.ColorLast);
                    Assert.AreEqual(expectedColor, group1.ColorHigh);
                    Assert.AreEqual(expectedColor, group1.ColorLow);
                    var copied = package.Workbook.Worksheets["Copied"];
                    this.ValidateGroup(copied.SparklineGroups.SparklineGroups.First());
                    copied.InsertRow(2, 2);
                    var group2 = copied.SparklineGroups.SparklineGroups[1];
                    var group3 = copied.SparklineGroups.SparklineGroups[2];
                    Assert.AreEqual("'Copied'!D9:F9", group2.Sparklines[0].Formula.Address);
                    Assert.AreEqual("'Copied'!D10:F10", group3.Sparklines[0].Formula.Address);
                    Assert.AreEqual("G9", group2.Sparklines[0].HostCell.Address);
                    Assert.AreEqual("G10", group3.Sparklines[0].HostCell.Address);
                }
            }
            finally
            {
                if (newFile.Exists)
                    newFile.Delete();
            }
        }

        private void ValidateGroup(ExcelSparklineGroup group)
        {
            Assert.AreEqual(SparklineType.Stacked, group.Type);
            Assert.AreEqual(1000.23, group.ManualMax);
            Assert.AreEqual(15.5, group.ManualMin);
            Assert.AreEqual(1.5, group.LineWeight);
            Assert.IsTrue(group.DateAxis);
            Assert.AreEqual(DispBlanksAs.Span, group.DisplayEmptyCellsAs);
            Assert.IsTrue(group.Markers);
            Assert.IsTrue(group.High);
            Assert.IsTrue(group.Low);
            Assert.IsTrue(group.First);
            Assert.IsTrue(group.Last);
            Assert.IsTrue(group.Negative);
            Assert.IsTrue(group.DisplayXAxis);
            Assert.IsTrue(group.DisplayHidden);
            Assert.AreEqual(SparklineAxisMinMax.Custom, group.MinAxisType);
            Assert.AreEqual(SparklineAxisMinMax.Group, group.MaxAxisType);
            Assert.AreEqual(Color.FromArgb(unchecked((int)0xAA375052)), group.ColorSeries);
            Assert.AreEqual(Color.FromArgb(unchecked((int)0xAB375052)), group.ColorNegative);
            Assert.AreEqual(Color.FromArgb(unchecked((int)0xAC375052)), group.ColorAxis);
            Assert.AreEqual(Color.FromArgb(unchecked((int)0xAD375052)), group.ColorMarkers);
            Assert.AreEqual(Color.FromArgb(unchecked((int)0xAE375052)), group.ColorFirst);
            Assert.AreEqual(Color.FromArgb(unchecked((int)0xAF375052)), group.ColorLast);
            Assert.AreEqual(Color.FromArgb(unchecked((int)0xBA375052)), group.ColorHigh);
            Assert.AreEqual(Color.FromArgb(unchecked((int)0xBB375052)), group.ColorLow);
            Assert.IsTrue(group.RightToLeft);
            Assert.AreEqual(2, group.Sparklines.Count);
            Assert.AreEqual("A1:B2", group.Sparklines.Last().Formula.Address);
            Assert.AreEqual("G9", group.Sparklines.Last().HostCell.Address);
        }
    }
}
