using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Sparkline;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace EPPlusTest
{
    [TestClass]
    public class SparklineTests
    {
        #region Test Constants
        string rootNode = @"<ext uri=""{05C60535-1F16-4fd2-B633-F4F36F0B64E0
    }"" xmlns:x14=""http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"">

            <x14:sparklineGroups xmlns:xm=""http://schemas.microsoft.com/office/excel/2006/main"">
				<x14:sparklineGroup type=""stacked"" displayEmptyCellsAs=""gap"" negative=""1"">
					<x14:colorSeries rgb=""FF376092""/>
					<x14:colorNegative rgb=""FFD00000""/>
					<x14:colorAxis rgb=""FF000000""/>
					<x14:colorMarkers rgb=""FFD00000""/>
					<x14:colorFirst rgb=""FFD00000""/>
					<x14:colorLast rgb=""FFD00000""/>
					<x14:colorHigh rgb=""FFD00000""/>
					<x14:colorLow rgb=""FFD00000""/>
					<x14:sparklines>
						<x14:sparkline>
							<xm:f>Sheet1!F7:F7</xm:f>
							<xm:sqref>G7</xm:sqref>
						</x14:sparkline>
						<x14:sparkline>
							<xm:f>Sheet1!F8:F8</xm:f>
							<xm:sqref>G8</xm:sqref>
						</x14:sparkline>
						<x14:sparkline>
							<xm:f>Sheet1!F9:F9</xm:f>
							<xm:sqref>G9</xm:sqref>
						</x14:sparkline>
						<x14:sparkline>
							<xm:f>Sheet1!F10:F10</xm:f>
							<xm:sqref>G10</xm:sqref>
						</x14:sparkline>
						<x14:sparkline>
							<xm:f>Sheet1!F11:F11</xm:f>
							<xm:sqref>G11</xm:sqref>
						</x14:sparkline>
						<x14:sparkline>
							<xm:f>Sheet1!F12:F12</xm:f>
							<xm:sqref>G12</xm:sqref>
						</x14:sparkline>
						<x14:sparkline>
							<xm:f>Sheet1!F13:F13</xm:f>
							<xm:sqref>G13</xm:sqref>
						</x14:sparkline>
						<x14:sparkline>
							<xm:f>Sheet1!F14:F14</xm:f>
							<xm:sqref>G14</xm:sqref>
						</x14:sparkline>
						<x14:sparkline>
							<xm:f>Sheet1!F15:F15</xm:f>
							<xm:sqref>G15</xm:sqref>
						</x14:sparkline>
						<x14:sparkline>
							<xm:f>Sheet1!F16:F16</xm:f>
							<xm:sqref>G16</xm:sqref>
						</x14:sparkline>
					</x14:sparklines>
				</x14:sparklineGroup>
				<x14:sparklineGroup type=""stacked"" displayEmptyCellsAs=""gap"" negative=""1"">
					<x14:colorSeries rgb=""FF376092""/>
					<x14:colorNegative rgb=""FFD00000""/>
					<x14:colorAxis rgb=""FF000000""/>
					<x14:colorMarkers rgb=""FFD00000""/>
					<x14:colorFirst rgb=""FFD00000""/>
					<x14:colorLast rgb=""FFD00000""/>
					<x14:colorHigh rgb=""FFD00000""/>
					<x14:colorLow rgb=""FFD00000""/>
					<x14:sparklines>
						<x14:sparkline>
							<xm:f>Sheet1!F6:F6</xm:f>
							<xm:sqref>G6</xm:sqref>
						</x14:sparkline>
						<x14:sparkline>
							<xm:f>Sheet1!F17:F17</xm:f>
							<xm:sqref>G17</xm:sqref>
						</x14:sparkline>
					</x14:sparklines>
				</x14:sparklineGroup>
			</x14:sparklineGroups>
        </ext>";
        #endregion

        [TestMethod]
        public void ParseExcelSparklineGroups()
        {
            XmlDocument extensions = new XmlDocument();
            extensions.LoadXml(this.rootNode);
            var sparklineGroupsXml = extensions.ChildNodes[0].ChildNodes[0];
            Assert.AreEqual(2, sparklineGroupsXml.ChildNodes.Count);
            XmlNamespaceManager manager = new XmlNamespaceManager(extensions.NameTable);
            // TODO: Figure out why I have to do this
            manager.AddNamespace("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            manager.AddNamespace("xm", "http://schemas.microsoft.com/office/excel/2006/main");

            var sparklineGroups = new ExcelSparklineGroups(null, manager, sparklineGroupsXml);
            Assert.AreEqual(2, sparklineGroups.SparklineGroups.Count);
            foreach (var group in sparklineGroups.SparklineGroups)
            {
                // Test that unused optional attributes are not set.
                Assert.IsNull(group.ManualMin);
                Assert.IsNull(group.ManualMax);
                Assert.IsNull(group.LineWeight);
                Assert.IsFalse(group.DateAxis);
                Assert.IsFalse(group.Markers);
                Assert.IsFalse(group.High);
                Assert.IsFalse(group.Low);
                Assert.IsFalse(group.First);
                Assert.IsFalse(group.Last);
                Assert.IsFalse(group.DisplayXAxis);
                //Assert.IsFalse(group.DisplayYAxis);
                Assert.AreEqual(SparklineAxisMinMax.Individual, group.MinAxisType);
                Assert.AreEqual(SparklineAxisMinMax.Individual, group.MaxAxisType);
                Assert.IsFalse(group.RightToLeft);

                // Test that attributes are parsed correctly.
                Assert.AreEqual(SparklineType.Stacked, group.Type);
                Assert.AreEqual(DispBlanksAs.Gap, group.DisplayEmptyCellsAs);
                Assert.AreEqual(true, group.Negative);

                // Test that color subnodes are parsed correctly.
                var darkRed = Color.FromArgb(unchecked((int)0xFFD00000));
                Assert.AreEqual(Color.FromArgb(unchecked((int)0xFF376092)), sparklineGroups.SparklineGroups[0].ColorSeries);
                Assert.AreEqual(darkRed, group.ColorNegative);
                Assert.AreEqual(Color.FromArgb(unchecked((int)0xFF000000)), group.ColorAxis);
                Assert.AreEqual(darkRed, group.ColorMarkers);
                Assert.AreEqual(darkRed, group.ColorFirst);
                Assert.AreEqual(darkRed, group.ColorLast);
                Assert.AreEqual(darkRed, group.ColorHigh);
                Assert.AreEqual(darkRed, group.ColorLow);
            }
            var group1 = sparklineGroups.SparklineGroups[0];
            var group2 = sparklineGroups.SparklineGroups[1];
            // Test that the Sparklines are parsed correctly.
            Assert.AreEqual(10, group1.Sparklines.Count);
            Assert.AreEqual(2, group2.Sparklines.Count);
            var firstCell = group1.Sparklines.First();
            Assert.AreEqual("G7", firstCell.HostCell.Address);
            Assert.AreEqual("Sheet1!F7:F7", firstCell.Formula.Address);
            var lastCell = group1.Sparklines.Last();
            Assert.AreEqual("G16", lastCell.HostCell.Address);
            Assert.AreEqual("Sheet1!F16:F16", lastCell.Formula.Address);
        }

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

                group1.Sparklines.Add(new ExcelSparkline(group1.NameSpaceManager) { Formula = new ExcelAddress("A1:B2"), HostCell = new ExcelAddress("G9") });
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

        [TestMethod]
        public void SaveSparklines()
        {
            FileInfo newFile = new FileInfo(Path.GetTempFileName());
            if (newFile.Exists)
                newFile.Delete();
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");
                var group = new ExcelSparklineGroup(sheet.NameSpaceManager);
                group.Sparklines.Add(new ExcelSparkline(sheet.NameSpaceManager) { Formula = new ExcelAddress("G1:G20"), HostCell = new ExcelAddress("B2") });
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

        // TODO: Delete this method once  you're done debugging
        [TestMethod]
        public void UpdateSparklinesAndSaveWorkbookDEBUG()
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

                group1.Sparklines.Add(new ExcelSparkline(group1.NameSpaceManager) { Formula = new ExcelAddress("A1:B2"), HostCell = new ExcelAddress("G9") });
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
                    var groups = package.Workbook.Worksheets.First().SparklineGroups;
                    var group1 = groups.SparklineGroups[0];
                    Assert.AreEqual(SparklineType.Stacked, group1.Type);
                    Assert.AreEqual(1000.23, group1.ManualMax);
                    Assert.AreEqual(15.5, group1.ManualMin);
                    Assert.AreEqual(1.5, group1.LineWeight);
                    Assert.IsTrue(group1.DateAxis);
                    Assert.AreEqual(DispBlanksAs.Span, group1.DisplayEmptyCellsAs);
                    Assert.IsTrue(group1.Markers);
                    Assert.IsTrue(group1.High);
                    Assert.IsTrue(group1.Low);
                    Assert.IsTrue(group1.First);
                    Assert.IsTrue(group1.Last);
                    Assert.IsTrue(group1.Negative);
                    Assert.IsTrue(group1.DisplayXAxis);
                    Assert.IsTrue(group1.DisplayHidden);
                    Assert.AreEqual(SparklineAxisMinMax.Custom, group1.MinAxisType);
                    Assert.AreEqual(SparklineAxisMinMax.Group, group1.MaxAxisType);
                    Assert.AreEqual(Color.FromArgb(unchecked((int)0xAA375052)), group1.ColorSeries);
                    Assert.AreEqual(Color.FromArgb(unchecked((int)0xAB375052)), group1.ColorNegative);
                    Assert.AreEqual(Color.FromArgb(unchecked((int)0xAC375052)), group1.ColorAxis);
                    Assert.AreEqual(Color.FromArgb(unchecked((int)0xAD375052)), group1.ColorMarkers);
                    Assert.AreEqual(Color.FromArgb(unchecked((int)0xAE375052)), group1.ColorFirst);
                    Assert.AreEqual(Color.FromArgb(unchecked((int)0xAF375052)), group1.ColorLast);
                    Assert.AreEqual(Color.FromArgb(unchecked((int)0xBA375052)), group1.ColorHigh);
                    Assert.AreEqual(Color.FromArgb(unchecked((int)0xBB375052)), group1.ColorLow);
                    Assert.IsTrue(group1.RightToLeft);
                    Assert.AreEqual(2, group1.Sparklines.Count);
                    Assert.AreEqual("A1:B2", group1.Sparklines.Last().Formula.Address);
                    Assert.AreEqual("G9", group1.Sparklines.Last().HostCell.Address);

                    // Ensure that unchanged Group2 and Group3 properties remain the same.
                    var group2 = groups.SparklineGroups[1];
                    var group3 = groups.SparklineGroups[2];
                    Assert.AreEqual(Color.FromArgb(unchecked((int)0xFF376092)), group2.ColorSeries);
                    Assert.AreEqual(Color.FromArgb(unchecked((int)0xFF323232)), group3.ColorSeries);
                    Assert.AreEqual(SparklineType.Line, group2.Type);
                    Assert.AreEqual(SparklineType.Stacked, group3.Type);
                    Assert.AreEqual("Sheet1!D7:F7", group2.Sparklines[0].Formula.Address);
                    Assert.AreEqual("Sheet1!D8:F8", group3.Sparklines[0].Formula.Address);
                    Assert.AreEqual("G7", group2.Sparklines[0].HostCell.Address);
                    Assert.AreEqual("G8", group3.Sparklines[0].HostCell.Address);

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
                    group1.ColorSeries = Color.Empty;
                    group1.ColorNegative = Color.Empty;
                    group1.ColorAxis = Color.Empty;
                    group1.ColorMarkers = Color.Empty;
                    group1.ColorFirst = Color.Empty;
                    group1.ColorLast = Color.Empty;
                    group1.ColorHigh = Color.Empty;
                    group1.ColorLow = Color.Empty;
                    package.Save();
                }

                // Ensure that the default property values are recognized as optional and not written to the worksheet.
                newFile.Refresh();
                Assert.IsTrue(newFile.Length < allPropertiesDefined);

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
                    Assert.IsTrue(group1.ColorSeries.IsEmpty);
                    Assert.IsTrue(group1.ColorNegative.IsEmpty);
                    Assert.IsTrue(group1.ColorAxis.IsEmpty);
                    Assert.IsTrue(group1.ColorMarkers.IsEmpty);
                    Assert.IsTrue(group1.ColorFirst.IsEmpty);
                    Assert.IsTrue(group1.ColorLast.IsEmpty);
                    Assert.IsTrue(group1.ColorHigh.IsEmpty);
                    Assert.IsTrue(group1.ColorLow.IsEmpty);
                }
            }
            finally
            {
                if (newFile.Exists)
                    newFile.Delete();
            }
        }

    }
}
