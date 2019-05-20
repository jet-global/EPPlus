using System.IO;
using EPPlusTest.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest.Table.PivotTable.Formats
{
	[TestClass]
	public class PivotTableConditionalFormattingTest
	{
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithConditionalFormatting.xlsx")]
		public void PivotTableRefreshConditionalFormattingWithDataBarAndGradedColorScale()
		{
			var file = new FileInfo("PivotTableWithConditionalFormatting.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable1"];
					package.Workbook.PivotCacheDefinitions.ForEach(c => c.UpdateData());
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B2:F17"), pivotTable.Address);
					Assert.AreEqual(10, pivotTable.Fields.Count);
					Assert.AreEqual(2, pivotTable.Worksheet.ConditionalFormatting.Count);
					Assert.AreEqual("E3:E17", pivotTable.Worksheet.ConditionalFormatting[0].Address.AddressSpaceSeparated);
					Assert.AreEqual("F3:F17", pivotTable.Worksheet.ConditionalFormatting[1].Address.AddressSpaceSeparated);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 2, 2, "Name"),
					new ExpectedCellValue(sheetName, 2, 3, "Country/Region Code"),
					new ExpectedCellValue(sheetName, 2, 4, "City"),
					new ExpectedCellValue(sheetName, 2, 5, "Sum of Net Change"),
					new ExpectedCellValue(sheetName, 2, 6, "Sum of Profit (LCY)"),
					new ExpectedCellValue(sheetName, 3, 2, "Antarcticopy"),
					new ExpectedCellValue(sheetName, 3, 3, "BE"),
					new ExpectedCellValue(sheetName, 3, 4, "Antwerpen"),
					new ExpectedCellValue(sheetName, 3, 5, 3218.14),
					new ExpectedCellValue(sheetName, 3, 6, 1119.44),
					new ExpectedCellValue(sheetName, 4, 2, "Autohaus Mielberg KG"),
					new ExpectedCellValue(sheetName, 4, 3, "DE"),
					new ExpectedCellValue(sheetName, 4, 4, "Hamburg 36"),
					new ExpectedCellValue(sheetName, 4, 5, 8794.86),
					new ExpectedCellValue(sheetName, 4, 6, 95.41),
					new ExpectedCellValue(sheetName, 5, 2, "BYT-KOMPLET s.r.o."),
					new ExpectedCellValue(sheetName, 5, 3, "CZ"),
					new ExpectedCellValue(sheetName, 5, 4, "Bojkovice"),
					new ExpectedCellValue(sheetName, 5, 5, 58518.63),
					new ExpectedCellValue(sheetName, 5, 6, 546),
					new ExpectedCellValue(sheetName, 6, 2, "Deerfield Graphics Company"),
					new ExpectedCellValue(sheetName, 6, 3, "US"),
					new ExpectedCellValue(sheetName, 6, 4, "Atlanta"),
					new ExpectedCellValue(sheetName, 6, 5, 1736.39),
					new ExpectedCellValue(sheetName, 6, 6, 1638.1),
					new ExpectedCellValue(sheetName, 7, 2, "Designstudio Gmunden"),
					new ExpectedCellValue(sheetName, 7, 3, "AT"),
					new ExpectedCellValue(sheetName, 7, 4, "Gmunden"),
					new ExpectedCellValue(sheetName, 7, 5, 3112.13),
					new ExpectedCellValue(sheetName, 7, 6, 847.2),
					new ExpectedCellValue(sheetName, 8, 2, "Englunds Kontorsmöbler AB"),
					new ExpectedCellValue(sheetName, 8, 3, "SE"),
					new ExpectedCellValue(sheetName, 8, 4, "Norrköbing"),
					new ExpectedCellValue(sheetName, 8, 5, 7841),
					new ExpectedCellValue(sheetName, 8, 6, 334.27),
					new ExpectedCellValue(sheetName, 9, 2, "Gagn & Gaman"),
					new ExpectedCellValue(sheetName, 9, 3, "IS"),
					new ExpectedCellValue(sheetName, 9, 4, "Hafnafjordur"),
					new ExpectedCellValue(sheetName, 9, 5, 86949.84),
					new ExpectedCellValue(sheetName, 9, 6, 259.97),
					new ExpectedCellValue(sheetName, 10, 2, "Guildford Water Department"),
					new ExpectedCellValue(sheetName, 10, 3, "US"),
					new ExpectedCellValue(sheetName, 10, 4, "Atlanta"),
					new ExpectedCellValue(sheetName, 10, 5, 822),
					new ExpectedCellValue(sheetName, 10, 6, 822),
					new ExpectedCellValue(sheetName, 11, 2, "Heimilisprydi"),
					new ExpectedCellValue(sheetName, 11, 3, "IS"),
					new ExpectedCellValue(sheetName, 11, 4, "Reykjavik"),
					new ExpectedCellValue(sheetName, 11, 5, 200615.42),
					new ExpectedCellValue(sheetName, 11, 6, 521.17),
					new ExpectedCellValue(sheetName, 12, 2, "John Haddock Insurance Co."),
					new ExpectedCellValue(sheetName, 12, 3, "US"),
					new ExpectedCellValue(sheetName, 12, 4, "Miami"),
					new ExpectedCellValue(sheetName, 12, 5, 537967),
					new ExpectedCellValue(sheetName, 12, 6, 4444.8),
					new ExpectedCellValue(sheetName, 13, 2, "Klubben"),
					new ExpectedCellValue(sheetName, 13, 3, "NO"),
					new ExpectedCellValue(sheetName, 13, 4, "Haslum"),
					new ExpectedCellValue(sheetName, 13, 5, 115966.31),
					new ExpectedCellValue(sheetName, 13, 6, 6349.7),
					new ExpectedCellValue(sheetName, 14, 2, "Progressive Home Furnishings"),
					new ExpectedCellValue(sheetName, 14, 3, "US"),
					new ExpectedCellValue(sheetName, 14, 4, "Chicago"),
					new ExpectedCellValue(sheetName, 14, 5, 2461),
					new ExpectedCellValue(sheetName, 14, 6, 621.6),
					new ExpectedCellValue(sheetName, 15, 2, "Selangorian Ltd."),
					new ExpectedCellValue(sheetName, 15, 3, "US"),
					new ExpectedCellValue(sheetName, 15, 4, "Chicago"),
					new ExpectedCellValue(sheetName, 15, 5, 147258.97),
					new ExpectedCellValue(sheetName, 15, 6, 3804.07),
					new ExpectedCellValue(sheetName, 16, 2, "The Cannon Group PLC"),
					new ExpectedCellValue(sheetName, 16, 3, "US"),
					new ExpectedCellValue(sheetName, 16, 4, "Atlanta"),
					new ExpectedCellValue(sheetName, 16, 5, 255797.35),
					new ExpectedCellValue(sheetName, 16, 6, 8148.48),
					new ExpectedCellValue(sheetName, 17, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 17, 5, 1431059.04),
					new ExpectedCellValue(sheetName, 17, 6, 29552.21),
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithConditionalFormatting.xlsx")]
		public void PivotTableRefreshConditionalFormattingWithDataBarGradedColorScaleAndIconSet()
		{
			var file = new FileInfo("PivotTableWithConditionalFormatting.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "Sheet2";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable1"];
					package.Workbook.PivotCacheDefinitions.ForEach(c => c.UpdateData());
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B2:G31"), pivotTable.Address);
					Assert.AreEqual(10, pivotTable.Fields.Count);
					Assert.AreEqual(3, pivotTable.Worksheet.ConditionalFormatting.Count);
					Assert.AreEqual("E3:E31", pivotTable.Worksheet.ConditionalFormatting[0].Address.AddressSpaceSeparated);
					Assert.AreEqual("F3:F31", pivotTable.Worksheet.ConditionalFormatting[1].Address.AddressSpaceSeparated);
					Assert.AreEqual("G3:G31", pivotTable.Worksheet.ConditionalFormatting[2].Address.AddressSpaceSeparated);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 2, 2, "Name"),
					new ExpectedCellValue(sheetName, 2, 3, "City"),
					new ExpectedCellValue(sheetName, 2, 4, "Last Date Modified"),
					new ExpectedCellValue(sheetName, 2, 5, "Sum of Net Change"),
					new ExpectedCellValue(sheetName, 2, 6, "Sum of Outstanding Orders"),
					new ExpectedCellValue(sheetName, 2, 7, "Sum of Profit (LCY)"),
					new ExpectedCellValue(sheetName, 3, 2, "Antarcticopy"),
					new ExpectedCellValue(sheetName, 3, 3, "Antwerpen"),
					new ExpectedCellValue(sheetName, 3, 5, null),
					new ExpectedCellValue(sheetName, 3, 6, null),
					new ExpectedCellValue(sheetName, 3, 7, null),
					new ExpectedCellValue(sheetName, 4, 4, "9/13/2012"),
					new ExpectedCellValue(sheetName, 4, 5, 3218.14),
					new ExpectedCellValue(sheetName, 4, 6, 0),
					new ExpectedCellValue(sheetName, 4, 7, 1119.44),
					new ExpectedCellValue(sheetName, 5, 2, "Autohaus Mielberg KG"),
					new ExpectedCellValue(sheetName, 5, 3, "Hamburg 36"),
					new ExpectedCellValue(sheetName, 5, 5, null),
					new ExpectedCellValue(sheetName, 5, 6, null),
					new ExpectedCellValue(sheetName, 5, 7, null),
					new ExpectedCellValue(sheetName, 6, 4, "9/13/2012"),
					new ExpectedCellValue(sheetName, 6, 5, 8794.86),
					new ExpectedCellValue(sheetName, 6, 6, 81399.66),
					new ExpectedCellValue(sheetName, 6, 7, 95.41),
					new ExpectedCellValue(sheetName, 7, 2, "BYT-KOMPLET s.r.o."),
					new ExpectedCellValue(sheetName, 7, 3, "Bojkovice"),
					new ExpectedCellValue(sheetName, 7, 5, null),
					new ExpectedCellValue(sheetName, 7, 6, null),
					new ExpectedCellValue(sheetName, 7, 7, null),
					new ExpectedCellValue(sheetName, 8, 4, "9/13/2012"),
					new ExpectedCellValue(sheetName, 8, 5, 58518.63),
					new ExpectedCellValue(sheetName, 8, 6, 0),
					new ExpectedCellValue(sheetName, 8, 7, 546),
					new ExpectedCellValue(sheetName, 9, 2, "Deerfield Graphics Company"),
					new ExpectedCellValue(sheetName, 9, 3, "Atlanta"),
					new ExpectedCellValue(sheetName, 9, 5, null),
					new ExpectedCellValue(sheetName, 9, 6, null),
					new ExpectedCellValue(sheetName, 9, 7, null),
					new ExpectedCellValue(sheetName, 10, 4, "9/13/2012"),
					new ExpectedCellValue(sheetName, 10, 5, 1736.39),
					new ExpectedCellValue(sheetName, 10, 6, 4390.52),
					new ExpectedCellValue(sheetName, 10, 7, 1638.1),
					new ExpectedCellValue(sheetName, 11, 2, "Designstudio Gmunden"),
					new ExpectedCellValue(sheetName, 11, 3, "Gmunden"),
					new ExpectedCellValue(sheetName, 11, 5, null),
					new ExpectedCellValue(sheetName, 11, 6, null),
					new ExpectedCellValue(sheetName, 11, 7, null),
					new ExpectedCellValue(sheetName, 12, 4, "9/13/2012"),
					new ExpectedCellValue(sheetName, 12, 5, 3112.13),
					new ExpectedCellValue(sheetName, 12, 6, 116491.69),
					new ExpectedCellValue(sheetName, 12, 7, 847.2),
					new ExpectedCellValue(sheetName, 13, 2, "Englunds Kontorsmöbler AB"),
					new ExpectedCellValue(sheetName, 13, 3, "Norrköbing"),
					new ExpectedCellValue(sheetName, 13, 5, null),
					new ExpectedCellValue(sheetName, 13, 6, null),
					new ExpectedCellValue(sheetName, 13, 7, null),
					new ExpectedCellValue(sheetName, 14, 4, "9/13/2012"),
					new ExpectedCellValue(sheetName, 14, 5, 7841),
					new ExpectedCellValue(sheetName, 14, 6, 6272.81),
					new ExpectedCellValue(sheetName, 14, 7, 334.27),
					new ExpectedCellValue(sheetName, 15, 2, "Gagn & Gaman"),
					new ExpectedCellValue(sheetName, 15, 3, "Hafnafjordur"),
					new ExpectedCellValue(sheetName, 15, 5, null),
					new ExpectedCellValue(sheetName, 15, 6, null),
					new ExpectedCellValue(sheetName, 15, 7, null),
					new ExpectedCellValue(sheetName, 16, 4, "9/13/2012"),
					new ExpectedCellValue(sheetName, 16, 5, 86949.84),
					new ExpectedCellValue(sheetName, 16, 6, 0),
					new ExpectedCellValue(sheetName, 16, 7, 259.97),
					new ExpectedCellValue(sheetName, 17, 2, "Guildford Water Department"),
					new ExpectedCellValue(sheetName, 17, 3, "Atlanta"),
					new ExpectedCellValue(sheetName, 17, 5, null),
					new ExpectedCellValue(sheetName, 17, 6, null),
					new ExpectedCellValue(sheetName, 17, 7, null),
					new ExpectedCellValue(sheetName, 18, 4, "9/13/2012"),
					new ExpectedCellValue(sheetName, 18, 5, 822),
					new ExpectedCellValue(sheetName, 18, 6, 0),
					new ExpectedCellValue(sheetName, 18, 7, 822),
					new ExpectedCellValue(sheetName, 19, 2, "Heimilisprydi"),
					new ExpectedCellValue(sheetName, 19, 3, "Reykjavik"),
					new ExpectedCellValue(sheetName, 19, 5, null),
					new ExpectedCellValue(sheetName, 19, 6, null),
					new ExpectedCellValue(sheetName, 19, 7, null),
					new ExpectedCellValue(sheetName, 20, 4, "9/13/2012"),
					new ExpectedCellValue(sheetName, 20, 5, 200615.42),
					new ExpectedCellValue(sheetName, 20, 6, 0),
					new ExpectedCellValue(sheetName, 20, 7, 521.17),
					new ExpectedCellValue(sheetName, 21, 2, "John Haddock Insurance Co."),
					new ExpectedCellValue(sheetName, 21, 3, "Miami"),
					new ExpectedCellValue(sheetName, 21, 5, null),
					new ExpectedCellValue(sheetName, 21, 6, null),
					new ExpectedCellValue(sheetName, 21, 7, null),
					new ExpectedCellValue(sheetName, 22, 4, "9/13/2012"),
					new ExpectedCellValue(sheetName, 22, 5, 537967),
					new ExpectedCellValue(sheetName, 22, 6, 10820.18),
					new ExpectedCellValue(sheetName, 22, 7, 4444.8),
					new ExpectedCellValue(sheetName, 23, 2, "Klubben"),
					new ExpectedCellValue(sheetName, 23, 3, "Haslum"),
					new ExpectedCellValue(sheetName, 23, 5, null),
					new ExpectedCellValue(sheetName, 23, 6, null),
					new ExpectedCellValue(sheetName, 23, 7, null),
					new ExpectedCellValue(sheetName, 24, 4, "9/13/2012"),
					new ExpectedCellValue(sheetName, 24, 5, 115966.31),
					new ExpectedCellValue(sheetName, 24, 6, 0),
					new ExpectedCellValue(sheetName, 24, 7, 6349.7),
					new ExpectedCellValue(sheetName, 25, 2, "Progressive Home Furnishings"),
					new ExpectedCellValue(sheetName, 25, 3, "Chicago"),
					new ExpectedCellValue(sheetName, 25, 5, null),
					new ExpectedCellValue(sheetName, 25, 6, null),
					new ExpectedCellValue(sheetName, 25, 7, null),
					new ExpectedCellValue(sheetName, 26, 4, "9/13/2012"),
					new ExpectedCellValue(sheetName, 26, 5, 2461),
					new ExpectedCellValue(sheetName, 26, 6, 0),
					new ExpectedCellValue(sheetName, 26, 7, 621.6),
					new ExpectedCellValue(sheetName, 27, 2, "Selangorian Ltd."),
					new ExpectedCellValue(sheetName, 27, 3, "Chicago"),
					new ExpectedCellValue(sheetName, 27, 5, null),
					new ExpectedCellValue(sheetName, 27, 6, null),
					new ExpectedCellValue(sheetName, 27, 7, null),
					new ExpectedCellValue(sheetName, 28, 4, "9/13/2012"),
					new ExpectedCellValue(sheetName, 28, 5, 147258.97),
					new ExpectedCellValue(sheetName, 28, 6, 8548.97),
					new ExpectedCellValue(sheetName, 28, 7, 3804.07),
					new ExpectedCellValue(sheetName, 29, 2, "The Cannon Group PLC"),
					new ExpectedCellValue(sheetName, 29, 3, "Atlanta"),
					new ExpectedCellValue(sheetName, 29, 5, null),
					new ExpectedCellValue(sheetName, 29, 6, null),
					new ExpectedCellValue(sheetName, 29, 7, null),
					new ExpectedCellValue(sheetName, 30, 4, "9/13/2012"),
					new ExpectedCellValue(sheetName, 30, 5, 255797.35),
					new ExpectedCellValue(sheetName, 30, 6, 1354.5),
					new ExpectedCellValue(sheetName, 30, 7, 8148.48),
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithConditionalFormatting.xlsx")]
		public void PivotTableRefreshConditionalFormattingSelectedCellsWithMultipleDataFields()
		{
			var file = new FileInfo("PivotTableWithConditionalFormatting.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "Sheet3";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable2"];
					package.Workbook.PivotCacheDefinitions.ForEach(c => c.UpdateData());
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B2:F17"), pivotTable.Address);
					Assert.AreEqual(10, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 2, 2, "City"),
					new ExpectedCellValue(sheetName, 2, 3, "Contact"),
					new ExpectedCellValue(sheetName, 2, 4, "Country/Region Code"),
					new ExpectedCellValue(sheetName, 2, 5, "Sum of Net Change"),
					new ExpectedCellValue(sheetName, 2, 6, "Sum of Outstanding Orders"),
					new ExpectedCellValue(sheetName, 3, 2, "Antwerpen"),
					new ExpectedCellValue(sheetName, 3, 3, "Michael Zeman"),
					new ExpectedCellValue(sheetName, 3, 4, "BE"),
					new ExpectedCellValue(sheetName, 3, 5, 3218.14),
					new ExpectedCellValue(sheetName, 3, 6, 0),
					new ExpectedCellValue(sheetName, 4, 2, "Atlanta"),
					new ExpectedCellValue(sheetName, 4, 3, "Mr. Andy Teal"),
					new ExpectedCellValue(sheetName, 4, 4, "US"),
					new ExpectedCellValue(sheetName, 4, 5, 255797.35),
					new ExpectedCellValue(sheetName, 4, 6, 1354.5),
					new ExpectedCellValue(sheetName, 5, 3, "Mr. Jim Stewart"),
					new ExpectedCellValue(sheetName, 5, 4, "US"),
					new ExpectedCellValue(sheetName, 5, 5, 822),
					new ExpectedCellValue(sheetName, 5, 6, 0),
					new ExpectedCellValue(sheetName, 6, 3, "Mr. Kevin Wright"),
					new ExpectedCellValue(sheetName, 6, 4, "US"),
					new ExpectedCellValue(sheetName, 6, 5, 1736.39),
					new ExpectedCellValue(sheetName, 6, 6, 4390.52),
					new ExpectedCellValue(sheetName, 7, 2, "Bojkovice"),
					new ExpectedCellValue(sheetName, 7, 3, "Milos Silhan"),
					new ExpectedCellValue(sheetName, 7, 4, "CZ"),
					new ExpectedCellValue(sheetName, 7, 5, 58518.63),
					new ExpectedCellValue(sheetName, 7, 6, 0),
					new ExpectedCellValue(sheetName, 8, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 8, 3, "Mr. Mark McArthur"),
					new ExpectedCellValue(sheetName, 8, 4, "US"),
					new ExpectedCellValue(sheetName, 8, 5, 147258.97),
					new ExpectedCellValue(sheetName, 8, 6, 8548.97),
					new ExpectedCellValue(sheetName, 9, 3, "Mr. Scott Mitchell"),
					new ExpectedCellValue(sheetName, 9, 4, "US"),
					new ExpectedCellValue(sheetName, 9, 5, 2461),
					new ExpectedCellValue(sheetName, 9, 6, 0),
					new ExpectedCellValue(sheetName, 10, 2, "Gmunden"),
					new ExpectedCellValue(sheetName, 10, 3, "Fr. Birgitte Vestphael"),
					new ExpectedCellValue(sheetName, 10, 4, "AT"),
					new ExpectedCellValue(sheetName, 10, 5, 3112.13),
					new ExpectedCellValue(sheetName, 10, 6, 116491.69),
					new ExpectedCellValue(sheetName, 11, 2, "Hafnafjordur"),
					new ExpectedCellValue(sheetName, 11, 3, "Ragnheidur K. Gudmundsdottir"),
					new ExpectedCellValue(sheetName, 11, 4, "IS"),
					new ExpectedCellValue(sheetName, 11, 5, 86949.84),
					new ExpectedCellValue(sheetName, 11, 6, 0),
					new ExpectedCellValue(sheetName, 12, 2, "Hamburg 36"),
					new ExpectedCellValue(sheetName, 12, 3, ""),
					new ExpectedCellValue(sheetName, 12, 4, "DE"),
					new ExpectedCellValue(sheetName, 12, 5, 8794.86),
					new ExpectedCellValue(sheetName, 12, 6, 81399.66),
					new ExpectedCellValue(sheetName, 13, 2, "Haslum"),
					new ExpectedCellValue(sheetName, 13, 3, "Thomas Andersen"),
					new ExpectedCellValue(sheetName, 13, 4, "NO"),
					new ExpectedCellValue(sheetName, 13, 5, 115966.31),
					new ExpectedCellValue(sheetName, 13, 6, 0),
					new ExpectedCellValue(sheetName, 14, 2, "Miami"),
					new ExpectedCellValue(sheetName, 14, 3, "Miss Patricia Doyle"),
					new ExpectedCellValue(sheetName, 14, 4, "US"),
					new ExpectedCellValue(sheetName, 14, 5, 537967),
					new ExpectedCellValue(sheetName, 14, 6, 10820.18),
					new ExpectedCellValue(sheetName, 15, 2, "Norrköbing"),
					new ExpectedCellValue(sheetName, 15, 3, ""),
					new ExpectedCellValue(sheetName, 15, 4, "SE"),
					new ExpectedCellValue(sheetName, 15, 5, 7841),
					new ExpectedCellValue(sheetName, 15, 6, 6272.81),
					new ExpectedCellValue(sheetName, 16, 2, "Reykjavik"),
					new ExpectedCellValue(sheetName, 16, 3, "Gunnar Orn Thorsteinsson"),
					new ExpectedCellValue(sheetName, 16, 4, "IS"),
					new ExpectedCellValue(sheetName, 16, 5, 200615.42),
					new ExpectedCellValue(sheetName, 16, 6, 0),
					new ExpectedCellValue(sheetName, 17, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 17, 5, 1431059.04),
					new ExpectedCellValue(sheetName, 17, 6, 229278.33),
				});
			}
		}
	}
}
