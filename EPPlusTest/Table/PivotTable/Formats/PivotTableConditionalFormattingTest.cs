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

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithConditionalFormatting.xlsx")]
		public void PivotTableConditionalFormattingApplyToDataFieldsWithDifferentTableForms()
		{
			var file = new FileInfo("PivotTableWithConditionalFormatting.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "Sheet3";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable1"];
					package.Workbook.PivotCacheDefinitions.ForEach(c => c.UpdateData());
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B23:F66"), pivotTable.Address);
					Assert.AreEqual(10, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 23, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 23, 3, "City"),
					new ExpectedCellValue(sheetName, 23, 4, "Country/Region Code"),
					new ExpectedCellValue(sheetName, 23, 5, "Sum of Net Change"),
					new ExpectedCellValue(sheetName, 23, 6, "Sum of Outstanding Orders"),
					new ExpectedCellValue(sheetName, 24, 2, "Antarcticopy"),
					new ExpectedCellValue(sheetName, 24, 3, "Antwerpen"),
					new ExpectedCellValue(sheetName, 24, 5, 3218.14),
					new ExpectedCellValue(sheetName, 24, 6, 0),
					new ExpectedCellValue(sheetName, 25, 4, "BE"),
					new ExpectedCellValue(sheetName, 25, 5, 3218.14),
					new ExpectedCellValue(sheetName, 25, 6, 0),
					new ExpectedCellValue(sheetName, 26, 2, "Antarcticopy Total"),
					new ExpectedCellValue(sheetName, 26, 5, 3218.14),
					new ExpectedCellValue(sheetName, 26, 6, 0),
					new ExpectedCellValue(sheetName, 27, 2, "Autohaus Mielberg KG"),
					new ExpectedCellValue(sheetName, 27, 3, "Hamburg 36"),
					new ExpectedCellValue(sheetName, 27, 5, 8794.86),
					new ExpectedCellValue(sheetName, 27, 6, 81399.66),
					new ExpectedCellValue(sheetName, 28, 4, "DE"),
					new ExpectedCellValue(sheetName, 28, 5, 8794.86),
					new ExpectedCellValue(sheetName, 28, 6, 81399.66),
					new ExpectedCellValue(sheetName, 29, 2, "Autohaus Mielberg KG Total"),
					new ExpectedCellValue(sheetName, 29, 5, 8794.86),
					new ExpectedCellValue(sheetName, 29, 6, 81399.66),
					new ExpectedCellValue(sheetName, 30, 2, "BYT-KOMPLET s.r.o."),
					new ExpectedCellValue(sheetName, 30, 3, "Bojkovice"),
					new ExpectedCellValue(sheetName, 30, 5, 58518.63),
					new ExpectedCellValue(sheetName, 30, 6, 0),
					new ExpectedCellValue(sheetName, 31, 4, "CZ"),
					new ExpectedCellValue(sheetName, 31, 5, 58518.63),
					new ExpectedCellValue(sheetName, 31, 6, 0),
					new ExpectedCellValue(sheetName, 32, 2, "BYT-KOMPLET s.r.o. Total"),
					new ExpectedCellValue(sheetName, 32, 5, 58518.63),
					new ExpectedCellValue(sheetName, 32, 6, 0),
					new ExpectedCellValue(sheetName, 33, 2, "Deerfield Graphics Company"),
					new ExpectedCellValue(sheetName, 33, 3, "Atlanta"),
					new ExpectedCellValue(sheetName, 33, 5, 1736.39),
					new ExpectedCellValue(sheetName, 33, 6, 4390.52),
					new ExpectedCellValue(sheetName, 34, 4, "US"),
					new ExpectedCellValue(sheetName, 34, 5, 1736.39),
					new ExpectedCellValue(sheetName, 34, 6, 4390.52),
					new ExpectedCellValue(sheetName, 35, 2, "Deerfield Graphics Company Total"),
					new ExpectedCellValue(sheetName, 35, 5, 1736.39),
					new ExpectedCellValue(sheetName, 35, 6, 4390.52),
					new ExpectedCellValue(sheetName, 36, 2, "Designstudio Gmunden"),
					new ExpectedCellValue(sheetName, 36, 3, "Gmunden"),
					new ExpectedCellValue(sheetName, 36, 5, 3112.13),
					new ExpectedCellValue(sheetName, 36, 6, 116491.69),
					new ExpectedCellValue(sheetName, 37, 4, "AT"),
					new ExpectedCellValue(sheetName, 37, 5, 3112.13),
					new ExpectedCellValue(sheetName, 37, 6, 116491.69),
					new ExpectedCellValue(sheetName, 38, 2, "Designstudio Gmunden Total"),
					new ExpectedCellValue(sheetName, 38, 5, 3112.13),
					new ExpectedCellValue(sheetName, 38, 6, 116491.69),
					new ExpectedCellValue(sheetName, 39, 2, "Englunds Kontorsmöbler AB"),
					new ExpectedCellValue(sheetName, 39, 3, "Norrköbing"),
					new ExpectedCellValue(sheetName, 39, 5, 7841),
					new ExpectedCellValue(sheetName, 39, 6, 6272.81),
					new ExpectedCellValue(sheetName, 40, 4, "SE"),
					new ExpectedCellValue(sheetName, 40, 5, 7841),
					new ExpectedCellValue(sheetName, 40, 6, 6272.81),
					new ExpectedCellValue(sheetName, 41, 2, "Englunds Kontorsmöbler AB Total"),
					new ExpectedCellValue(sheetName, 41, 5, 7841),
					new ExpectedCellValue(sheetName, 41, 6, 6272.81),
					new ExpectedCellValue(sheetName, 42, 2, "Gagn & Gaman"),
					new ExpectedCellValue(sheetName, 42, 3, "Hafnafjordur"),
					new ExpectedCellValue(sheetName, 42, 5, 86949.84),
					new ExpectedCellValue(sheetName, 42, 6, 0),
					new ExpectedCellValue(sheetName, 43, 4, "IS"),
					new ExpectedCellValue(sheetName, 43, 5, 86949.84),
					new ExpectedCellValue(sheetName, 43, 6, 0),
					new ExpectedCellValue(sheetName, 44, 2, "Gagn & Gaman Total"),
					new ExpectedCellValue(sheetName, 44, 5, 86949.84),
					new ExpectedCellValue(sheetName, 44, 6, 0),
					new ExpectedCellValue(sheetName, 45, 2, "Guildford Water Department"),
					new ExpectedCellValue(sheetName, 45, 3, "Atlanta"),
					new ExpectedCellValue(sheetName, 45, 5, 822),
					new ExpectedCellValue(sheetName, 45, 6, 0),
					new ExpectedCellValue(sheetName, 46, 4, "US"),
					new ExpectedCellValue(sheetName, 46, 5, 822),
					new ExpectedCellValue(sheetName, 46, 6, 0),
					new ExpectedCellValue(sheetName, 47, 2, "Guildford Water Department Total"),
					new ExpectedCellValue(sheetName, 47, 5, 822),
					new ExpectedCellValue(sheetName, 47, 6, 0),
					new ExpectedCellValue(sheetName, 48, 2, "Heimilisprydi"),
					new ExpectedCellValue(sheetName, 48, 3, "Reykjavik"),
					new ExpectedCellValue(sheetName, 48, 5, 200615.42),
					new ExpectedCellValue(sheetName, 48, 6, 0),
					new ExpectedCellValue(sheetName, 49, 4, "IS"),
					new ExpectedCellValue(sheetName, 49, 5, 200615.42),
					new ExpectedCellValue(sheetName, 49, 6, 0),
					new ExpectedCellValue(sheetName, 50, 2, "Heimilisprydi Total"),
					new ExpectedCellValue(sheetName, 50, 5, 200615.42),
					new ExpectedCellValue(sheetName, 50, 6, 0),
					new ExpectedCellValue(sheetName, 51, 2, "John Haddock Insurance Co."),
					new ExpectedCellValue(sheetName, 51, 3, "Miami"),
					new ExpectedCellValue(sheetName, 51, 5, 537967),
					new ExpectedCellValue(sheetName, 51, 6, 10820.18),
					new ExpectedCellValue(sheetName, 52, 4, "US"),
					new ExpectedCellValue(sheetName, 52, 5, 537967),
					new ExpectedCellValue(sheetName, 52, 6, 10820.18),
					new ExpectedCellValue(sheetName, 53, 2, "John Haddock Insurance Co. Total"),
					new ExpectedCellValue(sheetName, 53, 5, 537967),
					new ExpectedCellValue(sheetName, 53, 6, 10820.18),
					new ExpectedCellValue(sheetName, 54, 2, "Klubben"),
					new ExpectedCellValue(sheetName, 54, 3, "Haslum"),
					new ExpectedCellValue(sheetName, 54, 5, 115966.31),
					new ExpectedCellValue(sheetName, 54, 6, 0),
					new ExpectedCellValue(sheetName, 55, 4, "NO"),
					new ExpectedCellValue(sheetName, 55, 5, 115966.31),
					new ExpectedCellValue(sheetName, 55, 6, 0),
					new ExpectedCellValue(sheetName, 56, 2, "Klubben Total"),
					new ExpectedCellValue(sheetName, 56, 5, 115966.31),
					new ExpectedCellValue(sheetName, 56, 6, 0),
					new ExpectedCellValue(sheetName, 57, 2, "Progressive Home Furnishings"),
					new ExpectedCellValue(sheetName, 57, 3, "Chicago"),
					new ExpectedCellValue(sheetName, 57, 5, 2461),
					new ExpectedCellValue(sheetName, 57, 6, 0),
					new ExpectedCellValue(sheetName, 58, 4, "US"),
					new ExpectedCellValue(sheetName, 58, 5, 2461),
					new ExpectedCellValue(sheetName, 58, 6, 0),
					new ExpectedCellValue(sheetName, 59, 2, "Progressive Home Furnishings Total"),
					new ExpectedCellValue(sheetName, 59, 5, 2461),
					new ExpectedCellValue(sheetName, 59, 6, 0),
					new ExpectedCellValue(sheetName, 60, 2, "Selangorian Ltd."),
					new ExpectedCellValue(sheetName, 60, 3, "Chicago"),
					new ExpectedCellValue(sheetName, 60, 5, 147258.97),
					new ExpectedCellValue(sheetName, 60, 6, 8548.97),
					new ExpectedCellValue(sheetName, 61, 4, "US"),
					new ExpectedCellValue(sheetName, 61, 5, 147258.97),
					new ExpectedCellValue(sheetName, 61, 6, 8548.97),
					new ExpectedCellValue(sheetName, 62, 2, "Selangorian Ltd. Total"),
					new ExpectedCellValue(sheetName, 62, 5, 147258.97),
					new ExpectedCellValue(sheetName, 62, 6, 8548.97),
					new ExpectedCellValue(sheetName, 63, 2, "The Cannon Group PLC"),
					new ExpectedCellValue(sheetName, 63, 3, "Atlanta"),
					new ExpectedCellValue(sheetName, 63, 5, 255797.35),
					new ExpectedCellValue(sheetName, 63, 6, 1354.5),
					new ExpectedCellValue(sheetName, 64, 4, "US"),
					new ExpectedCellValue(sheetName, 64, 5, 255797.35),
					new ExpectedCellValue(sheetName, 64, 6, 1354.5),
					new ExpectedCellValue(sheetName, 65, 2, "The Cannon Group PLC Total"),
					new ExpectedCellValue(sheetName, 65, 5, 255797.35),
					new ExpectedCellValue(sheetName, 65, 6, 1354.5),
					new ExpectedCellValue(sheetName, 66, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 66, 5, 1431059.04),
					new ExpectedCellValue(sheetName, 66, 6, 229278.33),
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithConditionalFormatting.xlsx")]
		public void PivotTableConditionalFormattingApplyToDataFieldsWithCompactForms()
		{
			var file = new FileInfo("PivotTableWithConditionalFormatting.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "Sheet3";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable3"];
					package.Workbook.PivotCacheDefinitions.ForEach(c => c.UpdateData());
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("I2:K39"), pivotTable.Address);
					Assert.AreEqual(10, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 2, 9, "Row Labels"),
					new ExpectedCellValue(sheetName, 2, 10, "Sum of Net Change"),
					new ExpectedCellValue(sheetName, 2, 11, "Sum of Outstanding Orders"),
					new ExpectedCellValue(sheetName, 3, 9, "Antwerpen"),
					new ExpectedCellValue(sheetName, 3, 10, 3218.14),
					new ExpectedCellValue(sheetName, 3, 11, 0),
					new ExpectedCellValue(sheetName, 4, 9, "BE"),
					new ExpectedCellValue(sheetName, 4, 10, 3218.14),
					new ExpectedCellValue(sheetName, 4, 11, 0),
					new ExpectedCellValue(sheetName, 5, 9, "Michael Zeman"),
					new ExpectedCellValue(sheetName, 5, 10, 3218.14),
					new ExpectedCellValue(sheetName, 5, 11, 0),
					new ExpectedCellValue(sheetName, 6, 9, "Atlanta"),
					new ExpectedCellValue(sheetName, 6, 10, 258355.74),
					new ExpectedCellValue(sheetName, 6, 11, 5745.02),
					new ExpectedCellValue(sheetName, 7, 9, "US"),
					new ExpectedCellValue(sheetName, 7, 10, 258355.74),
					new ExpectedCellValue(sheetName, 7, 11, 5745.02),
					new ExpectedCellValue(sheetName, 8, 9, "Mr. Andy Teal"),
					new ExpectedCellValue(sheetName, 8, 10, 255797.35),
					new ExpectedCellValue(sheetName, 8, 11, 1354.5),
					new ExpectedCellValue(sheetName, 9, 9, "Mr. Jim Stewart"),
					new ExpectedCellValue(sheetName, 9, 10, 822),
					new ExpectedCellValue(sheetName, 9, 11, 0),
					new ExpectedCellValue(sheetName, 10, 9, "Mr. Kevin Wright"),
					new ExpectedCellValue(sheetName, 10, 10, 1736.39),
					new ExpectedCellValue(sheetName, 10, 11, 4390.52),
					new ExpectedCellValue(sheetName, 11, 9, "Bojkovice"),
					new ExpectedCellValue(sheetName, 11, 10, 58518.63),
					new ExpectedCellValue(sheetName, 11, 11, 0),
					new ExpectedCellValue(sheetName, 12, 9, "CZ"),
					new ExpectedCellValue(sheetName, 12, 10, 58518.63),
					new ExpectedCellValue(sheetName, 12, 11, 0),
					new ExpectedCellValue(sheetName, 13, 9, "Milos Silhan"),
					new ExpectedCellValue(sheetName, 13, 10, 58518.63),
					new ExpectedCellValue(sheetName, 13, 11, 0),
					new ExpectedCellValue(sheetName, 14, 9, "Chicago"),
					new ExpectedCellValue(sheetName, 14, 10, 149719.97),
					new ExpectedCellValue(sheetName, 14, 11, 8548.97),
					new ExpectedCellValue(sheetName, 15, 9, "US"),
					new ExpectedCellValue(sheetName, 15, 10, 149719.97),
					new ExpectedCellValue(sheetName, 15, 11, 8548.97),
					new ExpectedCellValue(sheetName, 16, 9, "Mr. Mark McArthur"),
					new ExpectedCellValue(sheetName, 16, 10, 147258.97),
					new ExpectedCellValue(sheetName, 16, 11, 8548.97),
					new ExpectedCellValue(sheetName, 17, 9, "Mr. Scott Mitchell"),
					new ExpectedCellValue(sheetName, 17, 10, 2461),
					new ExpectedCellValue(sheetName, 17, 11, 0),
					new ExpectedCellValue(sheetName, 18, 9, "Gmunden"),
					new ExpectedCellValue(sheetName, 18, 10, 3112.13),
					new ExpectedCellValue(sheetName, 18, 11, 116491.69),
					new ExpectedCellValue(sheetName, 19, 9, "AT"),
					new ExpectedCellValue(sheetName, 19, 10, 3112.13),
					new ExpectedCellValue(sheetName, 19, 11, 116491.69),
					new ExpectedCellValue(sheetName, 20, 9, "Fr. Birgitte Vestphael"),
					new ExpectedCellValue(sheetName, 20, 10, 3112.13),
					new ExpectedCellValue(sheetName, 20, 11, 116491.69),
					new ExpectedCellValue(sheetName, 21, 9, "Hafnafjordur"),
					new ExpectedCellValue(sheetName, 21, 10, 86949.84),
					new ExpectedCellValue(sheetName, 21, 11, 0),
					new ExpectedCellValue(sheetName, 22, 9, "IS"),
					new ExpectedCellValue(sheetName, 22, 10, 86949.84),
					new ExpectedCellValue(sheetName, 22, 11, 0),
					new ExpectedCellValue(sheetName, 23, 9, "Ragnheidur K. Gudmundsdottir"),
					new ExpectedCellValue(sheetName, 23, 10, 86949.84),
					new ExpectedCellValue(sheetName, 23, 11, 0),
					new ExpectedCellValue(sheetName, 24, 9, "Hamburg 36"),
					new ExpectedCellValue(sheetName, 24, 10, 8794.86),
					new ExpectedCellValue(sheetName, 24, 11, 81399.66),
					new ExpectedCellValue(sheetName, 25, 9, "DE"),
					new ExpectedCellValue(sheetName, 25, 10, 8794.86),
					new ExpectedCellValue(sheetName, 25, 11, 81399.66),
					new ExpectedCellValue(sheetName, 26, 9, ""),
					new ExpectedCellValue(sheetName, 26, 10, 8794.86),
					new ExpectedCellValue(sheetName, 26, 11, 81399.66),
					new ExpectedCellValue(sheetName, 27, 9, "Haslum"),
					new ExpectedCellValue(sheetName, 27, 10, 115966.31),
					new ExpectedCellValue(sheetName, 27, 11, 0),
					new ExpectedCellValue(sheetName, 28, 9, "NO"),
					new ExpectedCellValue(sheetName, 28, 10, 115966.31),
					new ExpectedCellValue(sheetName, 28, 11, 0),
					new ExpectedCellValue(sheetName, 29, 9, "Thomas Andersen"),
					new ExpectedCellValue(sheetName, 29, 10, 115966.31),
					new ExpectedCellValue(sheetName, 29, 11, 0),
					new ExpectedCellValue(sheetName, 30, 9, "Miami"),
					new ExpectedCellValue(sheetName, 30, 10, 537967),
					new ExpectedCellValue(sheetName, 30, 11, 10820.18),
					new ExpectedCellValue(sheetName, 31, 9, "US"),
					new ExpectedCellValue(sheetName, 31, 10, 537967),
					new ExpectedCellValue(sheetName, 31, 11, 10820.18),
					new ExpectedCellValue(sheetName, 32, 9, "Miss Patricia Doyle"),
					new ExpectedCellValue(sheetName, 32, 10, 537967),
					new ExpectedCellValue(sheetName, 32, 11, 10820.18),
					new ExpectedCellValue(sheetName, 33, 9, "Norrköbing"),
					new ExpectedCellValue(sheetName, 33, 10, 7841),
					new ExpectedCellValue(sheetName, 33, 11, 6272.81),
					new ExpectedCellValue(sheetName, 34, 9, "SE"),
					new ExpectedCellValue(sheetName, 34, 10, 7841),
					new ExpectedCellValue(sheetName, 34, 11, 6272.81),
					new ExpectedCellValue(sheetName, 35, 9, ""),
					new ExpectedCellValue(sheetName, 35, 10, 7841),
					new ExpectedCellValue(sheetName, 35, 11, 6272.81),
					new ExpectedCellValue(sheetName, 36, 9, "Reykjavik"),
					new ExpectedCellValue(sheetName, 36, 10, 200615.42),
					new ExpectedCellValue(sheetName, 36, 11, 0),
					new ExpectedCellValue(sheetName, 37, 9, "IS"),
					new ExpectedCellValue(sheetName, 37, 10, 200615.42),
					new ExpectedCellValue(sheetName, 37, 11, 0),
					new ExpectedCellValue(sheetName, 38, 9, "Gunnar Orn Thorsteinsson"),
					new ExpectedCellValue(sheetName, 38, 10, 200615.42),
					new ExpectedCellValue(sheetName, 38, 11, 0),
					new ExpectedCellValue(sheetName, 39, 9, "Grand Total"),
					new ExpectedCellValue(sheetName, 39, 10, 1431059.04),
					new ExpectedCellValue(sheetName, 39, 11, 229278.33),
				});
			}
		}
	}
}
