using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest
{
	[TestClass]
	public class ExcelRangeBaseTest : TestBase
	{
		#region Copy Tests
		[TestMethod]
		public void CopyCopiesCommentsFromSingleCellRanges()
		{
			InitBase();
			var pck = new ExcelPackage();
			var ws1 = pck.Workbook.Worksheets.Add("CommentCopying");
			var sourceExcelRange = ws1.Cells[3, 3];
			Assert.IsNull(sourceExcelRange.Comment);
			sourceExcelRange.AddComment("Testing comment 1", "test1");
			Assert.AreEqual("test1", sourceExcelRange.Comment.Author);
			Assert.AreEqual("Testing comment 1", sourceExcelRange.Comment.Text);
			var destinationExcelRange = ws1.Cells[5, 5];
			Assert.IsNull(destinationExcelRange.Comment);
			sourceExcelRange.Copy(destinationExcelRange);
			// Assert the original comment is intact.
			Assert.AreEqual("test1", sourceExcelRange.Comment.Author);
			Assert.AreEqual("Testing comment 1", sourceExcelRange.Comment.Text);
			// Assert the comment was copied.
			Assert.AreEqual("test1", destinationExcelRange.Comment.Author);
			Assert.AreEqual("Testing comment 1", destinationExcelRange.Comment.Text);
		}

		[TestMethod]
		public void CopyCopiesCommentsFromMultiCellRanges()
		{
			InitBase();
			var pck = new ExcelPackage();
			var ws1 = pck.Workbook.Worksheets.Add("CommentCopying");
			var sourceExcelRangeC3 = ws1.Cells[3, 3];
			var sourceExcelRangeD3 = ws1.Cells[3, 4];
			var sourceExcelRangeE3 = ws1.Cells[3, 5];
			Assert.IsNull(sourceExcelRangeC3.Comment);
			Assert.IsNull(sourceExcelRangeD3.Comment);
			Assert.IsNull(sourceExcelRangeE3.Comment);
			sourceExcelRangeC3.AddComment("Testing comment 1", "test1");
			sourceExcelRangeD3.AddComment("Testing comment 2", "test1");
			sourceExcelRangeE3.AddComment("Testing comment 3", "test1");
			Assert.AreEqual("test1", sourceExcelRangeC3.Comment.Author);
			Assert.AreEqual("Testing comment 1", sourceExcelRangeC3.Comment.Text);
			Assert.AreEqual("test1", sourceExcelRangeD3.Comment.Author);
			Assert.AreEqual("Testing comment 2", sourceExcelRangeD3.Comment.Text);
			Assert.AreEqual("test1", sourceExcelRangeE3.Comment.Author);
			Assert.AreEqual("Testing comment 3", sourceExcelRangeE3.Comment.Text);
			// Copy the full row to capture each cell at once.
			Assert.IsNull(ws1.Cells[5, 3].Comment);
			Assert.IsNull(ws1.Cells[5, 4].Comment);
			Assert.IsNull(ws1.Cells[5, 5].Comment);
			ws1.Cells["3:3"].Copy(ws1.Cells["5:5"]);
			// Assert the original comments are intact.
			Assert.AreEqual("test1", sourceExcelRangeC3.Comment.Author);
			Assert.AreEqual("Testing comment 1", sourceExcelRangeC3.Comment.Text);
			Assert.AreEqual("test1", sourceExcelRangeD3.Comment.Author);
			Assert.AreEqual("Testing comment 2", sourceExcelRangeD3.Comment.Text);
			Assert.AreEqual("test1", sourceExcelRangeE3.Comment.Author);
			Assert.AreEqual("Testing comment 3", sourceExcelRangeE3.Comment.Text);
			// Assert the comments were copied.
			var destinationExcelRangeC5 = ws1.Cells[5, 3];
			var destinationExcelRangeD5 = ws1.Cells[5, 4];
			var destinationExcelRangeE5 = ws1.Cells[5, 5];
			Assert.AreEqual("test1", destinationExcelRangeC5.Comment.Author);
			Assert.AreEqual("Testing comment 1", destinationExcelRangeC5.Comment.Text);
			Assert.AreEqual("test1", destinationExcelRangeD5.Comment.Author);
			Assert.AreEqual("Testing comment 2", destinationExcelRangeD5.Comment.Text);
			Assert.AreEqual("test1", destinationExcelRangeE5.Comment.Author);
			Assert.AreEqual("Testing comment 3", destinationExcelRangeE5.Comment.Text);
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\Sparkline Demos.xlsx")]
		public void CopySparklinesCopiesToSameSheet()
		{
			var file = new FileInfo("Sparkline Demos.xlsx");
			Assert.IsTrue(file.Exists);
			var temp = Path.GetTempFileName();
			File.Delete(temp);
			var copy = file.CopyTo(temp);
			try
			{
				using (var package = new ExcelPackage(copy))
				{
					var sheet1 = package.Workbook.Worksheets["Sheet1"];
					var sparklines = sheet1.SparklineGroups.SparklineGroups;
					Assert.AreEqual(7, sparklines.Count);
					sheet1.Cells[9, 3, 12, 4].Copy(new ExcelRange(sheet1, 13, 3, 16, 4));
					Assert.AreEqual(9, sparklines.Count);
					Assert.AreEqual("'Sheet2'!B6:I6", sparklines[8].Sparklines[0].Formula.Address);
					Assert.AreEqual("C16", sparklines[8].Sparklines[0].HostCell.Address);
					Assert.AreEqual("'Sheet1'!D10:D12", sparklines[7].Sparklines[0].Formula.Address);
					Assert.AreEqual("D13", sparklines[7].Sparklines[0].HostCell.Address);
					Assert.AreEqual("Sheet2!B2:I2", sparklines[6].Sparklines[0].Formula.Address);
					Assert.AreEqual("C12", sparklines[6].Sparklines[0].HostCell.Address);
					Assert.AreEqual("Sheet1!D6:F6", sparklines[5].Sparklines[0].Formula.Address);
					Assert.AreEqual("G6", sparklines[5].Sparklines[0].HostCell.Address);
					Assert.AreEqual("Sheet1!D7:F7", sparklines[4].Sparklines[0].Formula.Address);
					Assert.AreEqual("G7", sparklines[4].Sparklines[0].HostCell.Address);
					Assert.AreEqual("Sheet1!D8:F8", sparklines[3].Sparklines[0].Formula.Address);
					Assert.AreEqual("G8", sparklines[3].Sparklines[0].HostCell.Address);
					Assert.AreEqual("Sheet1!D6:D8", sparklines[2].Sparklines[0].Formula.Address);
					Assert.AreEqual("D9", sparklines[2].Sparklines[0].HostCell.Address);
					Assert.AreEqual("Sheet1!E6:E8", sparklines[1].Sparklines[0].Formula.Address);
					Assert.AreEqual("E9", sparklines[1].Sparklines[0].HostCell.Address);
					Assert.AreEqual("Sheet1!F6:F8", sparklines[0].Sparklines[0].Formula.Address);
					Assert.AreEqual("F9", sparklines[0].Sparklines[0].HostCell.Address);
					package.Save();
				}
				using (var package = new ExcelPackage(copy))
				{
					var sheet = package.Workbook.Worksheets["Sheet1"];
					var sparklines = sheet.SparklineGroups.SparklineGroups;
					Assert.AreEqual(9, sparklines.Count);
					Assert.AreEqual("'Sheet2'!B6:I6", sparklines[8].Sparklines[0].Formula.Address);
					Assert.AreEqual("C16", sparklines[8].Sparklines[0].HostCell.Address);
					Assert.AreEqual("'Sheet1'!D10:D12", sparklines[7].Sparklines[0].Formula.Address);
					Assert.AreEqual("D13", sparklines[7].Sparklines[0].HostCell.Address);
					Assert.AreEqual("Sheet2!B2:I2", sparklines[6].Sparklines[0].Formula.Address);
					Assert.AreEqual("C12", sparklines[6].Sparklines[0].HostCell.Address);
					Assert.AreEqual("Sheet1!D6:F6", sparklines[5].Sparklines[0].Formula.Address);
					Assert.AreEqual("G6", sparklines[5].Sparklines[0].HostCell.Address);
					Assert.AreEqual("Sheet1!D7:F7", sparklines[4].Sparklines[0].Formula.Address);
					Assert.AreEqual("G7", sparklines[4].Sparklines[0].HostCell.Address);
					Assert.AreEqual("Sheet1!D8:F8", sparklines[3].Sparklines[0].Formula.Address);
					Assert.AreEqual("G8", sparklines[3].Sparklines[0].HostCell.Address);
					Assert.AreEqual("Sheet1!D6:D8", sparklines[2].Sparklines[0].Formula.Address);
					Assert.AreEqual("D9", sparklines[2].Sparklines[0].HostCell.Address);
					Assert.AreEqual("Sheet1!E6:E8", sparklines[1].Sparklines[0].Formula.Address);
					Assert.AreEqual("E9", sparklines[1].Sparklines[0].HostCell.Address);
					Assert.AreEqual("Sheet1!F6:F8", sparklines[0].Sparklines[0].Formula.Address);
					Assert.AreEqual("F9", sparklines[0].Sparklines[0].HostCell.Address);
				}
			}
			finally
			{
				copy.Delete();
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\Sparkline Demos.xlsx")]
		public void CopySparklinesCopiesToDifferentSheet()
		{
			var file = new FileInfo("Sparkline Demos.xlsx");
			Assert.IsTrue(file.Exists);
			var temp = Path.GetTempFileName();
			File.Delete(temp);
			var copy = file.CopyTo(temp);
			try
			{
				string newSheetName = "Sheet3";
				using (var package = new ExcelPackage(copy))
				{
					var sheet1 = package.Workbook.Worksheets["Sheet1"];
					var sheet3 = package.Workbook.Worksheets.Add(newSheetName);
					var sheet1Sparklines = sheet1.SparklineGroups.SparklineGroups;
					Assert.AreEqual(7, sheet1Sparklines.Count);
					sheet1.Cells[9, 3, 12, 4].Copy(new ExcelRange(sheet3, 13, 3, 16, 4));
					Assert.AreEqual(7, sheet1Sparklines.Count);
					Assert.AreEqual("Sheet2!B2:I2", sheet1Sparklines[6].Sparklines[0].Formula.Address);
					Assert.AreEqual("C12", sheet1Sparklines[6].Sparklines[0].HostCell.Address);
					Assert.AreEqual("Sheet1!D6:F6", sheet1Sparklines[5].Sparklines[0].Formula.Address);
					Assert.AreEqual("G6", sheet1Sparklines[5].Sparklines[0].HostCell.Address);
					Assert.AreEqual("Sheet1!D7:F7", sheet1Sparklines[4].Sparklines[0].Formula.Address);
					Assert.AreEqual("G7", sheet1Sparklines[4].Sparklines[0].HostCell.Address);
					Assert.AreEqual("Sheet1!D8:F8", sheet1Sparklines[3].Sparklines[0].Formula.Address);
					Assert.AreEqual("G8", sheet1Sparklines[3].Sparklines[0].HostCell.Address);
					Assert.AreEqual("Sheet1!D6:D8", sheet1Sparklines[2].Sparklines[0].Formula.Address);
					Assert.AreEqual("D9", sheet1Sparklines[2].Sparklines[0].HostCell.Address);
					Assert.AreEqual("Sheet1!E6:E8", sheet1Sparklines[1].Sparklines[0].Formula.Address);
					Assert.AreEqual("E9", sheet1Sparklines[1].Sparklines[0].HostCell.Address);
					Assert.AreEqual("Sheet1!F6:F8", sheet1Sparklines[0].Sparklines[0].Formula.Address);
					Assert.AreEqual("F9", sheet1Sparklines[0].Sparklines[0].HostCell.Address);
					var sheet3Sparklines = sheet3.SparklineGroups.SparklineGroups;
					Assert.AreEqual(2, sheet3Sparklines.Count);
					Assert.AreEqual($"'{newSheetName}'!D10:D12", sheet3Sparklines[0].Sparklines[0].Formula.Address);
					Assert.AreEqual("D13", sheet3Sparklines[0].Sparklines[0].HostCell.Address);
					Assert.AreEqual("'Sheet2'!B6:I6", sheet3Sparklines[1].Sparklines[0].Formula.Address);
					Assert.AreEqual("C16", sheet3Sparklines[1].Sparklines[0].HostCell.Address);
					package.Save();
				}
				using (var package = new ExcelPackage(copy))
				{
					var sheet1 = package.Workbook.Worksheets["Sheet1"];
					var sheet3 = package.Workbook.Worksheets[newSheetName];
					var sheet1Sparklines = sheet1.SparklineGroups.SparklineGroups;
					Assert.AreEqual(7, sheet1Sparklines.Count);
					Assert.AreEqual("Sheet2!B2:I2", sheet1Sparklines[6].Sparklines[0].Formula.Address);
					Assert.AreEqual("C12", sheet1Sparklines[6].Sparklines[0].HostCell.Address);
					Assert.AreEqual("Sheet1!D6:F6", sheet1Sparklines[5].Sparklines[0].Formula.Address);
					Assert.AreEqual("G6", sheet1Sparklines[5].Sparklines[0].HostCell.Address);
					Assert.AreEqual("Sheet1!D7:F7", sheet1Sparklines[4].Sparklines[0].Formula.Address);
					Assert.AreEqual("G7", sheet1Sparklines[4].Sparklines[0].HostCell.Address);
					Assert.AreEqual("Sheet1!D8:F8", sheet1Sparklines[3].Sparklines[0].Formula.Address);
					Assert.AreEqual("G8", sheet1Sparklines[3].Sparklines[0].HostCell.Address);
					Assert.AreEqual("Sheet1!D6:D8", sheet1Sparklines[2].Sparklines[0].Formula.Address);
					Assert.AreEqual("D9", sheet1Sparklines[2].Sparklines[0].HostCell.Address);
					Assert.AreEqual("Sheet1!E6:E8", sheet1Sparklines[1].Sparklines[0].Formula.Address);
					Assert.AreEqual("E9", sheet1Sparklines[1].Sparklines[0].HostCell.Address);
					Assert.AreEqual("Sheet1!F6:F8", sheet1Sparklines[0].Sparklines[0].Formula.Address);
					Assert.AreEqual("F9", sheet1Sparklines[0].Sparklines[0].HostCell.Address);
					var sheet3Sparklines = sheet3.SparklineGroups.SparklineGroups;
					Assert.AreEqual(2, sheet3Sparklines.Count);
					Assert.AreEqual($"'{newSheetName}'!D10:D12", sheet3Sparklines[0].Sparklines[0].Formula.Address);
					Assert.AreEqual("D13", sheet3Sparklines[0].Sparklines[0].HostCell.Address);
					Assert.AreEqual("'Sheet2'!B6:I6", sheet3Sparklines[1].Sparklines[0].Formula.Address);
					Assert.AreEqual("C16", sheet3Sparklines[1].Sparklines[0].HostCell.Address);
				}
			}
			finally
			{
				copy.Delete();
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\Sparkline Demos.xlsx")]
		public void CopySparklinesCopiesToDifferentSheetBadReference()
		{
			var file = new FileInfo("Sparkline Demos.xlsx");
			Assert.IsTrue(file.Exists);
			var temp = Path.GetTempFileName();
			File.Delete(temp);
			var copy = file.CopyTo(temp);
			try
			{
				string newSheetName = "Sheet3";
				using (var package = new ExcelPackage(copy))
				{
					var sheet1 = package.Workbook.Worksheets["Sheet1"];
					var sheet3 = package.Workbook.Worksheets.Add(newSheetName);
					var sheet1Sparklines = sheet1.SparklineGroups.SparklineGroups;
					Assert.AreEqual(7, sheet1Sparklines.Count);
					sheet1.Cells["C12"].Copy(new ExcelRange(sheet3, 2, 2, 2, 2));
					Assert.AreEqual(7, sheet1Sparklines.Count);
					Assert.AreEqual("Sheet2!B2:I2", sheet1Sparklines[6].Sparklines[0].Formula.Address);
					Assert.AreEqual("Sheet1!D6:F6", sheet1Sparklines[5].Sparklines[0].Formula.Address);
					Assert.AreEqual("Sheet1!D7:F7", sheet1Sparklines[4].Sparklines[0].Formula.Address);
					Assert.AreEqual("Sheet1!D8:F8", sheet1Sparklines[3].Sparklines[0].Formula.Address);
					Assert.AreEqual("Sheet1!D6:D8", sheet1Sparklines[2].Sparklines[0].Formula.Address);
					Assert.AreEqual("Sheet1!E6:E8", sheet1Sparklines[1].Sparklines[0].Formula.Address);
					Assert.AreEqual("Sheet1!F6:F8", sheet1Sparklines[0].Sparklines[0].Formula.Address);
					var sheet3Sparklines = sheet3.SparklineGroups.SparklineGroups;
					Assert.AreEqual(1, sheet3Sparklines.Count);
					Assert.IsNull(sheet3Sparklines[0].Sparklines[0].Formula);
					Assert.AreEqual("B2", sheet3Sparklines[0].Sparklines[0].HostCell.Address);
					package.Save();
				}
				using (var package = new ExcelPackage(copy))
				{
					var sheet1 = package.Workbook.Worksheets["Sheet1"];
					var sheet3 = package.Workbook.Worksheets[newSheetName];
					var sparklines = sheet1.SparklineGroups.SparklineGroups;
					Assert.AreEqual(7, sparklines.Count);
					Assert.AreEqual("Sheet2!B2:I2", sparklines[6].Sparklines[0].Formula.Address);
					Assert.AreEqual("Sheet1!D6:F6", sparklines[5].Sparklines[0].Formula.Address);
					Assert.AreEqual("Sheet1!D7:F7", sparklines[4].Sparklines[0].Formula.Address);
					Assert.AreEqual("Sheet1!D8:F8", sparklines[3].Sparklines[0].Formula.Address);
					Assert.AreEqual("Sheet1!D6:D8", sparklines[2].Sparklines[0].Formula.Address);
					Assert.AreEqual("Sheet1!E6:E8", sparklines[1].Sparklines[0].Formula.Address);
					Assert.AreEqual("Sheet1!F6:F8", sparklines[0].Sparklines[0].Formula.Address);
					var sheet3Sparklines = sheet3.SparklineGroups.SparklineGroups;
					Assert.AreEqual(1, sheet3Sparklines.Count);
					Assert.AreEqual("B2", sheet3Sparklines[0].Sparklines[0].HostCell.Address);
					Assert.IsNull(sheet3Sparklines[0].Sparklines[0].Formula);
				}
			}
			finally
			{
				copy.Delete();
			}
		}

		[TestMethod]
		public void CopySparklineWithNoFormula()
		{
			var tempWorkbook = new FileInfo(Path.GetTempFileName());
			try
			{
				using (var package = new ExcelPackage())
				{
					var worksheet = package.Workbook.Worksheets.Add("Sheet");
					var worksheetSparklineGroups = worksheet.SparklineGroups.SparklineGroups;

					var sparklineGroup = new OfficeOpenXml.Drawing.Sparkline.ExcelSparklineGroup(worksheet, worksheet.NameSpaceManager);
					// Use NULL for Formula
					var sparkline = new OfficeOpenXml.Drawing.Sparkline.ExcelSparkline(new ExcelAddress("C3"), null, sparklineGroup, worksheet.NameSpaceManager);
					sparklineGroup.Sparklines.Add(sparkline);
					worksheet.SparklineGroups.SparklineGroups.Add(sparklineGroup);
					package.SaveAs(tempWorkbook);
				}
				using (var package = new ExcelPackage(tempWorkbook))
				{
					var worksheet = package.Workbook.Worksheets["Sheet"];
					var sparklineGroups = worksheet.SparklineGroups;
					Assert.AreEqual(1, sparklineGroups.SparklineGroups.Count);
					Assert.AreEqual(1, sparklineGroups.SparklineGroups[0].Sparklines.Count);
					var origin = worksheet.Cells["C3"];
					var target = worksheet.Cells["E5"];
					origin.Copy(target);
					Assert.AreEqual(2, sparklineGroups.SparklineGroups.Count);
					Assert.AreEqual(1, sparklineGroups.SparklineGroups[0].Sparklines.Count);
					Assert.AreEqual(1, sparklineGroups.SparklineGroups[1].Sparklines.Count);
				}
			}
			finally
			{
				tempWorkbook.Delete();
			}
		}
		#endregion

		#region Shared Formula Overwrite Tests
		[TestMethod]
		[DeploymentItem(@"Workbooks\SharedFormulasRows.xlsx")]
		public void OverwrittenSharedFormulaRowsAreRespected()
		{
			// In Excel, a cell in a shared formula range can have its formula replaced by something
			// that is not the shared formula. EP Plus overwrites this formula when splitting shared
			// formulas.
			FileInfo newFile = new FileInfo(Path.GetTempFileName());
			if (newFile.Exists)
				newFile.Delete();
			var dir = AppDomain.CurrentDomain.BaseDirectory;
			var file = new FileInfo("SharedFormulasRows.xlsx");
			Assert.IsTrue(file.Exists);
			try
			{
				using (var pck = new ExcelPackage(file))
				{
					var sheet = pck.Workbook.Worksheets["Sheet1"];
					// This formula is in the shared formula range, but was explicitly overwritten in Excel.
					var nonSharedFormulaOriginal = sheet.Cells["C5"].Formula;
					// Set some other cell's formula in the shared formula range to trigger a split.
					sheet.Cells["C4"].Formula = "SUM(1,2)";
					// Verify that the explicit formula in the shared formula range was NOT overwritten.
					var nonSharedFormulaUpdated = sheet.Cells["C5"].Formula;
					Assert.AreEqual(nonSharedFormulaOriginal, nonSharedFormulaUpdated);
					// Ensure that the shared formula was propagated to the region.
					var sharedFormulaUpdated = sheet.Cells["C6"].Formula;
					Assert.AreEqual("B6+D6", sharedFormulaUpdated);
					pck.SaveAs(newFile);
				}
				// Ensure the integrity of the package is maintained.
				using (var pck = new ExcelPackage(newFile))
				{
					var sheet = pck.Workbook.Worksheets["Sheet1"];
					Assert.AreEqual("B3+D3", sheet.Cells["C3"].Formula);
					Assert.AreEqual("SUM(1,2)", sheet.Cells["C4"].Formula);
					Assert.AreEqual("B5*D5", sheet.Cells["C5"].Formula);
					Assert.AreEqual("B6+D6", sheet.Cells["C6"].Formula);
				}
			}
			finally
			{
				if (newFile.Exists)
					newFile.Delete();
			}
		}

		[TestMethod]
		[DeploymentItem(@"Workbooks\SharedFormulasColumns.xlsx")]
		public void OverwrittenSharedFormulaColumnsAreRespected()
		{
			// In Excel, a cell in a shared formula range can have its formula replaced by something
			// that is not the shared formula. EP Plus overwrites this formula when splitting shared
			// formulas.
			FileInfo newFile = new FileInfo(Path.GetTempFileName());
			if (newFile.Exists)
				newFile.Delete();
			var dir = AppDomain.CurrentDomain.BaseDirectory;
			var file = new FileInfo("SharedFormulasColumns.xlsx");
			Assert.IsTrue(file.Exists);
			try
			{
				using (var pck = new ExcelPackage(file))
				{
					var sheet = pck.Workbook.Worksheets["Sheet1"];
					// This formula is in the shared formula range, but was explicitly overwritten in Excel.
					var nonSharedFormulaOriginal = sheet.Cells["E3"].Formula;
					// Set some other cell's formula in the shared formula range to trigger a split.
					sheet.Cells["D3"].Formula = "SUM(1,2)";
					// Verify that the explicit formula in the shared formula range was NOT overwritten.
					var nonSharedFormulaUpdated = sheet.Cells["E3"].Formula;
					Assert.AreEqual(nonSharedFormulaOriginal, nonSharedFormulaUpdated);
					// Ensure that the shared formula was propagated to the region.
					var sharedFormulaUpdated = sheet.Cells["F3"].Formula;
					Assert.AreEqual("F2+F4", sharedFormulaUpdated);
					pck.SaveAs(newFile);
				}
				// Ensure the integrity of the package is maintained.
				using (var pck = new ExcelPackage(newFile))
				{
					var sheet = pck.Workbook.Worksheets["Sheet1"];
					Assert.AreEqual("C2+C4", sheet.Cells["C3"].Formula);
					Assert.AreEqual("SUM(1,2)", sheet.Cells["D3"].Formula);
					Assert.AreEqual("E2*E4", sheet.Cells["E3"].Formula);
					Assert.AreEqual("F2+F4", sheet.Cells["F3"].Formula);
				}
			}
			finally
			{
				if (newFile.Exists)
					newFile.Delete();
			}
		}

		[TestMethod]
		public void FullyOverwrittenSharedRowFormulaIsRemovedCorrectly()
		{
			// There is a bug where Shared Formulas that are overwritten from the bottom up cause
			// an exception on the final cell in the Shared Formula range.
			FileInfo newFile = new FileInfo(Path.GetTempFileName());
			if (newFile.Exists)
				newFile.Delete();
			try
			{
				var dir = AppDomain.CurrentDomain.BaseDirectory;
				using (var pck = new ExcelPackage())
				{
					var sheet = pck.Workbook.Worksheets.Add("Sheet1");
					sheet.Cells["C3:C4"].Formula = "B3+B3";
					Assert.AreEqual("B3+B3", sheet.Cells[3, 3].Formula);
					Assert.AreEqual("B4+B4", sheet.Cells[4, 3].Formula);
					sheet.Cells[4, 3].Formula = "B4-B4";
					Assert.AreEqual("B3+B3", sheet.Cells[3, 3].Formula);
					Assert.AreEqual("B4-B4", sheet.Cells[4, 3].Formula);
					sheet.Cells[3, 3].Formula = "B3-B3";
					Assert.AreEqual("B3-B3", sheet.Cells[3, 3].Formula);
					Assert.AreEqual("B4-B4", sheet.Cells[4, 3].Formula);
					pck.SaveAs(newFile);
				}
				// Ensure the integrity of the package is maintained.
				using (var pck = new ExcelPackage(newFile))
				{
					var sheet = pck.Workbook.Worksheets["Sheet1"];
					Assert.AreEqual("B3-B3", sheet.Cells[3, 3].Formula);
					Assert.AreEqual("B4-B4", sheet.Cells[4, 3].Formula);
				}
			}
			finally
			{
				if (newFile.Exists)
					newFile.Delete();
			}
		}

		[TestMethod]
		public void FullyOverwrittenSharedColumnFormulaIsRemovedCorrectly()
		{
			// There is a bug where Shared Formulas that are overwritten from the bottom up cause
			// an exception on the final cell in the Shared Formula range.
			FileInfo newFile = new FileInfo(Path.GetTempFileName());
			if (newFile.Exists)
				newFile.Delete();
			try
			{
				var dir = AppDomain.CurrentDomain.BaseDirectory;
				using (var pck = new ExcelPackage())
				{
					var sheet = pck.Workbook.Worksheets.Add("Sheet1");
					sheet.Cells["C3:D3"].Formula = "C4+C4";
					Assert.AreEqual("C4+C4", sheet.Cells[3, 3].Formula);
					Assert.AreEqual("D4+D4", sheet.Cells[3, 4].Formula);
					sheet.Cells[3, 4].Formula = "D4-D4";
					Assert.AreEqual("C4+C4", sheet.Cells[3, 3].Formula);
					Assert.AreEqual("D4-D4", sheet.Cells[3, 4].Formula);
					sheet.Cells[3, 3].Formula = "C4-C4";
					Assert.AreEqual("C4-C4", sheet.Cells[3, 3].Formula);
					Assert.AreEqual("D4-D4", sheet.Cells[3, 4].Formula);
					pck.SaveAs(newFile);
				}
				// Ensure the integrity of the package is maintained.
				using (var pck = new ExcelPackage(newFile))
				{
					var sheet = pck.Workbook.Worksheets["Sheet1"];
					Assert.AreEqual("C4-C4", sheet.Cells[3, 3].Formula);
					Assert.AreEqual("D4-D4", sheet.Cells[3, 4].Formula);
				}
			}
			finally
			{
				if (newFile.Exists)
					newFile.Delete();
			}
		}
		#endregion

		#region SetFormula Tests
		[TestMethod]
		public void SetFormulaRemovesLeadingEquals()
		{
			using (var pkg = new ExcelPackage())
			{
				var sheet = pkg.Workbook.Worksheets.Add("Sheet");
				sheet.Cells[3, 3].Formula = "=SUM(1,2)";
				Assert.AreEqual("SUM(1,2)", sheet.Cells[3, 3].Formula);
			}
		}

		[TestMethod]
		public void SetFormulaDoesNotOverwriteValue()
		{
			using (var pkg = new ExcelPackage())
			{
				var sheet = pkg.Workbook.Worksheets.Add("Sheet");
				string expectedValue = "some value";
				string expectedFormula = "SUM(1, 2)";
				sheet.Cells[3, 3].Value = expectedValue;
				sheet.Cells[3, 3].SetFormula($"={expectedFormula}", false);
				Assert.AreEqual("some value", sheet.Cells[3, 3].Value);
				Assert.AreEqual(expectedFormula, sheet.Cells[3, 3].Formula);
			}
		}

		[TestMethod]
		public void SetFormulaOverwritesValue()
		{
			using (var pkg = new ExcelPackage())
			{
				var sheet = pkg.Workbook.Worksheets.Add("Sheet");
				string expectedFormula = "SUM(1, 2)";
				sheet.Cells[3, 3].Value = "some value";
				sheet.Cells[3, 3].SetFormula($"={expectedFormula}", true);
				Assert.IsNull(sheet.Cells[3, 3].Value);
				Assert.AreEqual(expectedFormula, sheet.Cells[3, 3].Formula);
			}
		}

		[TestMethod]
		public void SetNullOrEmptyFormulaOverwritesValue()
		{
			using (var pkg = new ExcelPackage())
			{
				var sheet = pkg.Workbook.Worksheets.Add("Sheet");
				sheet.Cells[3, 3].Value = "some value";
				sheet.Cells[3, 3].SetFormula(null, true);
				Assert.IsNull(sheet.Cells[3, 3].Value);
				Assert.AreEqual(string.Empty, sheet.Cells[3, 3].Formula);

				sheet.Cells[3, 3].Value = "some value";
				sheet.Cells[3, 3].SetFormula(string.Empty, true);
				Assert.IsNull(sheet.Cells[3, 3].Value);
				Assert.AreEqual(string.Empty, sheet.Cells[3, 3].Formula);
			}
		}

		[TestMethod]
		public void SetNullOrEmptyFormulaDoesNotOverwriteValue()
		{
			using (var pkg = new ExcelPackage())
			{
				var sheet = pkg.Workbook.Worksheets.Add("Sheet");
				string expectedValue = "some value";
				sheet.Cells[3, 3].Value = expectedValue;
				sheet.Cells[3, 3].SetFormula(null, false);
				Assert.AreEqual(expectedValue, sheet.Cells[3, 3].Value);
				Assert.AreEqual(string.Empty, sheet.Cells[3, 3].Formula);

				sheet.Cells[3, 3].Value = expectedValue;
				sheet.Cells[3, 3].SetFormula(string.Empty, false);
				Assert.AreEqual(expectedValue, sheet.Cells[3, 3].Value);
				Assert.AreEqual(string.Empty, sheet.Cells[3, 3].Formula);
			}
		}

		[TestMethod]
		public void SetFormulaOnEmptyCellCreatesEmptyValue()
		{
			const string formula = "1 + 1";
			var file = new FileInfo(Path.GetTempFileName());
			if (file.Exists)
				file.Delete();
			try
			{
				using (var package = new ExcelPackage())
				{
					var sheet = package.Workbook.Worksheets.Add("Sheet1");
					var cell = package.Workbook.Worksheets["Sheet1"].Cells[2, 2];
					cell.SetFormula(formula, false);
					package.SaveAs(file);
				}
				using (var package = new ExcelPackage(file))
				{
					var sheet = package.Workbook.Worksheets["Sheet1"];
					Assert.AreEqual(formula, package.Workbook.Worksheets["Sheet1"].Cells[2, 2].Formula);
				}
			}
			finally
			{
				if (file.Exists)
					file.Delete();
			}
		}
		#endregion

		#region SetAddress Tests
		[TestMethod]
		public void SetAddress()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet1");
				var sheet2 = package.Workbook.Worksheets.Add("Sheet2");
				// It probably shouldn't be allowed to change the address of a cell, but since
				// it is the backing worksheet reference should be set correctly.
				var cell = sheet.Cells[3, 3];
				cell.Address = "'Sheet2'!D4";
				Assert.AreEqual(sheet2.Name, cell.Worksheet.Name);
				Assert.AreNotEqual(sheet2.Name, sheet.Cells[3, 3].Worksheet.Name);
			}
		}
		#endregion

		#region SetSharedFormula Tests
		[TestMethod]
		public void SetSharedFormulaRemovesLeadingEquals()
		{
			using (var pkg = new ExcelPackage())
			{
				var sheet = pkg.Workbook.Worksheets.Add("Sheet");
				sheet.Cells[3, 3, 5, 5].Formula = "=SUM(1,2)";
				Assert.AreEqual("SUM(1,2)", sheet.Cells[3, 3].Formula);
			}
		}
		#endregion

		#region Equals Tests with array values
		[TestMethod]
		public void ArrayEquality()
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells[3, 3].Formula = "{\"test\"}=\"test\"";
				worksheet.Cells[3, 4].Formula = "{\"test\"}=\"\"\"test\"\"\"";
				worksheet.Cells[3, 5].Formula = "{\"test1\",\"test2\"}=\"test1\"";
				worksheet.Cells[3, 6].Formula = "{\"test1\",\"test2\"}={\"test1\"}";
				worksheet.Cells[3, 7].Formula = "{\"test1\",\"test2\"}={\"test1\",\"testB\"}";
				worksheet.Cells[3, 8].Formula = "{1,2,3}={1,2,3}";
				worksheet.Cells[3, 9].Formula = "{1,2,3}={1}";
				worksheet.Cells[3, 10].Formula = "{1,2,3}+4";
				worksheet.Calculate();
				Assert.IsTrue((bool)worksheet.Cells[3, 3].Value);
				Assert.IsFalse((bool)worksheet.Cells[3, 4].Value);
				Assert.IsTrue((bool)worksheet.Cells[3, 5].Value);
				Assert.IsTrue((bool)worksheet.Cells[3, 6].Value);
				Assert.IsTrue((bool)worksheet.Cells[3, 7].Value);
				Assert.IsTrue((bool)worksheet.Cells[3, 7].Value);
				Assert.IsTrue((bool)worksheet.Cells[3, 8].Value);
				Assert.IsTrue((bool)worksheet.Cells[3, 9].Value);
				Assert.AreEqual(5d, worksheet.Cells[3, 10].Value);
			}
		}

		[TestMethod]
		public void ArrayCell()
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells[3, 3].Formula = "{\"test1\",\"test2\"}";
				worksheet.Calculate();
				Assert.AreEqual("test1", worksheet.Cells[3, 3].Value);
			}
		}

		[TestMethod]
		public void IfWithArray()
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells[3, 3].Formula = "IF(FALSE,\"true\",{\"false\"})";
				worksheet.Cells[3, 3].Calculate();
				Assert.AreEqual("false", worksheet.Cells[3, 3].Value);
			}
		}
		#endregion

		#region AutoFit Tests
		[TestMethod]
		public void AutoFitResizesColumnToFitContentsWithDefaultRowHeight()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				sheet.Cells[1, 2].Value = "Header to set width";
				sheet.Cells[1, 2].Style.WrapText = true;
				var contents = "=0";
				sheet.Cells[2, 2].Formula = contents;
				sheet.Cells[2, 2].Style.Numberformat.Format = "_(* #,##0_);_(* (#,##0);_(* \" - \"??_);_(@_)";
				sheet.Column(2).Width = 2;
				Assert.AreEqual(2, sheet.Column(2).Width);
				sheet.Cells[2, 3].Formula = "=\"Next cell contents.\"";
				sheet.Cells[2, 2].Calculate();
				sheet.Column(2).AutoFit();
				var actualWidth = sheet.Column(2).Width;
				Assert.AreEqual(21, Math.Round(sheet.Column(2).Width));
			}
		}

		[TestMethod]
		public void AutoFitResizesColumnToFitContentsWithNoSpacesRegardlessOfRowHeight()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				sheet.Row(1).Height = 33; // Just a bit taller than two lines of text.
				sheet.Cells[1, 2].Value = "HeaderWithNoSpacesShouldNotBreak";
				sheet.Cells[1, 2].Style.WrapText = true;
				var contents = "=0";
				sheet.Cells[2, 2].Formula = contents;
				sheet.Cells[2, 2].Style.Numberformat.Format = "_(* #,##0_);_(* (#,##0);_(* \" - \"??_);_(@_)";
				sheet.Column(2).Width = 2;
				Assert.AreEqual(2, sheet.Column(2).Width);
				sheet.Cells[2, 3].Formula = "=\"Next cell contents.\"";
				sheet.Cells[2, 2].Calculate();
				sheet.Column(2).AutoFit();
				Assert.AreEqual(35, Math.Round(sheet.Column(2).Width));
			}
		}

		[TestMethod]
		public void AutoFitResizesColumnToFitContentsWithSpecifiedRowHeight()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				sheet.Row(1).Height = 33; // Just a bit taller than two lines of text.
				sheet.Cells[1, 2].Value = "Header to set width";
				sheet.Cells[1, 2].Style.WrapText = true;
				var contents = "=0";
				sheet.Cells[2, 2].Formula = contents;
				sheet.Cells[2, 2].Style.Numberformat.Format = "_(* #,##0_);_(* (#,##0);_(* \" - \"??_);_(@_)";
				sheet.Column(2).Width = 2;
				Assert.AreEqual(2, sheet.Column(2).Width);
				sheet.Cells[2, 3].Formula = "=\"Next cell contents.\"";
				sheet.Cells[2, 2].Calculate();
				sheet.Column(2).AutoFit();
				Assert.AreEqual(10, Math.Round(sheet.Column(2).Width));
			}
		}

		[TestMethod]
		public void AutoFitResizesColumnToFitContentsWithSpecifiedRowHeightForFiveLines()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				sheet.Row(1).Height = 76; // Just a bit taller than five lines of text.
				sheet.Cells[1, 2].Value = "Set width to five lines";
				sheet.Cells[1, 2].Style.WrapText = true;
				var contents = "=0";
				sheet.Cells[2, 2].Formula = contents;
				sheet.Cells[2, 2].Style.Numberformat.Format = "_(* #,##0_);_(* (#,##0);_(* \" - \"??_);_(@_)";
				sheet.Column(2).Width = 2;
				Assert.AreEqual(2, sheet.Column(2).Width);
				sheet.Cells[2, 3].Formula = "=\"Next cell contents.\"";
				sheet.Cells[2, 2].Calculate();
				sheet.Column(2).AutoFit();
				Assert.AreEqual(10, Math.Round(sheet.Column(2).Width));
			}
		}

		[TestMethod]
		public void AutoFitAccountsForConditionalFormatting()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				sheet.Cells[1, 2].Value = "iiiiiiiiiiiiiiiiiiii";
				sheet.Column(2).AutoFit();
				var normalWidth = sheet.Column(2).Width;
				var conditionalFormattingRule = sheet.Cells[1, 2].ConditionalFormatting.AddEqual();
				// The formula doesn't actually have to match, as auto-fitting will take the largest
				// possible font.
				conditionalFormattingRule.Formula = "iiiiiiiiiiiiiiiiiiii";
				conditionalFormattingRule.Style.Font.Bold = true;
				sheet.Column(2).AutoFit();
				var boldWidth = sheet.Column(2).Width;
				Assert.IsTrue(boldWidth > normalWidth);
			}
		}

		[TestMethod]
		public void AutoFitHandlesRotatedContents()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				sheet.Cells[1, 2].Value = "iiiiiiiiiiiiiiiiiiii";
				sheet.Column(2).AutoFit();
				var normalWidth = sheet.Column(2).Width;
				sheet.Cells[1, 2].Style.TextRotation = 30;
				sheet.Column(2).AutoFit();
				var rotatedWidth = sheet.Column(2).Width;
				Assert.IsTrue(rotatedWidth < normalWidth);
			}
		}
		#endregion

		#region IsEquivalentRange Tests
		[TestMethod]
		public void IsEquivalentRangeAreEqual()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("sheet1");
				var range1 = new ExcelRangeBase(worksheet, "A1:F6");
				var range2 = new ExcelRangeBase(worksheet, "A1:F6");
				Assert.IsTrue(range1.IsEquivalentRange(range2));
			}
		}

		[TestMethod]
		public void IsEquivalentRangeAreNotEqual()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet1 = package.Workbook.Worksheets.Add("sheet1");
				var worksheet2 = package.Workbook.Worksheets.Add("sheet2");
				var range1 = new ExcelRangeBase(worksheet1, "A1:F6");
				var range2 = new ExcelRangeBase(worksheet2, "A1:F6");
				Assert.IsFalse(range1.IsEquivalentRange(range2));
				range2 = new ExcelRangeBase(worksheet1, "D1:F6");
				Assert.IsFalse(range1.IsEquivalentRange(range2));
				range2 = new ExcelRangeBase(worksheet1, "A2:F6");
				Assert.IsFalse(range1.IsEquivalentRange(range2));
			}
		}
		#endregion
	}
}
