/*******************************************************************************
* You may amend and distribute as you like, but don't remove this header!
*
* EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
* See http://www.codeplex.com/EPPlus for details.
*
* Copyright (C) 2011-2018 Michelle Lau, Evan Schallerer, and others as noted in the source history.
*
* This library is free software; you can redistribute it and/or
* modify it under the terms of the GNU Lesser General Public
* License as published by the Free Software Foundation; either
* version 2.1 of the License, or (at your option) any later version.
* This library is distributed in the hope that it will be useful,
* but WITHOUT ANY WARRANTY; without even the implied warranty of
* MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
* See the GNU Lesser General Public License for more details.
*
* The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
* If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
*
* All code and executables are provided "as is" with no warranty either express or implied. 
* The author accepts no liability for any damage or loss of business that this product may cause.
*
* For code change notes, see the source control history.
*******************************************************************************/
using System.IO;
using System.Linq;
using EPPlusTest.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest.Table.PivotTable
{
	[TestClass]
	public class ExcelPivotTableTest
	{
		#region Integration Tests
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTablesWorksheetSources.xlsx")]
		public void PivotTableXmlLoadsCorrectly()
		{
			var testFile = new FileInfo(@"PivotTablesWorksheetSources.xlsx");
			var tempFile = new FileInfo(Path.GetTempFileName());
			if (tempFile.Exists)
				tempFile.Delete();
			testFile.CopyTo(tempFile.FullName);
			try
			{
				using (var package = new ExcelPackage(tempFile))
				{
					Assert.AreEqual(2, package.Workbook.PivotCacheDefinitions.Count());

					var cacheRecords1 = package.Workbook.PivotCacheDefinitions[0].CacheRecords;
					var cacheRecords2 = package.Workbook.PivotCacheDefinitions[1].CacheRecords;

					Assert.AreNotEqual(cacheRecords1, cacheRecords2);
					Assert.AreEqual(22, cacheRecords1.Count);
					Assert.AreEqual(36, cacheRecords2.Count);
					Assert.AreEqual(cacheRecords1.Count, cacheRecords1.Count);
					Assert.AreEqual(cacheRecords2.Count, cacheRecords2.Count);

					var worksheet1 = package.Workbook.Worksheets["sheet1"];
					var worksheet2 = package.Workbook.Worksheets["sheet2"];
					var worksheet3 = package.Workbook.Worksheets["sheet3"];

					Assert.AreEqual(0, worksheet1.PivotTables.Count());
					Assert.AreEqual(2, worksheet2.PivotTables.Count());
					Assert.AreEqual(1, worksheet3.PivotTables.Count());

					Assert.AreEqual(worksheet2.PivotTables[0].CacheDefinition, worksheet2.PivotTables[1].CacheDefinition);
					Assert.AreNotEqual(worksheet2.PivotTables[0].CacheDefinition, worksheet3.PivotTables[0].CacheDefinition);
				}
			}
			finally
			{
				tempFile.Delete();
			}
		}
		#endregion

		#region Refresh Tests
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWorksheetSource.xlsx")]
		public void PivotTableRefreshFromCacheWithChangedData()
		{
			var file = new FileInfo("PivotTableWorksheetSource.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets.First();
					var pivotTable = worksheet.PivotTables.First();
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					worksheet.Cells[4, 5].Value = "Blue";
					worksheet.Cells[5, 5].Value = "Green";
					worksheet.Cells[6, 5].Value = "Purple";
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("I10:J17"), pivotTable.Address);
					Assert.AreEqual(4, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(6, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(7, pivotTable.RowItems.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "Sheet1";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 11, 9, "Blue"),
					new ExpectedCellValue(sheetName, 12, 9, "Bike"),
					new ExpectedCellValue(sheetName, 13, 9, "Green"),
					new ExpectedCellValue(sheetName, 14, 9, "Car"),
					new ExpectedCellValue(sheetName, 15, 9, "Purple"),
					new ExpectedCellValue(sheetName, 16, 9, "Skateboard"),
					new ExpectedCellValue(sheetName, 17, 9, "Grand Total"),

					new ExpectedCellValue(sheetName, 10, 10, "Sum of Cost"),
					new ExpectedCellValue(sheetName, 11, 10, 100d),
					new ExpectedCellValue(sheetName, 12, 10, 100d),
					new ExpectedCellValue(sheetName, 13, 10, 90000d),
					new ExpectedCellValue(sheetName, 14, 10, 90000d),
					new ExpectedCellValue(sheetName, 15, 10, 10d),
					new ExpectedCellValue(sheetName, 16, 10, 10d),
					new ExpectedCellValue(sheetName, 17, 10, 90110d)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWorksheetSource.xlsx")]
		public void PivotTableRefreshFromCacheWithAddedData()
		{
			var file = new FileInfo("PivotTableWorksheetSource.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets.First();
					var pivotTable = worksheet.PivotTables.First();
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					worksheet.Cells[7, 3].Value = 4;
					worksheet.Cells[7, 4].Value = "Scooter";
					worksheet.Cells[7, 5].Value = "Purple";
					worksheet.Cells[7, 6].Value = 28;
					cacheDefinition.SetSourceRangeAddress(worksheet, worksheet.Cells["C3:F7"]);
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("I10:J18"), pivotTable.Address);
					Assert.AreEqual(4, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(8, pivotTable.RowItems.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "Sheet1";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 11, 9, "Black"),
					new ExpectedCellValue(sheetName, 12, 9, "Bike"),
					new ExpectedCellValue(sheetName, 13, 9, "Skateboard"),
					new ExpectedCellValue(sheetName, 14, 9, "Purple"),
					new ExpectedCellValue(sheetName, 15, 9, "Scooter"),
					new ExpectedCellValue(sheetName, 16, 9, "Red"),
					new ExpectedCellValue(sheetName, 17, 9, "Car"),
					new ExpectedCellValue(sheetName, 18, 9, "Grand Total"),

					new ExpectedCellValue(sheetName, 10, 10, "Sum of Cost"),
					new ExpectedCellValue(sheetName, 11, 10, 110d),
					new ExpectedCellValue(sheetName, 12, 10, 100d),
					new ExpectedCellValue(sheetName, 13, 10, 10d),
					new ExpectedCellValue(sheetName, 14, 10, 28d),
					new ExpectedCellValue(sheetName, 15, 10, 28d),
					new ExpectedCellValue(sheetName, 16, 10, 90000d),
					new ExpectedCellValue(sheetName, 17, 10, 90000d),
					new ExpectedCellValue(sheetName, 18, 10, 90138d)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWorksheetSource.xlsx")]
		public void PivotTableRefreshFromCacheRemoveRow()
		{
			var file = new FileInfo("PivotTableWorksheetSource.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets.First();
					var pivotTable = worksheet.PivotTables.First();
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.SetSourceRangeAddress(worksheet, worksheet.Cells["C3:F5"]);
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("I10:J15"), pivotTable.Address);
					Assert.AreEqual(4, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(3, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(5, pivotTable.RowItems.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "Sheet1";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 11, 9, "Black"),
					new ExpectedCellValue(sheetName, 12, 9, "Bike"),
					new ExpectedCellValue(sheetName, 13, 9, "Red"),
					new ExpectedCellValue(sheetName, 14, 9, "Car"),
					new ExpectedCellValue(sheetName, 15, 9, "Grand Total"),

					new ExpectedCellValue(sheetName, 10, 10, "Sum of Cost"),
					new ExpectedCellValue(sheetName, 11, 10, 100d),
					new ExpectedCellValue(sheetName, 12, 10, 100d),
					new ExpectedCellValue(sheetName, 13, 10, 90000d),
					new ExpectedCellValue(sheetName, 14, 10, 90000d),
					new ExpectedCellValue(sheetName, 15, 10, 90100d)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithMultipleColumnFields.xlsx")]
		public void PivotTableRefreshColumnItemsWithChangedData()
		{
			var file = new FileInfo("PivotTableWithMultipleColumnFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets.First();
					var pivotTable = worksheet.PivotTables.First();
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					worksheet.Cells[4, 3].Value = "January";
					worksheet.Cells[7, 3].Value = "January";
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B12:O23"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(8, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "Sheet1";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 16, 2, "20100007"),
					new ExpectedCellValue(sheetName, 17, 2, "20100017"),
					new ExpectedCellValue(sheetName, 18, 2, "20100070"),
					new ExpectedCellValue(sheetName, 19, 2, "20100076"),
					new ExpectedCellValue(sheetName, 20, 2, "20100083"),
					new ExpectedCellValue(sheetName, 21, 2, "20100085"),
					new ExpectedCellValue(sheetName, 22, 2, "20100090"),
					new ExpectedCellValue(sheetName, 23, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 13, 3, "January"),
					new ExpectedCellValue(sheetName, 14, 3, "Car Rack"),
					new ExpectedCellValue(sheetName, 15, 3, "Chicago"),
					new ExpectedCellValue(sheetName, 16, 3, 415.75),
					new ExpectedCellValue(sheetName, 23, 3, 415.75),
					new ExpectedCellValue(sheetName, 15, 4, "Nashville"),
					new ExpectedCellValue(sheetName, 17, 4, 415.75),
					new ExpectedCellValue(sheetName, 22, 4, 415.75),
					new ExpectedCellValue(sheetName, 23, 4, 831.5),
					new ExpectedCellValue(sheetName, 15, 5, "San Francisco"),
					new ExpectedCellValue(sheetName, 19, 5, 415.75),
					new ExpectedCellValue(sheetName, 23, 5, 415.75),
					new ExpectedCellValue(sheetName, 14, 6, "Car Rack Total"),
					new ExpectedCellValue(sheetName, 16, 6, 415.75),
					new ExpectedCellValue(sheetName, 17, 6, 415.75),
					new ExpectedCellValue(sheetName, 19, 6, 415.75),
					new ExpectedCellValue(sheetName, 22, 6, 415.75),
					new ExpectedCellValue(sheetName, 23, 6, 1663d),
					new ExpectedCellValue(sheetName, 14, 7, "Headlamp"),
					new ExpectedCellValue(sheetName, 15, 7, "Chicago"),
					new ExpectedCellValue(sheetName, 20, 7, 24.99),
					new ExpectedCellValue(sheetName, 23, 7, 24.99),
					new ExpectedCellValue(sheetName, 14, 8, "Headlamp Total"),
					new ExpectedCellValue(sheetName, 20, 8, 24.99),
					new ExpectedCellValue(sheetName, 23, 8, 24.99),
					new ExpectedCellValue(sheetName, 13, 9, "January Total"),
					new ExpectedCellValue(sheetName, 16, 9, 415.75),
					new ExpectedCellValue(sheetName, 17, 9, 415.75),
					new ExpectedCellValue(sheetName, 19, 9, 415.75),
					new ExpectedCellValue(sheetName, 20, 9, 24.99),
					new ExpectedCellValue(sheetName, 22, 9, 415.75),
					new ExpectedCellValue(sheetName, 23, 9, 1687.99),
					new ExpectedCellValue(sheetName, 13, 10, "February"),
					new ExpectedCellValue(sheetName, 14, 10, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 15, 10, "San Francisco"),
					new ExpectedCellValue(sheetName, 21, 10, 99d),
					new ExpectedCellValue(sheetName, 23, 10, 99d),
					new ExpectedCellValue(sheetName, 14, 11, "Sleeping Bag Total"),
					new ExpectedCellValue(sheetName, 21, 11, 99d),
					new ExpectedCellValue(sheetName, 23, 11, 99d),
					new ExpectedCellValue(sheetName, 14, 12, "Tent"),
					new ExpectedCellValue(sheetName, 15, 12, "Nashville"),
					new ExpectedCellValue(sheetName, 18, 12, 199d),
					new ExpectedCellValue(sheetName, 23, 12, 199d),
					new ExpectedCellValue(sheetName, 14, 13, "Tent Total"),
					new ExpectedCellValue(sheetName, 18, 13, 199d),
					new ExpectedCellValue(sheetName, 13, 14, "February Total"),
					new ExpectedCellValue(sheetName, 18, 14, 199d),
					new ExpectedCellValue(sheetName, 21, 14, 99d),
					new ExpectedCellValue(sheetName, 23, 14, 298d),
					new ExpectedCellValue(sheetName, 13, 15, "Grand Total"),
					new ExpectedCellValue(sheetName, 16, 15, 415.75),
					new ExpectedCellValue(sheetName, 17, 15, 415.75),
					new ExpectedCellValue(sheetName, 18, 15, 199d),
					new ExpectedCellValue(sheetName, 19, 15, 415.75),
					new ExpectedCellValue(sheetName, 20, 15, 24.99),
					new ExpectedCellValue(sheetName, 21, 15, 99d),
					new ExpectedCellValue(sheetName, 22, 15, 415.75),
					new ExpectedCellValue(sheetName, 23, 15, 1985.99)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithMultipleColumnFields.xlsx")]
		public void PivotTableRefreshColumnItemsWithAddedData()
		{
			var file = new FileInfo("PivotTableWithMultipleColumnFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets.First();
					var pivotTable = worksheet.PivotTables.First();
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					worksheet.Cells[9, 1].Value = 20100091;
					worksheet.Cells[9, 2].Value = "Texas";
					worksheet.Cells[9, 3].Value = "December";
					worksheet.Cells[9, 4].Value = "Bike";
					worksheet.Cells[9, 5].Value = 20;
					worksheet.Cells[9, 6].Value = 1;
					worksheet.Cells[9, 7].Value = 20;
					cacheDefinition.SetSourceRangeAddress(worksheet, worksheet.Cells["A1:G9"]);
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B12:U24"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(9, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(6, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "Sheet1";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 16, 2, "20100007"),
					new ExpectedCellValue(sheetName, 17, 2, "20100017"),
					new ExpectedCellValue(sheetName, 18, 2, "20100070"),
					new ExpectedCellValue(sheetName, 19, 2, "20100076"),
					new ExpectedCellValue(sheetName, 20, 2, "20100083"),
					new ExpectedCellValue(sheetName, 21, 2, "20100085"),
					new ExpectedCellValue(sheetName, 22, 2, "20100090"),
					new ExpectedCellValue(sheetName, 23, 2, "20100091"),
					new ExpectedCellValue(sheetName, 24, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 13, 3, "January"),
					new ExpectedCellValue(sheetName, 14, 3, "Car Rack"),
					new ExpectedCellValue(sheetName, 15, 3, "Chicago"),
					new ExpectedCellValue(sheetName, 16, 3, 415.75),
					new ExpectedCellValue(sheetName, 24, 3, 415.75),
					new ExpectedCellValue(sheetName, 15, 4, "Nashville"),
					new ExpectedCellValue(sheetName, 22, 4, 415.75),
					new ExpectedCellValue(sheetName, 24, 4, 415.75),
					new ExpectedCellValue(sheetName, 15, 5, "San Francisco"),
					new ExpectedCellValue(sheetName, 19, 5, 415.75),
					new ExpectedCellValue(sheetName, 24, 5, 415.75),
					new ExpectedCellValue(sheetName, 14, 6, "Car Rack Total"),
					new ExpectedCellValue(sheetName, 16, 6, 415.75),
					new ExpectedCellValue(sheetName, 19, 6, 415.75),
					new ExpectedCellValue(sheetName, 22, 6, 415.75),
					new ExpectedCellValue(sheetName, 24, 6, 1247.25),
					new ExpectedCellValue(sheetName, 13, 7, "January Total"),
					new ExpectedCellValue(sheetName, 16, 7, 415.75),
					new ExpectedCellValue(sheetName, 19, 7, 415.75),
					new ExpectedCellValue(sheetName, 22, 7, 415.75),
					new ExpectedCellValue(sheetName, 24, 7, 1247.25),
					new ExpectedCellValue(sheetName, 13, 8, "February"),
					new ExpectedCellValue(sheetName, 14, 8, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 15, 8, "San Francisco"),
					new ExpectedCellValue(sheetName, 21, 8, 99d),
					new ExpectedCellValue(sheetName, 24, 8, 99d),
					new ExpectedCellValue(sheetName, 14, 9, "Sleeping Bag Total"),
					new ExpectedCellValue(sheetName, 21, 9, 99d),
					new ExpectedCellValue(sheetName, 24, 9, 99d),
					new ExpectedCellValue(sheetName, 14, 10, "Tent"),
					new ExpectedCellValue(sheetName, 15, 10, "Nashville"),
					new ExpectedCellValue(sheetName, 18, 10, 199d),
					new ExpectedCellValue(sheetName, 24, 10, 199d),
					new ExpectedCellValue(sheetName, 14, 11, "Tent Total"),
					new ExpectedCellValue(sheetName, 18, 11, 199d),
					new ExpectedCellValue(sheetName, 24, 11, 199d),
					new ExpectedCellValue(sheetName, 13, 12, "February Total"),
					new ExpectedCellValue(sheetName, 18, 12, 199d),
					new ExpectedCellValue(sheetName, 21, 12, 99d),
					new ExpectedCellValue(sheetName, 24, 12, 298d),
					new ExpectedCellValue(sheetName, 13, 13, "March"),
					new ExpectedCellValue(sheetName, 14, 13, "Car Rack"),
					new ExpectedCellValue(sheetName, 15, 13, "Nashville"),
					new ExpectedCellValue(sheetName, 17, 13, 415.75),
					new ExpectedCellValue(sheetName, 24, 13, 415.75),
					new ExpectedCellValue(sheetName, 14, 14, "Car Rack Total"),
					new ExpectedCellValue(sheetName, 17, 14, 415.75),
					new ExpectedCellValue(sheetName, 24, 14, 415.75),
					new ExpectedCellValue(sheetName, 14, 15, "Headlamp"),
					new ExpectedCellValue(sheetName, 15, 15, "Chicago"),
					new ExpectedCellValue(sheetName, 20, 15, 24.99),
					new ExpectedCellValue(sheetName, 24, 15, 24.99),
					new ExpectedCellValue(sheetName, 14, 16, "Headlamp Total"),
					new ExpectedCellValue(sheetName, 20, 16, 24.99),
					new ExpectedCellValue(sheetName, 24, 16, 24.99),
					new ExpectedCellValue(sheetName, 13, 17, "March Total"),
					new ExpectedCellValue(sheetName, 17, 17, 415.75),
					new ExpectedCellValue(sheetName, 20, 17, 24.99),
					new ExpectedCellValue(sheetName, 24, 17, 440.74),
					new ExpectedCellValue(sheetName, 13, 18, "December"),
					new ExpectedCellValue(sheetName, 14, 18, "Bike"),
					new ExpectedCellValue(sheetName, 15, 18, "Texas"),
					new ExpectedCellValue(sheetName, 23, 18, 20d),
					new ExpectedCellValue(sheetName, 24, 18, 20d),
					new ExpectedCellValue(sheetName, 14, 19, "Bike Total"),
					new ExpectedCellValue(sheetName, 23, 19, 20d),
					new ExpectedCellValue(sheetName, 24, 19, 20d),
					new ExpectedCellValue(sheetName, 13, 20, "December Total"),
					new ExpectedCellValue(sheetName, 23, 20, 20d),
					new ExpectedCellValue(sheetName, 24, 20, 20d),
					new ExpectedCellValue(sheetName, 13, 21, "Grand Total"),
					new ExpectedCellValue(sheetName, 16, 21, 415.75),
					new ExpectedCellValue(sheetName, 17, 21, 415.75),
					new ExpectedCellValue(sheetName, 18, 21, 199d),
					new ExpectedCellValue(sheetName, 19, 21, 415.75),
					new ExpectedCellValue(sheetName, 20, 21, 24.99),
					new ExpectedCellValue(sheetName, 21, 21, 99d),
					new ExpectedCellValue(sheetName, 22, 21, 415.75),
					new ExpectedCellValue(sheetName, 23, 21, 20d),
					new ExpectedCellValue(sheetName, 24, 21, 2005.99)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithMultipleColumnFields.xlsx")]
		public void PivotTableRefreshColumnItemsWithRemoveData()
		{
			var file = new FileInfo("PivotTableWithMultipleColumnFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets.First();
					var pivotTable = worksheet.PivotTables.First();
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.SetSourceRangeAddress(worksheet, worksheet.Cells["A1:G5"]);
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B12:M20"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(8, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "Sheet1";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 16, 2, "20100007"),
					new ExpectedCellValue(sheetName, 17, 2, "20100076"),
					new ExpectedCellValue(sheetName, 18, 2, "20100083"),
					new ExpectedCellValue(sheetName, 19, 2, "20100085"),
					new ExpectedCellValue(sheetName, 20, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 13, 3, "January"),
					new ExpectedCellValue(sheetName, 14, 3, "Car Rack"),
					new ExpectedCellValue(sheetName, 15, 3, "Chicago"),
					new ExpectedCellValue(sheetName, 16, 3, 415.75),
					new ExpectedCellValue(sheetName, 20, 3, 415.75),
					new ExpectedCellValue(sheetName, 15, 4, "San Francisco"),
					new ExpectedCellValue(sheetName, 17, 4, 415.75),
					new ExpectedCellValue(sheetName, 20, 4, 415.75),
					new ExpectedCellValue(sheetName, 14, 5, "Car Rack Total"),
					new ExpectedCellValue(sheetName, 16, 5, 415.75),
					new ExpectedCellValue(sheetName, 17, 5, 415.75),
					new ExpectedCellValue(sheetName, 20, 5, 831.5),
					new ExpectedCellValue(sheetName, 13, 6, "January Total"),
					new ExpectedCellValue(sheetName, 16, 6, 415.75),
					new ExpectedCellValue(sheetName, 17, 6, 415.75),
					new ExpectedCellValue(sheetName, 20, 6, 831.5),
					new ExpectedCellValue(sheetName, 13, 7, "February"),
					new ExpectedCellValue(sheetName, 14, 7, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 15, 7, "San Francisco"),
					new ExpectedCellValue(sheetName, 19, 7, 99d),
					new ExpectedCellValue(sheetName, 20, 7, 99d),
					new ExpectedCellValue(sheetName, 14, 8, "Sleeping Bag Total"),
					new ExpectedCellValue(sheetName, 19, 8, 99d),
					new ExpectedCellValue(sheetName, 20, 8, 99d),
					new ExpectedCellValue(sheetName, 13, 9, "February Total"),
					new ExpectedCellValue(sheetName, 19, 9, 99d),
					new ExpectedCellValue(sheetName, 20, 9, 99d),
					new ExpectedCellValue(sheetName, 13, 10, "March"),
					new ExpectedCellValue(sheetName, 14, 10, "Headlamp"),
					new ExpectedCellValue(sheetName, 15, 10, "Chicago"),
					new ExpectedCellValue(sheetName, 18, 10, 24.99),
					new ExpectedCellValue(sheetName, 20, 10, 24.99),
					new ExpectedCellValue(sheetName, 14, 11, "Headlamp Total"),
					new ExpectedCellValue(sheetName, 18, 11, 24.99),
					new ExpectedCellValue(sheetName, 20, 11, 24.99),
					new ExpectedCellValue(sheetName, 13, 12, "March Total"),
					new ExpectedCellValue(sheetName, 18, 12, 24.99),
					new ExpectedCellValue(sheetName, 20, 12, 24.99),
					new ExpectedCellValue(sheetName, 13, 13, "Grand Total"),
					new ExpectedCellValue(sheetName, 16, 13, 415.75),
					new ExpectedCellValue(sheetName, 17, 13, 415.75),
					new ExpectedCellValue(sheetName, 18, 13, 24.99),
					new ExpectedCellValue(sheetName, 19, 13, 99d),
					new ExpectedCellValue(sheetName, 20, 13, 955.49)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithMultipleColumnFields.xlsx")]
		public void PivotTableRefreshDeletingSourceRow()
		{
			var file = new FileInfo("PivotTableWithMultipleColumnFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets.First();
					var pivotTable = worksheet.PivotTables.First();
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					worksheet.DeleteRow(6);
					cacheDefinition.SetSourceRangeAddress(worksheet, worksheet.Cells["A1:G7"]);
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B11:P21"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(8, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "Sheet1";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 15, 2, "20100007"),
					new ExpectedCellValue(sheetName, 16, 2, "20100017"),
					new ExpectedCellValue(sheetName, 17, 2, "20100076"),
					new ExpectedCellValue(sheetName, 18, 2, "20100083"),
					new ExpectedCellValue(sheetName, 19, 2, "20100085"),
					new ExpectedCellValue(sheetName, 20, 2, "20100090"),
					new ExpectedCellValue(sheetName, 21, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 12, 3, "January"),
					new ExpectedCellValue(sheetName, 13, 3, "Car Rack"),
					new ExpectedCellValue(sheetName, 14, 3, "Chicago"),
					new ExpectedCellValue(sheetName, 15, 3, 415.75),
					new ExpectedCellValue(sheetName, 21, 3, 415.75),
					new ExpectedCellValue(sheetName, 14, 4, "Nashville"),
					new ExpectedCellValue(sheetName, 20, 4, 415.75),
					new ExpectedCellValue(sheetName, 21, 4, 415.75),
					new ExpectedCellValue(sheetName, 14, 5, "San Francisco"),
					new ExpectedCellValue(sheetName, 17, 5, 415.75),
					new ExpectedCellValue(sheetName, 21, 5, 415.75),
					new ExpectedCellValue(sheetName, 13, 6, "Car Rack Total"),
					new ExpectedCellValue(sheetName, 15, 6, 415.75),
					new ExpectedCellValue(sheetName, 17, 6, 415.75),
					new ExpectedCellValue(sheetName, 20, 6, 415.75),
					new ExpectedCellValue(sheetName, 21, 6, 1247.25),
					new ExpectedCellValue(sheetName, 12, 7, "January Total"),
					new ExpectedCellValue(sheetName, 15, 7, 415.75),
					new ExpectedCellValue(sheetName, 17, 7, 415.75),
					new ExpectedCellValue(sheetName, 20, 7, 415.75),
					new ExpectedCellValue(sheetName, 21, 7, 1247.25),
					new ExpectedCellValue(sheetName, 12, 8, "February"),
					new ExpectedCellValue(sheetName, 13, 8, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 14, 8, "San Francisco"),
					new ExpectedCellValue(sheetName, 19, 8, 99d),
					new ExpectedCellValue(sheetName, 21, 8, 99d),
					new ExpectedCellValue(sheetName, 13, 9, "Sleeping Bag Total"),
					new ExpectedCellValue(sheetName, 19, 9, 99d),
					new ExpectedCellValue(sheetName, 21, 9, 99d),
					new ExpectedCellValue(sheetName, 12, 10, "February Total"),
					new ExpectedCellValue(sheetName, 19, 10, 99d),
					new ExpectedCellValue(sheetName, 21, 10, 99d),
					new ExpectedCellValue(sheetName, 12, 11, "March"),
					new ExpectedCellValue(sheetName, 13, 11, "Car Rack"),
					new ExpectedCellValue(sheetName, 14, 11, "Nashville"),
					new ExpectedCellValue(sheetName, 16, 11, 415.75),
					new ExpectedCellValue(sheetName, 21, 11, 415.75),
					new ExpectedCellValue(sheetName, 13, 12, "Car Rack Total"),
					new ExpectedCellValue(sheetName, 16, 12, 415.75),
					new ExpectedCellValue(sheetName, 21, 12, 415.75),
					new ExpectedCellValue(sheetName, 13, 13, "Headlamp"),
					new ExpectedCellValue(sheetName, 14, 13, "Chicago"),
					new ExpectedCellValue(sheetName, 18, 13, 24.99),
					new ExpectedCellValue(sheetName, 21, 13, 24.99),
					new ExpectedCellValue(sheetName, 13, 14, "Headlamp Total"),
					new ExpectedCellValue(sheetName, 18, 14, 24.99),
					new ExpectedCellValue(sheetName, 21, 14, 24.99),
					new ExpectedCellValue(sheetName, 12, 15, "March Total"),
					new ExpectedCellValue(sheetName, 16, 15, 415.75),
					new ExpectedCellValue(sheetName, 18, 15, 24.99),
					new ExpectedCellValue(sheetName, 21, 15, 440.74),
					new ExpectedCellValue(sheetName, 12, 16, "Grand Total"),
					new ExpectedCellValue(sheetName, 15, 16, 415.75),
					new ExpectedCellValue(sheetName, 16, 16, 415.75),
					new ExpectedCellValue(sheetName, 17, 16, 415.75),
					new ExpectedCellValue(sheetName, 18, 16, 24.99),
					new ExpectedCellValue(sheetName, 19, 16, 99d),
					new ExpectedCellValue(sheetName, 20, 16, 415.75),
					new ExpectedCellValue(sheetName, 21, 16, 1786.99)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithRowFieldsOnly.xlsx")]
		public void PivotTableRefreshSingleColumnNoDataFields()
		{
			var file = new FileInfo("PivotTableWithRowFieldsOnly.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["PivotTables"];
					var pivotTable = worksheet.PivotTables["PivotTable1"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("A1:A5"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "PivotTables";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 2, 1, "January"),
					new ExpectedCellValue(sheetName, 3, 1, "February"),
					new ExpectedCellValue(sheetName, 4, 1, "March"),
					new ExpectedCellValue(sheetName, 5, 1, "Grand Total")
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithRowFieldsOnly.xlsx")]
		public void PivotTableRefreshSingleColumnTwoRowFieldsAndNoDataFields()
		{
			var file = new FileInfo("PivotTableWithRowFieldsOnly.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["PivotTables"];
					var pivotTable = worksheet.PivotTables["PivotTable2"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("E1:E12"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "PivotTables";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 2, 5, "January"),
					new ExpectedCellValue(sheetName, 3, 5, "Chicago"),
					new ExpectedCellValue(sheetName, 4, 5, "Nashville"),
					new ExpectedCellValue(sheetName, 5, 5, "San Francisco"),
					new ExpectedCellValue(sheetName, 6, 5, "February"),
					new ExpectedCellValue(sheetName, 7, 5, "Nashville"),
					new ExpectedCellValue(sheetName, 8, 5, "San Francisco"),
					new ExpectedCellValue(sheetName, 9, 5, "March"),
					new ExpectedCellValue(sheetName, 10, 5, "Chicago"),
					new ExpectedCellValue(sheetName, 11, 5, "Nashville"),
					new ExpectedCellValue(sheetName, 12, 5, "Grand Total")
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithMAttrbuteFieldItems.xlsx")]
		public void PivotTableRefreshFieldItemsWithMAttributes()
		{
			var file = new FileInfo("PivotTableWithMAttrbuteFieldItems.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["Sheet1"];
					var pivotTable = worksheet.PivotTables["PivotTable1"];
					worksheet.Cells[2, 3].Value = "December";
					worksheet.Cells[5, 3].Value = "December";
					worksheet.Cells[8, 3].Value = "December";
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("C15:D19"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					foreach (var field in pivotTable.Fields)
					{
						if (field.Items.Count > 0)
						{
							foreach (var item in field.Items)
							{
								Assert.IsNull(item.TopNode.Attributes["m"]);
								Assert.AreEqual(1, item.TopNode.Attributes.Count);
							}
						}
					}
					package.SaveAs(newFile.File);
				}
				string sheetName = "Sheet1";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 16, 3, "February"),
					new ExpectedCellValue(sheetName, 17, 3, "March"),
					new ExpectedCellValue(sheetName, 18, 3, "December"),
					new ExpectedCellValue(sheetName, 19, 3, "Grand Total"),
					new ExpectedCellValue(sheetName, 15, 4, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 16, 4, 7d),
					new ExpectedCellValue(sheetName, 17, 4, 3d),
					new ExpectedCellValue(sheetName, 18, 4, 5d),
					new ExpectedCellValue(sheetName, 19, 4, 15d)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithMonthInSpanish.xlsx")]
		public void PivotTableRefreshSortMonthsInSpanishCorrectly()
		{
			var file = new FileInfo("PivotTableWithMonthInSpanish.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets.First();
					var pivotTable = worksheet.PivotTables.First();
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("A13:F25"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "Sheet1";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 15, 1, "Chicago"),
					new ExpectedCellValue(sheetName, 16, 1, "enero"),
					new ExpectedCellValue(sheetName, 17, 1, "febrero"),
					new ExpectedCellValue(sheetName, 18, 1, "Nashville"),
					new ExpectedCellValue(sheetName, 19, 1, "enero"),
					new ExpectedCellValue(sheetName, 20, 1, "febrero"),
					new ExpectedCellValue(sheetName, 21, 1, "marzo"),
					new ExpectedCellValue(sheetName, 22, 1, "San Francisco"),
					new ExpectedCellValue(sheetName, 23, 1, "enero"),
					new ExpectedCellValue(sheetName, 24, 1, "marzo"),
					new ExpectedCellValue(sheetName, 25, 1, "Grand Total"),
					new ExpectedCellValue(sheetName, 14, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 15, 2, 831.5),
					new ExpectedCellValue(sheetName, 16, 2, 831.5),
					new ExpectedCellValue(sheetName, 18, 2, 1663),
					new ExpectedCellValue(sheetName, 19, 2, 831.5),
					new ExpectedCellValue(sheetName, 21, 2, 831.5),
					new ExpectedCellValue(sheetName, 22, 2, 415.75),
					new ExpectedCellValue(sheetName, 24, 2, 415.75),
					new ExpectedCellValue(sheetName, 25, 2, 2910.25),
					new ExpectedCellValue(sheetName, 14, 3, "Headlamp"),
					new ExpectedCellValue(sheetName, 15, 3, 24.99),
					new ExpectedCellValue(sheetName, 17, 3, 24.99),
					new ExpectedCellValue(sheetName, 25, 3, 24.99),
					new ExpectedCellValue(sheetName, 14, 4, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 22, 4, 99d),
					new ExpectedCellValue(sheetName, 23, 4, 99d),
					new ExpectedCellValue(sheetName, 25, 4, 99d),
					new ExpectedCellValue(sheetName, 14, 5, "Tent"),
					new ExpectedCellValue(sheetName, 18, 5, 1194d),
					new ExpectedCellValue(sheetName, 20, 5, 1194d),
					new ExpectedCellValue(sheetName, 25, 5, 1194d),
					new ExpectedCellValue(sheetName, 14, 6, "Grand Total"),
					new ExpectedCellValue(sheetName, 15, 6, 856.49),
					new ExpectedCellValue(sheetName, 16, 6, 831.5),
					new ExpectedCellValue(sheetName, 17, 6, 24.99),
					new ExpectedCellValue(sheetName, 18, 6, 2857d),
					new ExpectedCellValue(sheetName, 19, 6, 831.5),
					new ExpectedCellValue(sheetName, 20, 6, 1194d),
					new ExpectedCellValue(sheetName, 21, 6, 831.5),
					new ExpectedCellValue(sheetName, 22, 6, 514.75),
					new ExpectedCellValue(sheetName, 23, 6, 99d),
					new ExpectedCellValue(sheetName, 24, 6, 415.75),
					new ExpectedCellValue(sheetName, 25, 6, 4228.24)

				});
			}
		}

		#region Calculated Fields Tests
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithCalculatedFields.xlsx")]
		public void PivotTableRefreshCalculatedField()
		{
			var file = new FileInfo("PivotTableWithCalculatedFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "Sheet2";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable1"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					var formulaCacheField = cacheDefinition.CacheFields.First(c => !string.IsNullOrEmpty(c.Formula));
					formulaCacheField.Formula = "'Wholesale Price'";
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B2:D13"), pivotTable.Address);
					Assert.AreEqual(9, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[7].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[8].Items.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 2, 3, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 2, 4, "Sum of CalculatedField"),
					new ExpectedCellValue(sheetName, 3, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 3, 3, 7),
					new ExpectedCellValue(sheetName, 3, 4, 1663),
					new ExpectedCellValue(sheetName, 4, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 4, 3, 2),
					new ExpectedCellValue(sheetName, 4, 4, 415.75),
					new ExpectedCellValue(sheetName, 5, 2, "Nashville"),
					new ExpectedCellValue(sheetName, 5, 3, 4),
					new ExpectedCellValue(sheetName, 5, 4, 831.5),
					new ExpectedCellValue(sheetName, 6, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 6, 3, 1),
					new ExpectedCellValue(sheetName, 6, 4, 415.75),
					new ExpectedCellValue(sheetName, 7, 2, "Headlamp"),
					new ExpectedCellValue(sheetName, 7, 3, 1),
					new ExpectedCellValue(sheetName, 7, 4, 24.99),
					new ExpectedCellValue(sheetName, 8, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 8, 3, 1),
					new ExpectedCellValue(sheetName, 8, 4, 24.99),
					new ExpectedCellValue(sheetName, 9, 2, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 9, 3, 1),
					new ExpectedCellValue(sheetName, 9, 4, 99),
					new ExpectedCellValue(sheetName, 10, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 10, 3, 1),
					new ExpectedCellValue(sheetName, 10, 4, 99),
					new ExpectedCellValue(sheetName, 11, 2, "Tent"),
					new ExpectedCellValue(sheetName, 11, 3, 6),
					new ExpectedCellValue(sheetName, 11, 4, 199),
					new ExpectedCellValue(sheetName, 12, 2, "Nashville"),
					new ExpectedCellValue(sheetName, 12, 3, 6),
					new ExpectedCellValue(sheetName, 12, 4, 199),
					new ExpectedCellValue(sheetName, 13, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 13, 3, 15),
					new ExpectedCellValue(sheetName, 13, 4, 1985.99),
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithCalculatedFields.xlsx")]
		public void PivotTableRefreshCalculatedFieldFormulaMultipliesFields()
		{
			var file = new FileInfo("PivotTableWithCalculatedFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "Sheet2";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable1"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B2:D13"), pivotTable.Address);
					Assert.AreEqual(9, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[7].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[8].Items.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 2, 3, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 2, 4, "Sum of CalculatedField"),
					new ExpectedCellValue(sheetName, 3, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 3, 3, 7),
					new ExpectedCellValue(sheetName, 3, 4, 11641),
					new ExpectedCellValue(sheetName, 4, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 4, 3, 2),
					new ExpectedCellValue(sheetName, 4, 4, 831.5),
					new ExpectedCellValue(sheetName, 5, 2, "Nashville"),
					new ExpectedCellValue(sheetName, 5, 3, 4),
					new ExpectedCellValue(sheetName, 5, 4, 3326),
					new ExpectedCellValue(sheetName, 6, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 6, 3, 1),
					new ExpectedCellValue(sheetName, 6, 4, 415.75),
					new ExpectedCellValue(sheetName, 7, 2, "Headlamp"),
					new ExpectedCellValue(sheetName, 7, 3, 1),
					new ExpectedCellValue(sheetName, 7, 4, 24.99),
					new ExpectedCellValue(sheetName, 8, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 8, 3, 1),
					new ExpectedCellValue(sheetName, 8, 4, 24.99),
					new ExpectedCellValue(sheetName, 9, 2, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 9, 3, 1),
					new ExpectedCellValue(sheetName, 9, 4, 99),
					new ExpectedCellValue(sheetName, 10, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 10, 3, 1),
					new ExpectedCellValue(sheetName, 10, 4, 99),
					new ExpectedCellValue(sheetName, 11, 2, "Tent"),
					new ExpectedCellValue(sheetName, 11, 3, 6),
					new ExpectedCellValue(sheetName, 11, 4, 1194),
					new ExpectedCellValue(sheetName, 12, 2, "Nashville"),
					new ExpectedCellValue(sheetName, 12, 3, 6),
					new ExpectedCellValue(sheetName, 12, 4, 1194),
					new ExpectedCellValue(sheetName, 13, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 13, 3, 15),
					new ExpectedCellValue(sheetName, 13, 4, 29789.85),
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithCalculatedFields.xlsx")]
		public void PivotTableRefreshCalculatedFieldFormulaContainsStringField()
		{
			var file = new FileInfo("PivotTableWithCalculatedFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "Sheet2";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable1"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					var formulaCacheField = cacheDefinition.CacheFields.First(c => !string.IsNullOrEmpty(c.Formula));
					formulaCacheField.Formula = "'Wholesale Price' * Item";
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B2:D13"), pivotTable.Address);
					Assert.AreEqual(9, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[7].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[8].Items.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 2, 3, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 2, 4, "Sum of CalculatedField"),
					new ExpectedCellValue(sheetName, 3, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 3, 3, 7),
					new ExpectedCellValue(sheetName, 3, 4, 0),
					new ExpectedCellValue(sheetName, 4, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 4, 3, 2),
					new ExpectedCellValue(sheetName, 4, 4, 0),
					new ExpectedCellValue(sheetName, 5, 2, "Nashville"),
					new ExpectedCellValue(sheetName, 5, 3, 4),
					new ExpectedCellValue(sheetName, 5, 4, 0),
					new ExpectedCellValue(sheetName, 6, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 6, 3, 1),
					new ExpectedCellValue(sheetName, 6, 4, 0),
					new ExpectedCellValue(sheetName, 7, 2, "Headlamp"),
					new ExpectedCellValue(sheetName, 7, 3, 1),
					new ExpectedCellValue(sheetName, 7, 4, 0),
					new ExpectedCellValue(sheetName, 8, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 8, 3, 1),
					new ExpectedCellValue(sheetName, 8, 4, 0),
					new ExpectedCellValue(sheetName, 9, 2, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 9, 3, 1),
					new ExpectedCellValue(sheetName, 9, 4, 0),
					new ExpectedCellValue(sheetName, 10, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 10, 3, 1),
					new ExpectedCellValue(sheetName, 10, 4, 0),
					new ExpectedCellValue(sheetName, 11, 2, "Tent"),
					new ExpectedCellValue(sheetName, 11, 3, 6),
					new ExpectedCellValue(sheetName, 11, 4, 0),
					new ExpectedCellValue(sheetName, 12, 2, "Nashville"),
					new ExpectedCellValue(sheetName, 12, 3, 6),
					new ExpectedCellValue(sheetName, 12, 4, 0),
					new ExpectedCellValue(sheetName, 13, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 13, 3, 15),
					new ExpectedCellValue(sheetName, 13, 4, 0),
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithCalculatedFields.xlsx")]
		public void PivotTableRefreshCalculatedFieldFormulaReferencesOtherCalculatedField()
		{
			var file = new FileInfo("PivotTableWithCalculatedFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "Sheet2";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable2"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("B18:D27"), pivotTable.Address);
					Assert.AreEqual(9, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[7].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[8].Items.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 18, 3, "Sum of CalculatedField"),
					new ExpectedCellValue(sheetName, 18, 4, "Sum of CalculatedField2"),
					new ExpectedCellValue(sheetName, 19, 2, "January"),
					new ExpectedCellValue(sheetName, 19, 3, 6236.25),
					new ExpectedCellValue(sheetName, 19, 4, 5),
					new ExpectedCellValue(sheetName, 20, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 20, 3, 6236.25),
					new ExpectedCellValue(sheetName, 20, 4, 5),
					new ExpectedCellValue(sheetName, 21, 2, "February"),
					new ExpectedCellValue(sheetName, 21, 3, 2086.00),
					new ExpectedCellValue(sheetName, 21, 4, 7),
					new ExpectedCellValue(sheetName, 22, 2, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 22, 3, 99d),
					new ExpectedCellValue(sheetName, 22, 4, 1),
					new ExpectedCellValue(sheetName, 23, 2, "Tent"),
					new ExpectedCellValue(sheetName, 23, 3, 1194d),
					new ExpectedCellValue(sheetName, 23, 4, 6),
					new ExpectedCellValue(sheetName, 24, 2, "March"),
					new ExpectedCellValue(sheetName, 24, 3, 1322.22),
					new ExpectedCellValue(sheetName, 24, 4, 3),
					new ExpectedCellValue(sheetName, 25, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 25, 3, 831.5),
					new ExpectedCellValue(sheetName, 25, 4, 2),
					new ExpectedCellValue(sheetName, 26, 2, "Headlamp"),
					new ExpectedCellValue(sheetName, 26, 3, 24.99),
					new ExpectedCellValue(sheetName, 26, 4, 1),
					new ExpectedCellValue(sheetName, 27, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 27, 3, 29789.85),
					new ExpectedCellValue(sheetName, 27, 4, 15),
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithCalculatedFields.xlsx")]
		public void PivotTableRefreshColumnRowAndPageFieldsWithCalculatedField()
		{
			var file = new FileInfo("PivotTableWithCalculatedFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "Sheet2";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable3"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("G4:L14"), pivotTable.Address);
					Assert.AreEqual(9, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[7].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[8].Items.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 4, 7, null),
					new ExpectedCellValue(sheetName, 4, 8, "Column Labels"),
					new ExpectedCellValue(sheetName, 4, 9, null),
					new ExpectedCellValue(sheetName, 4, 10, null),
					new ExpectedCellValue(sheetName, 4, 11, null),
					new ExpectedCellValue(sheetName, 4, 12, null),
					new ExpectedCellValue(sheetName, 5, 7, null),
					new ExpectedCellValue(sheetName, 5, 8, "Chicago"),
					new ExpectedCellValue(sheetName, 5, 9, "Chicago Total"),
					new ExpectedCellValue(sheetName, 5, 10, "Nashville"),
					new ExpectedCellValue(sheetName, 5, 11, "Nashville Total"),
					new ExpectedCellValue(sheetName, 5, 12, "Grand Total"),
					new ExpectedCellValue(sheetName, 6, 7, "Row Labels"),
					new ExpectedCellValue(sheetName, 6, 8, "Car Rack"),
					new ExpectedCellValue(sheetName, 6, 9, null),
					new ExpectedCellValue(sheetName, 6, 10, "Car Rack"),
					new ExpectedCellValue(sheetName, 6, 11, null),
					new ExpectedCellValue(sheetName, 6, 12, null),
					new ExpectedCellValue(sheetName, 7, 7, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 7, 8, null),
					new ExpectedCellValue(sheetName, 7, 9, null),
					new ExpectedCellValue(sheetName, 7, 10, null),
					new ExpectedCellValue(sheetName, 7, 11, null),
					new ExpectedCellValue(sheetName, 7, 12, null),
					new ExpectedCellValue(sheetName, 8, 7, "January"),
					new ExpectedCellValue(sheetName, 8, 8, 415.75),
					new ExpectedCellValue(sheetName, 8, 9, 415.75),
					new ExpectedCellValue(sheetName, 8, 10, 415.75),
					new ExpectedCellValue(sheetName, 8, 11, 415.75),
					new ExpectedCellValue(sheetName, 8, 12, 831.5),
					new ExpectedCellValue(sheetName, 9, 7, "March"),
					new ExpectedCellValue(sheetName, 9, 8, null),
					new ExpectedCellValue(sheetName, 9, 9, null),
					new ExpectedCellValue(sheetName, 9, 10, 415.75),
					new ExpectedCellValue(sheetName, 9, 11, 415.75),
					new ExpectedCellValue(sheetName, 9, 12, 415.75),
					new ExpectedCellValue(sheetName, 10, 7, "Sum of CalculatedField"),
					new ExpectedCellValue(sheetName, 10, 8, null),
					new ExpectedCellValue(sheetName, 10, 9, null),
					new ExpectedCellValue(sheetName, 10, 10, null),
					new ExpectedCellValue(sheetName, 10, 11, null),
					new ExpectedCellValue(sheetName, 10, 12, null),
					new ExpectedCellValue(sheetName, 11, 7, "January"),
					new ExpectedCellValue(sheetName, 11, 8, 831.5),
					new ExpectedCellValue(sheetName, 11, 9, 831.5),
					new ExpectedCellValue(sheetName, 11, 10, 831.5),
					new ExpectedCellValue(sheetName, 11, 11, 831.5),
					new ExpectedCellValue(sheetName, 11, 12, 3326d),
					new ExpectedCellValue(sheetName, 12, 7, "March"),
					new ExpectedCellValue(sheetName, 12, 8, 0d),
					new ExpectedCellValue(sheetName, 12, 9, 0d),
					new ExpectedCellValue(sheetName, 12, 10, 831.5),
					new ExpectedCellValue(sheetName, 12, 11, 831.5),
					new ExpectedCellValue(sheetName, 12, 12, 831.5),
					new ExpectedCellValue(sheetName, 13, 7, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 13, 8, 415.75),
					new ExpectedCellValue(sheetName, 13, 9, 415.75),
					new ExpectedCellValue(sheetName, 13, 10, 831.5),
					new ExpectedCellValue(sheetName, 13, 11, 831.5),
					new ExpectedCellValue(sheetName, 13, 12, 1247.25),
					new ExpectedCellValue(sheetName, 14, 7, "Total Sum of CalculatedField"),
					new ExpectedCellValue(sheetName, 14, 8, 831.5),
					new ExpectedCellValue(sheetName, 14, 9, 831.5),
					new ExpectedCellValue(sheetName, 14, 10, 3326d),
					new ExpectedCellValue(sheetName, 14, 11, 3326d),
					new ExpectedCellValue(sheetName, 14, 12, 7483.5),
				});
			}
		}
		#endregion
		#endregion

		#region UpdateData Field Values Tests
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithNoColumnFields.xlsx")]
		public void PivotTableRefreshDataFieldOneRowFieldWithTrueSubtotalTop()
		{
			var file = new FileInfo("PivotTableWithNoColumnFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["PivotTables"];
					var pivotTable = worksheet.PivotTables["PivotTable1"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("A1:B5"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "PivotTables";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 2, 1, "January"),
					new ExpectedCellValue(sheetName, 3, 1, "February"),
					new ExpectedCellValue(sheetName, 4, 1, "March"),
					new ExpectedCellValue(sheetName, 5, 1, "Grand Total"),
					new ExpectedCellValue(sheetName, 2, 2, 2078.75),
					new ExpectedCellValue(sheetName, 3, 2, 1293d),
					new ExpectedCellValue(sheetName, 4, 2, 856.49),
					new ExpectedCellValue(sheetName, 5, 2, 4228.24)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithNoColumnFields.xlsx")]
		public void PivotTableRefreshDataFieldTwoRowFieldsWithTrueSubtotalTop()
		{
			var file = new FileInfo("PivotTableWithNoColumnFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["PivotTables"];
					var pivotTable = worksheet.PivotTables["PivotTable2"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("F1:G10"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "PivotTables";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 2, 6, "January"),
					new ExpectedCellValue(sheetName, 3, 6, "Car Rack"),
					new ExpectedCellValue(sheetName, 4, 6, "February"),
					new ExpectedCellValue(sheetName, 5, 6, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 6, 6, "Tent"),
					new ExpectedCellValue(sheetName, 7, 6, "March"),
					new ExpectedCellValue(sheetName, 8, 6, "Car Rack"),
					new ExpectedCellValue(sheetName, 9, 6, "Headlamp"),
					new ExpectedCellValue(sheetName, 10, 6, "Grand Total"),
					new ExpectedCellValue(sheetName, 2, 7, 2078.75),
					new ExpectedCellValue(sheetName, 3, 7, 2078.75),
					new ExpectedCellValue(sheetName, 4, 7, 1293d),
					new ExpectedCellValue(sheetName, 5, 7, 99d),
					new ExpectedCellValue(sheetName, 6, 7, 1194d),
					new ExpectedCellValue(sheetName, 7, 7, 856.49),
					new ExpectedCellValue(sheetName, 8, 7, 831.5),
					new ExpectedCellValue(sheetName, 9, 7, 24.99),
					new ExpectedCellValue(sheetName, 10, 7, 4228.24)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithNoColumnFields.xlsx")]
		public void PivotTableRefreshDataFieldTwoRowFieldsWithFalseSubtotalTop()
		{
			var file = new FileInfo("PivotTableWithNoColumnFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["PivotTables"];
					var pivotTable = worksheet.PivotTables["PivotTable2"];
					foreach (var field in pivotTable.Fields)
					{
						field.SubtotalTop = false;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("F1:G13"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "PivotTables";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 2, 6, "January"),
					new ExpectedCellValue(sheetName, 3, 6, "Car Rack"),
					new ExpectedCellValue(sheetName, 4, 6, "January Total"),
					new ExpectedCellValue(sheetName, 5, 6, "February"),
					new ExpectedCellValue(sheetName, 6, 6, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 7, 6, "Tent"),
					new ExpectedCellValue(sheetName, 8, 6, "February Total"),
					new ExpectedCellValue(sheetName, 9, 6, "March"),
					new ExpectedCellValue(sheetName, 10, 6, "Car Rack"),
					new ExpectedCellValue(sheetName, 11, 6, "Headlamp"),
					new ExpectedCellValue(sheetName, 12, 6, "March Total"),
					new ExpectedCellValue(sheetName, 13, 6, "Grand Total"),
					new ExpectedCellValue(sheetName, 3, 7, 2078.75),
					new ExpectedCellValue(sheetName, 4, 7, 2078.75),
					new ExpectedCellValue(sheetName, 6, 7, 99d),
					new ExpectedCellValue(sheetName, 7, 7, 1194d),
					new ExpectedCellValue(sheetName, 8, 7, 1293d),
					new ExpectedCellValue(sheetName, 10, 7, 831.5),
					new ExpectedCellValue(sheetName, 11, 7, 24.99),
					new ExpectedCellValue(sheetName, 12, 7, 856.49),
					new ExpectedCellValue(sheetName, 13, 7, 4228.24)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithNoColumnFields.xlsx")]
		public void PivotTableRefreshDataFieldTwoRowFieldsWithNoSubtotal()
		{
			var file = new FileInfo("PivotTableWithNoColumnFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["PivotTables"];
					var pivotTable = worksheet.PivotTables["PivotTable2"];
					foreach (var field in pivotTable.Fields)
					{
						field.DisableDefaultSubtotal();
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("F1:G10"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(3, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(3, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "PivotTables";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 2, 6, "January"),
					new ExpectedCellValue(sheetName, 3, 6, "Car Rack"),
					new ExpectedCellValue(sheetName, 4, 6, "February"),
					new ExpectedCellValue(sheetName, 5, 6, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 6, 6, "Tent"),
					new ExpectedCellValue(sheetName, 7, 6, "March"),
					new ExpectedCellValue(sheetName, 8, 6, "Car Rack"),
					new ExpectedCellValue(sheetName, 9, 6, "Headlamp"),
					new ExpectedCellValue(sheetName, 10, 6, "Grand Total"),
					new ExpectedCellValue(sheetName, 3, 7, 2078.75),
					new ExpectedCellValue(sheetName, 5, 7, 99d),
					new ExpectedCellValue(sheetName, 6, 7, 1194d),
					new ExpectedCellValue(sheetName, 8, 7, 831.5),
					new ExpectedCellValue(sheetName, 9, 7, 24.99),
					new ExpectedCellValue(sheetName, 10, 7, 4228.24)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithNoColumnFields.xlsx")]
		public void PivotTableRefreshDataFieldThreeRowFieldsWithTrueSubtotalTop()
		{
			var file = new FileInfo("PivotTableWithNoColumnFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["PivotTables"];
					var pivotTable = worksheet.PivotTables["PivotTable3"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("L1:M17"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "PivotTables";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 2, 12, "January"),
					new ExpectedCellValue(sheetName, 3, 12, "Car Rack"),
					new ExpectedCellValue(sheetName, 4, 12, "Chicago"),
					new ExpectedCellValue(sheetName, 5, 12, "Nashville"),
					new ExpectedCellValue(sheetName, 6, 12, "San Francisco"),
					new ExpectedCellValue(sheetName, 7, 12, "February"),
					new ExpectedCellValue(sheetName, 8, 12, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 9, 12, "San Francisco"),
					new ExpectedCellValue(sheetName, 10, 12, "Tent"),
					new ExpectedCellValue(sheetName, 11, 12, "Nashville"),
					new ExpectedCellValue(sheetName, 12, 12, "March"),
					new ExpectedCellValue(sheetName, 13, 12, "Car Rack"),
					new ExpectedCellValue(sheetName, 14, 12, "Nashville"),
					new ExpectedCellValue(sheetName, 15, 12, "Headlamp"),
					new ExpectedCellValue(sheetName, 16, 12, "Chicago"),
					new ExpectedCellValue(sheetName, 17, 12, "Grand Total"),
					new ExpectedCellValue(sheetName, 2, 13, 2078.75),
					new ExpectedCellValue(sheetName, 3, 13, 2078.75),
					new ExpectedCellValue(sheetName, 4, 13, 831.5),
					new ExpectedCellValue(sheetName, 5, 13, 831.5),
					new ExpectedCellValue(sheetName, 6, 13, 415.75),
					new ExpectedCellValue(sheetName, 7, 13, 1293d),
					new ExpectedCellValue(sheetName, 8, 13, 99d),
					new ExpectedCellValue(sheetName, 9, 13, 99d),
					new ExpectedCellValue(sheetName, 10, 13, 1194d),
					new ExpectedCellValue(sheetName, 11, 13, 1194d),
					new ExpectedCellValue(sheetName, 12, 13, 856.49),
					new ExpectedCellValue(sheetName, 13, 13, 831.5),
					new ExpectedCellValue(sheetName, 14, 13, 831.5),
					new ExpectedCellValue(sheetName, 15, 13, 24.99),
					new ExpectedCellValue(sheetName, 16, 13, 24.99),
					new ExpectedCellValue(sheetName, 17, 13, 4228.24)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithNoColumnFields.xlsx")]
		public void PivotTableRefreshDataFieldThreeRowFieldsWithFalseSubtotalTop()
		{
			var file = new FileInfo("PivotTableWithNoColumnFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["PivotTables"];
					var pivotTable = worksheet.PivotTables["PivotTable3"];
					foreach (var field in pivotTable.Fields)
					{
						field.SubtotalTop = false;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("L1:M25"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "PivotTables";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 2, 12, "January"),
					new ExpectedCellValue(sheetName, 3, 12, "Car Rack"),
					new ExpectedCellValue(sheetName, 4, 12, "Chicago"),
					new ExpectedCellValue(sheetName, 5, 12, "Nashville"),
					new ExpectedCellValue(sheetName, 6, 12, "San Francisco"),
					new ExpectedCellValue(sheetName, 7, 12, "Car Rack Total"),
					new ExpectedCellValue(sheetName, 8, 12, "January Total"),
					new ExpectedCellValue(sheetName, 9, 12, "February"),
					new ExpectedCellValue(sheetName, 10, 12, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 11, 12, "San Francisco"),
					new ExpectedCellValue(sheetName, 12, 12, "Sleeping Bag Total"),
					new ExpectedCellValue(sheetName, 13, 12, "Tent"),
					new ExpectedCellValue(sheetName, 14, 12, "Nashville"),
					new ExpectedCellValue(sheetName, 15, 12, "Tent Total"),
					new ExpectedCellValue(sheetName, 16, 12, "February Total"),
					new ExpectedCellValue(sheetName, 17, 12, "March"),
					new ExpectedCellValue(sheetName, 18, 12, "Car Rack"),
					new ExpectedCellValue(sheetName, 19, 12, "Nashville"),
					new ExpectedCellValue(sheetName, 20, 12, "Car Rack Total"),
					new ExpectedCellValue(sheetName, 21, 12, "Headlamp"),
					new ExpectedCellValue(sheetName, 22, 12, "Chicago"),
					new ExpectedCellValue(sheetName, 23, 12, "Headlamp Total"),
					new ExpectedCellValue(sheetName, 24, 12, "March Total"),
					new ExpectedCellValue(sheetName, 25, 12, "Grand Total"),
					new ExpectedCellValue(sheetName, 4, 13, 831.5),
					new ExpectedCellValue(sheetName, 5, 13, 831.5),
					new ExpectedCellValue(sheetName, 6, 13, 415.75),
					new ExpectedCellValue(sheetName, 7, 13, 2078.75),
					new ExpectedCellValue(sheetName, 8, 13, 2078.75),
					new ExpectedCellValue(sheetName, 11, 13, 99d),
					new ExpectedCellValue(sheetName, 12, 13, 99d),
					new ExpectedCellValue(sheetName, 14, 13, 1194d),
					new ExpectedCellValue(sheetName, 15, 13, 1194d),
					new ExpectedCellValue(sheetName, 16, 13, 1293d),
					new ExpectedCellValue(sheetName, 19, 13, 831.5),
					new ExpectedCellValue(sheetName, 20, 13, 831.5),
					new ExpectedCellValue(sheetName, 22, 13, 24.99),
					new ExpectedCellValue(sheetName, 23, 13, 24.99),
					new ExpectedCellValue(sheetName, 24, 13, 856.49),
					new ExpectedCellValue(sheetName, 25, 13, 4228.24)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithNoSubtotals.xlsx")]
		public void PivotTableRefreshDataFieldsRowsAndColumnsWithNoSubtotal()
		{
			var file = new FileInfo("PivotTableWithNoSubtotals.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["NoSubtotals"];
					var pivotTable = worksheet.PivotTables["PivotTable1"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("A1:I13"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(3, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(3, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(3, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "NoSubtotals";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 4, 1, "Car Rack"),
					new ExpectedCellValue(sheetName, 5, 1, "1"),
					new ExpectedCellValue(sheetName, 6, 1, "2"),
					new ExpectedCellValue(sheetName, 7, 1, "Headlamp"),
					new ExpectedCellValue(sheetName, 8, 1, "1"),
					new ExpectedCellValue(sheetName, 9, 1, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 10, 1, "1"),
					new ExpectedCellValue(sheetName, 11, 1, "Tent"),
					new ExpectedCellValue(sheetName, 12, 1, "6"),
					new ExpectedCellValue(sheetName, 13, 1, "Grand Total"),
					new ExpectedCellValue(sheetName, 2, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 3, 2, "January"),
					new ExpectedCellValue(sheetName, 6, 2, 415.75),
					new ExpectedCellValue(sheetName, 13, 2, 415.75),
					new ExpectedCellValue(sheetName, 3, 3, "March"),
					new ExpectedCellValue(sheetName, 8, 3, 24.99),
					new ExpectedCellValue(sheetName, 13, 3, 24.99),
					new ExpectedCellValue(sheetName, 2, 4, "Nashville"),
					new ExpectedCellValue(sheetName, 3, 4, "January"),
					new ExpectedCellValue(sheetName, 6, 4, 415.75),
					new ExpectedCellValue(sheetName, 13, 2, 415.75),
					new ExpectedCellValue(sheetName, 3, 5, "February"),
					new ExpectedCellValue(sheetName, 12, 5, 199d),
					new ExpectedCellValue(sheetName, 13, 5, 199d),
					new ExpectedCellValue(sheetName, 3, 6, "March"),
					new ExpectedCellValue(sheetName, 6, 6, 415.75),
					new ExpectedCellValue(sheetName, 13, 6, 415.75),
					new ExpectedCellValue(sheetName, 2, 7, "San Francisco"),
					new ExpectedCellValue(sheetName, 3, 7, "January"),
					new ExpectedCellValue(sheetName, 5, 7, 415.75),
					new ExpectedCellValue(sheetName, 13, 7, 415.75),
					new ExpectedCellValue(sheetName, 3, 8, "February"),
					new ExpectedCellValue(sheetName, 10, 8, 99d),
					new ExpectedCellValue(sheetName, 13, 8, 99d),
					new ExpectedCellValue(sheetName, 2, 9, "Grand Total"),
					new ExpectedCellValue(sheetName, 5, 9, 415.75),
					new ExpectedCellValue(sheetName, 6, 9, 1247.25),
					new ExpectedCellValue(sheetName, 8, 9, 24.99),
					new ExpectedCellValue(sheetName, 10, 9, 99d),
					new ExpectedCellValue(sheetName, 12, 9, 199d),
					new ExpectedCellValue(sheetName, 13, 9, 1985.99),
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithGrandTotals.xlsx")]
		public void PivotTableRefreshDataFieldsRowsAndColumnsGrandTotalOff()
		{
			var file = new FileInfo("PivotTableWithGrandTotals.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["GrandTotals"];
					var pivotTable = worksheet.PivotTables["PivotTable1"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					pivotTable.RowGrandTotals = false;
					pivotTable.ColumnGrandTotals = false;
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("A1:K7"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "GrandTotals";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 4, 1, "Car Rack"),
					new ExpectedCellValue(sheetName, 5, 1, "Headlamp"),
					new ExpectedCellValue(sheetName, 6, 1, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 7, 1, "Tent"),
					new ExpectedCellValue(sheetName, 2, 2, "January"),
					new ExpectedCellValue(sheetName, 3, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 4, 2, 831.5),
					new ExpectedCellValue(sheetName, 3, 3, "Nashville"),
					new ExpectedCellValue(sheetName, 4, 3, 831.5),
					new ExpectedCellValue(sheetName, 3, 4, "San Francisco"),
					new ExpectedCellValue(sheetName, 4, 4, 415.75),
					new ExpectedCellValue(sheetName, 2, 5, "January Total"),
					new ExpectedCellValue(sheetName, 4, 5, 2078.75),
					new ExpectedCellValue(sheetName, 2, 6, "February"),
					new ExpectedCellValue(sheetName, 3, 6, "Nashville"),
					new ExpectedCellValue(sheetName, 7, 6, 1194d),
					new ExpectedCellValue(sheetName, 3, 7, "San Francisco"),
					new ExpectedCellValue(sheetName, 6, 7, 99d),
					new ExpectedCellValue(sheetName, 2, 8, "February Total"),
					new ExpectedCellValue(sheetName, 6, 8, 99d),
					new ExpectedCellValue(sheetName, 7, 8, 1194d),
					new ExpectedCellValue(sheetName, 2, 9, "March"),
					new ExpectedCellValue(sheetName, 3, 9, "Chicago"),
					new ExpectedCellValue(sheetName, 5, 9, 24.99),
					new ExpectedCellValue(sheetName, 3, 10, "Nashville"),
					new ExpectedCellValue(sheetName, 4, 10, 831.5),
					new ExpectedCellValue(sheetName, 2, 11, "March Total"),
					new ExpectedCellValue(sheetName, 4, 11, 831.5),
					new ExpectedCellValue(sheetName, 5, 11, 24.99)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithGrandTotals.xlsx")]
		public void PivotTableRefreshDataFieldsColumnGrandTotalOff()
		{
			var file = new FileInfo("PivotTableWithGrandTotals.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["GrandTotals"];
					var pivotTable = worksheet.PivotTables["PivotTable1"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					pivotTable.ColumnGrandTotals = false;
					pivotTable.RowGrandTotals = true;
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("A1:K8"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "GrandTotals";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 4, 1, "Car Rack"),
					new ExpectedCellValue(sheetName, 5, 1, "Headlamp"),
					new ExpectedCellValue(sheetName, 6, 1, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 7, 1, "Tent"),
					new ExpectedCellValue(sheetName, 8, 1, "Grand Total"),
					new ExpectedCellValue(sheetName, 2, 2, "January"),
					new ExpectedCellValue(sheetName, 3, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 4, 2, 831.5),
					new ExpectedCellValue(sheetName, 8, 2, 831.5),
					new ExpectedCellValue(sheetName, 3, 3, "Nashville"),
					new ExpectedCellValue(sheetName, 4, 3, 831.5),
					new ExpectedCellValue(sheetName, 8, 3, 831.5),
					new ExpectedCellValue(sheetName, 3, 4, "San Francisco"),
					new ExpectedCellValue(sheetName, 4, 4, 415.75),
					new ExpectedCellValue(sheetName, 8, 4, 415.75),
					new ExpectedCellValue(sheetName, 2, 5, "January Total"),
					new ExpectedCellValue(sheetName, 4, 5, 2078.75),
					new ExpectedCellValue(sheetName, 8, 5, 2078.75),
					new ExpectedCellValue(sheetName, 2, 6, "February"),
					new ExpectedCellValue(sheetName, 3, 6, "Nashville"),
					new ExpectedCellValue(sheetName, 7, 6, 1194d),
					new ExpectedCellValue(sheetName, 8, 6, 1194d),
					new ExpectedCellValue(sheetName, 3, 7, "San Francisco"),
					new ExpectedCellValue(sheetName, 6, 7, 99d),
					new ExpectedCellValue(sheetName, 8, 7, 99d),
					new ExpectedCellValue(sheetName, 2, 8, "February Total"),
					new ExpectedCellValue(sheetName, 6, 8, 99d),
					new ExpectedCellValue(sheetName, 7, 8, 1194d),
					new ExpectedCellValue(sheetName, 8, 8, 1293d),
					new ExpectedCellValue(sheetName, 2, 9, "March"),
					new ExpectedCellValue(sheetName, 3, 9, "Chicago"),
					new ExpectedCellValue(sheetName, 5, 9, 24.99),
					new ExpectedCellValue(sheetName, 8, 9, 24.99),
					new ExpectedCellValue(sheetName, 3, 10, "Nashville"),
					new ExpectedCellValue(sheetName, 4, 10, 831.5),
					new ExpectedCellValue(sheetName, 8, 10, 831.5),
					new ExpectedCellValue(sheetName, 2, 11, "March Total"),
					new ExpectedCellValue(sheetName, 4, 11, 831.5),
					new ExpectedCellValue(sheetName, 5, 11, 24.99),
					new ExpectedCellValue(sheetName, 8, 11, 856.49)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithGrandTotals.xlsx")]
		public void PivotTableRefreshDataFieldsRowGrandTotalOff()
		{
			var file = new FileInfo("PivotTableWithGrandTotals.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["GrandTotals"];
					var pivotTable = worksheet.PivotTables["PivotTable1"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					pivotTable.RowGrandTotals = false;
					pivotTable.ColumnGrandTotals = true;
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("A1:L7"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "GrandTotals";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 4, 1, "Car Rack"),
					new ExpectedCellValue(sheetName, 5, 1, "Headlamp"),
					new ExpectedCellValue(sheetName, 6, 1, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 7, 1, "Tent"),
					new ExpectedCellValue(sheetName, 2, 2, "January"),
					new ExpectedCellValue(sheetName, 3, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 4, 2, 831.5),
					new ExpectedCellValue(sheetName, 3, 3, "Nashville"),
					new ExpectedCellValue(sheetName, 4, 3, 831.5),
					new ExpectedCellValue(sheetName, 3, 4, "San Francisco"),
					new ExpectedCellValue(sheetName, 4, 4, 415.75),
					new ExpectedCellValue(sheetName, 2, 5, "January Total"),
					new ExpectedCellValue(sheetName, 4, 5, 2078.75),
					new ExpectedCellValue(sheetName, 2, 6, "February"),
					new ExpectedCellValue(sheetName, 3, 6, "Nashville"),
					new ExpectedCellValue(sheetName, 7, 6, 1194d),
					new ExpectedCellValue(sheetName, 3, 7, "San Francisco"),
					new ExpectedCellValue(sheetName, 6, 7, 99d),
					new ExpectedCellValue(sheetName, 2, 8, "February Total"),
					new ExpectedCellValue(sheetName, 6, 8, 99d),
					new ExpectedCellValue(sheetName, 7, 8, 1194d),
					new ExpectedCellValue(sheetName, 2, 9, "March"),
					new ExpectedCellValue(sheetName, 3, 9, "Chicago"),
					new ExpectedCellValue(sheetName, 5, 9, 24.99),
					new ExpectedCellValue(sheetName, 3, 10, "Nashville"),
					new ExpectedCellValue(sheetName, 4, 10, 831.5),
					new ExpectedCellValue(sheetName, 2, 11, "March Total"),
					new ExpectedCellValue(sheetName, 4, 11, 831.5),
					new ExpectedCellValue(sheetName, 5, 11, 24.99),
					new ExpectedCellValue(sheetName, 2, 12, "Grand Total"),
					new ExpectedCellValue(sheetName, 4, 12, 2910.25),
					new ExpectedCellValue(sheetName, 5, 12, 24.99),
					new ExpectedCellValue(sheetName, 6, 12, 99d),
					new ExpectedCellValue(sheetName, 7, 12, 1194d)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithMultipleDataFields.xlsx")]
		public void PivotTableRefreshMultipleDataFields()
		{
			var file = new FileInfo("PivotTableWithMultipleDataFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["MultipleDataFields"];
					var pivotTable = worksheet.PivotTables["PivotTable1"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("A1:C5"), pivotTable.Address);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "MultipleDataFields";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 2, 1, "Chicago"),
					new ExpectedCellValue(sheetName, 3, 1, "Nashville"),
					new ExpectedCellValue(sheetName, 4, 1, "San Francisco"),
					new ExpectedCellValue(sheetName, 5, 1, "Grand Total"),
					new ExpectedCellValue(sheetName, 1, 2, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 2, 2, 3),
					new ExpectedCellValue(sheetName, 3, 2, 10),
					new ExpectedCellValue(sheetName, 4, 2, 2),
					new ExpectedCellValue(sheetName, 5, 2, 15),
					new ExpectedCellValue(sheetName, 1, 3, "Sum of Total"),
					new ExpectedCellValue(sheetName, 2, 3, 856.49),
					new ExpectedCellValue(sheetName, 3, 3, 2857d),
					new ExpectedCellValue(sheetName, 4, 3, 514.75),
					new ExpectedCellValue(sheetName, 5, 3, 4228.24),
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithMultipleDataFields.xlsx")]
		public void PivotTableRefreshMultipleDataFieldsNoGrandTotal()
		{
			var file = new FileInfo("PivotTableWithMultipleDataFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["MultipleDataFields"];
					var pivotTable = worksheet.PivotTables["PivotTable1"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					pivotTable.RowGrandTotals = false;
					pivotTable.ColumnGrandTotals = true;
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("A1:C4"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "MultipleDataFields";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 2, 1, "Chicago"),
					new ExpectedCellValue(sheetName, 3, 1, "Nashville"),
					new ExpectedCellValue(sheetName, 4, 1, "San Francisco"),
					new ExpectedCellValue(sheetName, 1, 2, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 2, 2, 3),
					new ExpectedCellValue(sheetName, 3, 2, 10),
					new ExpectedCellValue(sheetName, 4, 2, 2),
					new ExpectedCellValue(sheetName, 1, 3, "Sum of Total"),
					new ExpectedCellValue(sheetName, 2, 3, 856.49),
					new ExpectedCellValue(sheetName, 3, 3, 2857d),
					new ExpectedCellValue(sheetName, 4, 3, 514.75)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithMultipleDataFields.xlsx")]
		public void PivotTableRefreshMultipleDataFieldsWithColumnHeaders()
		{
			var file = new FileInfo("PivotTableWithMultipleDataFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["MultipleDataFields"];
					var pivotTable = worksheet.PivotTables["PivotTable2"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("A10:W17"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "MultipleDataFields";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 14, 1, "January"),
					new ExpectedCellValue(sheetName, 15, 1, "February"),
					new ExpectedCellValue(sheetName, 16, 1, "March"),
					new ExpectedCellValue(sheetName, 17, 1, "Grand Total"),
					new ExpectedCellValue(sheetName, 11, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 12, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 13, 2, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 14, 2, 2d),
					new ExpectedCellValue(sheetName, 17, 2, 2d),
					new ExpectedCellValue(sheetName, 13, 3, "Sum of Total"),
					new ExpectedCellValue(sheetName, 14, 3, 831.5),
					new ExpectedCellValue(sheetName, 17, 3, 831.5),
					new ExpectedCellValue(sheetName, 12, 4, "Nashville"),
					new ExpectedCellValue(sheetName, 13, 4, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 14, 4, 2d),
					new ExpectedCellValue(sheetName, 16, 4, 2d),
					new ExpectedCellValue(sheetName, 17, 4, 4d),
					new ExpectedCellValue(sheetName, 13, 5, "Sum of Total"),
					new ExpectedCellValue(sheetName, 14, 5, 831.5),
					new ExpectedCellValue(sheetName, 16, 5, 831.5),
					new ExpectedCellValue(sheetName, 17, 5, 1663d),
					new ExpectedCellValue(sheetName, 12, 6, "San Francisco"),
					new ExpectedCellValue(sheetName, 13, 6, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 14, 6, 1d),
					new ExpectedCellValue(sheetName, 17, 6, 1d),
					new ExpectedCellValue(sheetName, 13, 7, "Sum of Total"),
					new ExpectedCellValue(sheetName, 14, 7, 415.75),
					new ExpectedCellValue(sheetName, 17, 7, 415.75),
					new ExpectedCellValue(sheetName, 11, 8, "Car Rack Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 14, 8, 5d),
					new ExpectedCellValue(sheetName, 16, 8, 2d),
					new ExpectedCellValue(sheetName, 17, 8, 7d),
					new ExpectedCellValue(sheetName, 11, 9, "Car Rack Sum of Total"),
					new ExpectedCellValue(sheetName, 14, 9, 2078.75),
					new ExpectedCellValue(sheetName, 16, 9, 831.5),
					new ExpectedCellValue(sheetName, 17, 9, 2910.25),
					new ExpectedCellValue(sheetName, 11, 10, "Headlamp"),
					new ExpectedCellValue(sheetName, 12, 10, "Chicago"),
					new ExpectedCellValue(sheetName, 13, 10, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 16, 10, 1d),
					new ExpectedCellValue(sheetName, 17, 10, 1d),
					new ExpectedCellValue(sheetName, 13, 11, "Sum of Total"),
					new ExpectedCellValue(sheetName, 16, 11, 24.99),
					new ExpectedCellValue(sheetName, 17, 11, 24.99),
					new ExpectedCellValue(sheetName, 11, 12, "Headlamp Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 16, 12, 1d),
					new ExpectedCellValue(sheetName, 17, 12, 1d),
					new ExpectedCellValue(sheetName, 11, 13, "Headlamp Sum of Total"),
					new ExpectedCellValue(sheetName, 16, 13, 24.99),
					new ExpectedCellValue(sheetName, 17, 13, 24.99),
					new ExpectedCellValue(sheetName, 11, 14, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 12, 14, "San Francisco"),
					new ExpectedCellValue(sheetName, 13, 14, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 15, 14, 1d),
					new ExpectedCellValue(sheetName, 17, 14, 1d),
					new ExpectedCellValue(sheetName, 13, 15, "Sum of Total"),
					new ExpectedCellValue(sheetName, 15, 15, 99d),
					new ExpectedCellValue(sheetName, 17, 15, 99d),
					new ExpectedCellValue(sheetName, 11, 16, "Sleeping Bag Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 15, 16, 1d),
					new ExpectedCellValue(sheetName, 17, 16, 1d),
					new ExpectedCellValue(sheetName, 11, 17, "Sleeping Bag Sum of Total"),
					new ExpectedCellValue(sheetName, 15, 17, 99d),
					new ExpectedCellValue(sheetName, 17, 17, 99d),
					new ExpectedCellValue(sheetName, 11, 18, "Tent"),
					new ExpectedCellValue(sheetName, 12, 18, "Nashville"),
					new ExpectedCellValue(sheetName, 13, 18, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 15, 18, 6d),
					new ExpectedCellValue(sheetName, 17, 18, 6d),
					new ExpectedCellValue(sheetName, 13, 19, "Sum of Total"),
					new ExpectedCellValue(sheetName, 15, 19, 1194d),
					new ExpectedCellValue(sheetName, 17, 19, 1194d),
					new ExpectedCellValue(sheetName, 11, 20, "Tent Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 15, 20, 6d),
					new ExpectedCellValue(sheetName, 17, 20, 6d),
					new ExpectedCellValue(sheetName, 11, 21, "Tent Sum of Total"),
					new ExpectedCellValue(sheetName, 15, 21, 1194d),
					new ExpectedCellValue(sheetName, 17, 21, 1194d),
					new ExpectedCellValue(sheetName, 11, 22, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 14, 22, 5d),
					new ExpectedCellValue(sheetName, 15, 22, 7d),
					new ExpectedCellValue(sheetName, 16, 22, 3d),
					new ExpectedCellValue(sheetName, 17, 22, 15d),
					new ExpectedCellValue(sheetName, 11, 23, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 14, 23, 2078.75),
					new ExpectedCellValue(sheetName, 15, 23, 1293d),
					new ExpectedCellValue(sheetName, 16, 23, 856.49),
					new ExpectedCellValue(sheetName, 17, 23, 4228.24)
				});
			}
		}

		#region Multiple Row Data Fields
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithMultipleRowDataFieldsOneRowFieldOneColumnField.xlsx")]
		public void PivotTableRefreshMultipleRowDataFieldsOneRowAndOneColumn()
		{
			var file = new FileInfo("PivotTableWithMultipleRowDataFieldsOneRowFieldOneColumnField.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["RowDataFields"];
					var pivotTable = worksheet.PivotTables["PivotTable1"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("A1:E13"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "RowDataFields";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 3, 1, "Chicago"),
					new ExpectedCellValue(sheetName, 4, 1, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 5, 1, "Sum of Total"),
					new ExpectedCellValue(sheetName, 6, 1, "Nashville"),
					new ExpectedCellValue(sheetName, 7, 1, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 8, 1, "Sum of Total"),
					new ExpectedCellValue(sheetName, 9, 1, "San Francisco"),
					new ExpectedCellValue(sheetName, 10, 1, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 11, 1, "Sum of Total"),
					new ExpectedCellValue(sheetName, 12, 1, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 13, 1, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 2, 2, "January"),
					new ExpectedCellValue(sheetName, 4, 2, 2d),
					new ExpectedCellValue(sheetName, 5, 2, 831.5),
					new ExpectedCellValue(sheetName, 7, 2, 2d),
					new ExpectedCellValue(sheetName, 8, 2, 831.5),
					new ExpectedCellValue(sheetName, 10, 2, 1d),
					new ExpectedCellValue(sheetName, 11, 2, 415.75),
					new ExpectedCellValue(sheetName, 12, 2, 5d),
					new ExpectedCellValue(sheetName, 13, 2, 2078.75),
					new ExpectedCellValue(sheetName, 2, 3, "February"),
					new ExpectedCellValue(sheetName, 7, 3, 6d),
					new ExpectedCellValue(sheetName, 8, 3, 1194d),
					new ExpectedCellValue(sheetName, 10, 3, 1d),
					new ExpectedCellValue(sheetName, 11, 3, 99d),
					new ExpectedCellValue(sheetName, 12, 3, 7d),
					new ExpectedCellValue(sheetName, 13, 3, 1293d),
					new ExpectedCellValue(sheetName, 2, 4, "March"),
					new ExpectedCellValue(sheetName, 4, 4, 1d),
					new ExpectedCellValue(sheetName, 5, 4, 24.99),
					new ExpectedCellValue(sheetName, 7, 4, 2d),
					new ExpectedCellValue(sheetName, 8, 4, 831.5),
					new ExpectedCellValue(sheetName, 12, 4, 3d),
					new ExpectedCellValue(sheetName, 13, 4, 856.49),
					new ExpectedCellValue(sheetName, 2, 5, "Grand Total"),
					new ExpectedCellValue(sheetName, 4, 5, 3d),
					new ExpectedCellValue(sheetName, 5, 5, 856.49),
					new ExpectedCellValue(sheetName, 7, 5, 10d),
					new ExpectedCellValue(sheetName, 8, 5, 2857d),
					new ExpectedCellValue(sheetName, 10, 5, 2d),
					new ExpectedCellValue(sheetName, 11, 5, 514.75),
					new ExpectedCellValue(sheetName, 12, 5, 15d),
					new ExpectedCellValue(sheetName, 13, 5, 4228.24)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithMultipleRowDataFieldsTwoRowFieldOneColumnField.xlsx")]
		public void PivotTableRefreshMultipleRowDataFieldsTwoRowsAndOneColumnSubtotalsOff()
		{
			var file = new FileInfo("PivotTableWithMultipleRowDataFieldsTwoRowFieldOneColumnField.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["RowDataFields"];
					var pivotTable = worksheet.PivotTables["PivotTable1"];
					foreach (var field in pivotTable.Fields)
					{
						field.DisableDefaultSubtotal();
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("A1:E25"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(3, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(3, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "RowDataFields";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 3, 1, "Chicago"),
					new ExpectedCellValue(sheetName, 4, 1, "Car Rack"),
					new ExpectedCellValue(sheetName, 5, 1, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 6, 1, "Sum of Total"),
					new ExpectedCellValue(sheetName, 7, 1, "Headlamp"),
					new ExpectedCellValue(sheetName, 8, 1, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 9, 1, "Sum of Total"),
					new ExpectedCellValue(sheetName, 10, 1, "Nashville"),
					new ExpectedCellValue(sheetName, 11, 1, "Car Rack"),
					new ExpectedCellValue(sheetName, 12, 1, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 13, 1, "Sum of Total"),
					new ExpectedCellValue(sheetName, 14, 1, "Tent"),
					new ExpectedCellValue(sheetName, 15, 1, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 16, 1, "Sum of Total"),
					new ExpectedCellValue(sheetName, 17, 1, "San Francisco"),
					new ExpectedCellValue(sheetName, 18, 1, "Car Rack"),
					new ExpectedCellValue(sheetName, 19, 1, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 20, 1, "Sum of Total"),
					new ExpectedCellValue(sheetName, 21, 1, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 22, 1, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 23, 1, "Sum of Total"),
					new ExpectedCellValue(sheetName, 24, 1, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 25, 1, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 2, 2, "January"),
					new ExpectedCellValue(sheetName, 5, 2, 2d),
					new ExpectedCellValue(sheetName, 6, 2, 831.5),
					new ExpectedCellValue(sheetName, 12, 2, 2d),
					new ExpectedCellValue(sheetName, 13, 2, 831.5),
					new ExpectedCellValue(sheetName, 19, 2, 1d),
					new ExpectedCellValue(sheetName, 20, 2, 415.75),
					new ExpectedCellValue(sheetName, 24, 2, 5d),
					new ExpectedCellValue(sheetName, 25, 2, 2078.75),
					new ExpectedCellValue(sheetName, 2, 3, "February"),
					new ExpectedCellValue(sheetName, 15, 3, 6d),
					new ExpectedCellValue(sheetName, 16, 3, 1194d),
					new ExpectedCellValue(sheetName, 22, 3, 1d),
					new ExpectedCellValue(sheetName, 23, 3, 99d),
					new ExpectedCellValue(sheetName, 24, 3, 7d),
					new ExpectedCellValue(sheetName, 25, 3, 1293d),
					new ExpectedCellValue(sheetName, 2, 4, "March"),
					new ExpectedCellValue(sheetName, 8, 4, 1d),
					new ExpectedCellValue(sheetName, 9, 4, 24.99),
					new ExpectedCellValue(sheetName, 12, 4, 2d),
					new ExpectedCellValue(sheetName, 13, 4, 831.5),
					new ExpectedCellValue(sheetName, 24, 4, 3d),
					new ExpectedCellValue(sheetName, 25, 4, 856.49),
					new ExpectedCellValue(sheetName, 2, 5, "Grand Total"),
					new ExpectedCellValue(sheetName, 5, 5, 2d),
					new ExpectedCellValue(sheetName, 6, 5, 831.5),
					new ExpectedCellValue(sheetName, 8, 5, 1d),
					new ExpectedCellValue(sheetName, 9, 5, 24.99),
					new ExpectedCellValue(sheetName, 12, 5, 4d),
					new ExpectedCellValue(sheetName, 13, 5, 1663d),
					new ExpectedCellValue(sheetName, 15, 5, 6d),
					new ExpectedCellValue(sheetName, 16, 5, 1194d),
					new ExpectedCellValue(sheetName, 19, 5, 1d),
					new ExpectedCellValue(sheetName, 20, 5, 415.75),
					new ExpectedCellValue(sheetName, 22, 5, 1d),
					new ExpectedCellValue(sheetName, 23, 5, 99d),
					new ExpectedCellValue(sheetName, 24, 5, 15d),
					new ExpectedCellValue(sheetName, 25, 5, 4228.24)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithMultipleRowDataFieldsTwoRowFieldOneColumnField.xlsx")]
		public void PivotTableRefreshMultipleRowDataFieldsTwoRowsAndOneColumnSubtotalsOn()
		{
			var file = new FileInfo("PivotTableWithMultipleRowDataFieldsTwoRowFieldOneColumnField.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["RowDataFields"];
					var pivotTable = worksheet.PivotTables["PivotTable1"];
					foreach (var field in pivotTable.Fields)
					{
						field.SubtotalTop = false;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("A1:E31"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "RowDataFields";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 3, 1, "Chicago"),
					new ExpectedCellValue(sheetName, 4, 1, "Car Rack"),
					new ExpectedCellValue(sheetName, 5, 1, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 6, 1, "Sum of Total"),
					new ExpectedCellValue(sheetName, 7, 1, "Headlamp"),
					new ExpectedCellValue(sheetName, 8, 1, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 9, 1, "Sum of Total"),
					new ExpectedCellValue(sheetName, 10, 1, "Chicago Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 11, 1, "Chicago Sum of Total"),
					new ExpectedCellValue(sheetName, 12, 1, "Nashville"),
					new ExpectedCellValue(sheetName, 13, 1, "Car Rack"),
					new ExpectedCellValue(sheetName, 14, 1, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 15, 1, "Sum of Total"),
					new ExpectedCellValue(sheetName, 16, 1, "Tent"),
					new ExpectedCellValue(sheetName, 17, 1, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 18, 1, "Sum of Total"),
					new ExpectedCellValue(sheetName, 19, 1, "Nashville Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 20, 1, "Nashville Sum of Total"),
					new ExpectedCellValue(sheetName, 21, 1, "San Francisco"),
					new ExpectedCellValue(sheetName, 22, 1, "Car Rack"),
					new ExpectedCellValue(sheetName, 23, 1, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 24, 1, "Sum of Total"),
					new ExpectedCellValue(sheetName, 25, 1, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 26, 1, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 27, 1, "Sum of Total"),
					new ExpectedCellValue(sheetName, 28, 1, "San Francisco Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 29, 1, "San Francisco Sum of Total"),
					new ExpectedCellValue(sheetName, 30, 1, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 31, 1, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 2, 2, "January"),
					new ExpectedCellValue(sheetName, 5, 2, 2d),
					new ExpectedCellValue(sheetName, 6, 2, 831.5),
					new ExpectedCellValue(sheetName, 10, 2, 2d),
					new ExpectedCellValue(sheetName, 11, 2, 831.5),
					new ExpectedCellValue(sheetName, 14, 2, 2d),
					new ExpectedCellValue(sheetName, 15, 2, 831.5),
					new ExpectedCellValue(sheetName, 19, 2, 2d),
					new ExpectedCellValue(sheetName, 20, 2, 831.5),
					new ExpectedCellValue(sheetName, 23, 2, 1d),
					new ExpectedCellValue(sheetName, 24, 2, 415.75),
					new ExpectedCellValue(sheetName, 28, 2, 1d),
					new ExpectedCellValue(sheetName, 29, 2, 415.75),
					new ExpectedCellValue(sheetName, 30, 2, 5d),
					new ExpectedCellValue(sheetName, 31, 2, 2078.75),
					new ExpectedCellValue(sheetName, 2, 3, "February"),
					new ExpectedCellValue(sheetName, 17, 3, 6d),
					new ExpectedCellValue(sheetName, 18, 3, 1194d),
					new ExpectedCellValue(sheetName, 19, 3, 6d),
					new ExpectedCellValue(sheetName, 20, 3, 1194d),
					new ExpectedCellValue(sheetName, 26, 3, 1d),
					new ExpectedCellValue(sheetName, 27, 3, 99d),
					new ExpectedCellValue(sheetName, 28, 3, 1d),
					new ExpectedCellValue(sheetName, 29, 3, 99d),
					new ExpectedCellValue(sheetName, 30, 3, 7d),
					new ExpectedCellValue(sheetName, 31, 3, 1293d),
					new ExpectedCellValue(sheetName, 2, 4, "March"),
					new ExpectedCellValue(sheetName, 8, 4, 1d),
					new ExpectedCellValue(sheetName, 9, 4, 24.99),
					new ExpectedCellValue(sheetName, 10, 4, 1d),
					new ExpectedCellValue(sheetName, 11, 4, 24.99),
					new ExpectedCellValue(sheetName, 14, 4, 2d),
					new ExpectedCellValue(sheetName, 15, 4, 831.5),
					new ExpectedCellValue(sheetName, 19, 4, 2d),
					new ExpectedCellValue(sheetName, 20, 4, 831.5),
					new ExpectedCellValue(sheetName, 30, 4, 3d),
					new ExpectedCellValue(sheetName, 31, 4, 856.49),
					new ExpectedCellValue(sheetName, 2, 5, "Grand Total"),
					new ExpectedCellValue(sheetName, 5, 5, 2d),
					new ExpectedCellValue(sheetName, 6, 5, 831.5),
					new ExpectedCellValue(sheetName, 8, 5, 1d),
					new ExpectedCellValue(sheetName, 9, 5, 24.99),
					new ExpectedCellValue(sheetName, 10, 5, 3d),
					new ExpectedCellValue(sheetName, 11, 5, 856.49),
					new ExpectedCellValue(sheetName, 14, 5, 4d),
					new ExpectedCellValue(sheetName, 15, 5, 1663d),
					new ExpectedCellValue(sheetName, 17, 5, 6d),
					new ExpectedCellValue(sheetName, 18, 5, 1194d),
					new ExpectedCellValue(sheetName, 19, 5, 10d),
					new ExpectedCellValue(sheetName, 20, 5, 2857d),
					new ExpectedCellValue(sheetName, 23, 5, 1d),
					new ExpectedCellValue(sheetName, 24, 5, 415.75),
					new ExpectedCellValue(sheetName, 26, 5, 1d),
					new ExpectedCellValue(sheetName, 27, 5, 99d),
					new ExpectedCellValue(sheetName, 28, 5, 2d),
					new ExpectedCellValue(sheetName, 29, 5, 514.75),
					new ExpectedCellValue(sheetName, 30, 5, 15d),
					new ExpectedCellValue(sheetName, 31, 5, 4228.24)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithMultipleRowDataFieldsThreeRowFieldOneColumnField.xlsx")]
		public void PivotTableRefreshMultipleRowDataFieldsThreeRowsAndOneColumnSubtotalsOff()
		{
			var file = new FileInfo("PivotTableWithMultipleRowDataFieldsThreeRowFieldOneColumnField.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["RowDataFields"];
					var pivotTable = worksheet.PivotTables["PivotTable1"];
					foreach (var field in pivotTable.Fields)
					{
						field.DisableDefaultSubtotal();
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("A1:E34"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(7, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(3, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(3, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "RowDataFields";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 3, 1, "Chicago"),
					new ExpectedCellValue(sheetName, 4, 1, "Car Rack"),
					new ExpectedCellValue(sheetName, 5, 1, "20100007"),
					new ExpectedCellValue(sheetName, 6, 1, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 7, 1, "Sum of Total"),
					new ExpectedCellValue(sheetName, 8, 1, "Headlamp"),
					new ExpectedCellValue(sheetName, 9, 1, "20100083"),
					new ExpectedCellValue(sheetName, 10, 1, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 11, 1, "Sum of Total"),
					new ExpectedCellValue(sheetName, 12, 1, "Nashville"),
					new ExpectedCellValue(sheetName, 13, 1, "Car Rack"),
					new ExpectedCellValue(sheetName, 14, 1, "20100017"),
					new ExpectedCellValue(sheetName, 15, 1, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 16, 1, "Sum of Total"),
					new ExpectedCellValue(sheetName, 17, 1, "20100090"),
					new ExpectedCellValue(sheetName, 18, 1, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 19, 1, "Sum of Total"),
					new ExpectedCellValue(sheetName, 20, 1, "Tent"),
					new ExpectedCellValue(sheetName, 21, 1, "20100070"),
					new ExpectedCellValue(sheetName, 22, 1, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 23, 1, "Sum of Total"),
					new ExpectedCellValue(sheetName, 24, 1, "San Francisco"),
					new ExpectedCellValue(sheetName, 25, 1, "Car Rack"),
					new ExpectedCellValue(sheetName, 26, 1, "20100076"),
					new ExpectedCellValue(sheetName, 27, 1, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 28, 1, "Sum of Total"),
					new ExpectedCellValue(sheetName, 29, 1, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 30, 1, "20100085"),
					new ExpectedCellValue(sheetName, 31, 1, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 32, 1, "Sum of Total"),
					new ExpectedCellValue(sheetName, 33, 1, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 34, 1, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 2, 2, "January"),
					new ExpectedCellValue(sheetName, 6, 2, 2d),
					new ExpectedCellValue(sheetName, 7, 2, 831.5),
					new ExpectedCellValue(sheetName, 18, 2, 2d),
					new ExpectedCellValue(sheetName, 19, 2, 831.5),
					new ExpectedCellValue(sheetName, 27, 2, 1d),
					new ExpectedCellValue(sheetName, 28, 2, 415.75),
					new ExpectedCellValue(sheetName, 33, 2, 5d),
					new ExpectedCellValue(sheetName, 34, 2, 2078.75),
					new ExpectedCellValue(sheetName, 2, 3, "February"),
					new ExpectedCellValue(sheetName, 22, 3, 6d),
					new ExpectedCellValue(sheetName, 23, 3, 1194d),
					new ExpectedCellValue(sheetName, 31, 3, 1d),
					new ExpectedCellValue(sheetName, 32, 3, 99d),
					new ExpectedCellValue(sheetName, 33, 3, 7d),
					new ExpectedCellValue(sheetName, 34, 3, 1293d),
					new ExpectedCellValue(sheetName, 2, 4, "March"),
					new ExpectedCellValue(sheetName, 10, 4, 1d),
					new ExpectedCellValue(sheetName, 11, 4, 24.99),
					new ExpectedCellValue(sheetName, 15, 4, 2d),
					new ExpectedCellValue(sheetName, 16, 4, 831.5),
					new ExpectedCellValue(sheetName, 33, 4, 3d),
					new ExpectedCellValue(sheetName, 34, 4, 856.49),
					new ExpectedCellValue(sheetName, 2, 5, "Grand Total"),
					new ExpectedCellValue(sheetName, 6, 5, 2d),
					new ExpectedCellValue(sheetName, 7, 5, 831.5),
					new ExpectedCellValue(sheetName, 10, 5, 1d),
					new ExpectedCellValue(sheetName, 11, 5, 24.99),
					new ExpectedCellValue(sheetName, 15, 5, 2d),
					new ExpectedCellValue(sheetName, 16, 5, 831.5),
					new ExpectedCellValue(sheetName, 18, 5, 2d),
					new ExpectedCellValue(sheetName, 19, 5, 831.5),
					new ExpectedCellValue(sheetName, 22, 5, 6d),
					new ExpectedCellValue(sheetName, 23, 5, 1194d),
					new ExpectedCellValue(sheetName, 27, 5, 1d),
					new ExpectedCellValue(sheetName, 28, 5, 415.75),
					new ExpectedCellValue(sheetName, 31, 5, 1d),
					new ExpectedCellValue(sheetName, 32, 5, 99d),
					new ExpectedCellValue(sheetName, 33, 5, 15d),
					new ExpectedCellValue(sheetName, 34, 5, 4228.24)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithMultipleRowDataFieldsThreeRowFieldOneColumnField.xlsx")]
		public void PivotTableRefreshMultipleRowDataFieldsThreeRowsAndOneColumnSubtotalsOn()
		{
			var file = new FileInfo("PivotTableWithMultipleRowDataFieldsThreeRowFieldOneColumnField.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["RowDataFields"];
					var pivotTable = worksheet.PivotTables["PivotTable1"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("A1:E52"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(8, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "RowDataFields";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 3, 1, "Chicago"),
					new ExpectedCellValue(sheetName, 4, 1, "Car Rack"),
					new ExpectedCellValue(sheetName, 5, 1, "20100007"),
					new ExpectedCellValue(sheetName, 6, 1, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 7, 1, "Sum of Total"),
					new ExpectedCellValue(sheetName, 8, 1, "Car Rack Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 9, 1, "Car Rack Sum of Total"),
					new ExpectedCellValue(sheetName, 10, 1, "Headlamp"),
					new ExpectedCellValue(sheetName, 11, 1, "20100083"),
					new ExpectedCellValue(sheetName, 12, 1, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 13, 1, "Sum of Total"),
					new ExpectedCellValue(sheetName, 14, 1, "Headlamp Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 15, 1, "Headlamp Sum of Total"),
					new ExpectedCellValue(sheetName, 16, 1, "Chicago Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 17, 1, "Chicago Sum of Total"),
					new ExpectedCellValue(sheetName, 18, 1, "Nashville"),
					new ExpectedCellValue(sheetName, 19, 1, "Car Rack"),
					new ExpectedCellValue(sheetName, 20, 1, "20100017"),
					new ExpectedCellValue(sheetName, 21, 1, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 22, 1, "Sum of Total"),
					new ExpectedCellValue(sheetName, 23, 1, "20100090"),
					new ExpectedCellValue(sheetName, 24, 1, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 25, 1, "Sum of Total"),
					new ExpectedCellValue(sheetName, 26, 1, "Car Rack Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 27, 1, "Car Rack Sum of Total"),
					new ExpectedCellValue(sheetName, 28, 1, "Tent"),
					new ExpectedCellValue(sheetName, 29, 1, "20100070"),
					new ExpectedCellValue(sheetName, 30, 1, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 31, 1, "Sum of Total"),
					new ExpectedCellValue(sheetName, 32, 1, "Tent Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 33, 1, "Tent Sum of Total"),
					new ExpectedCellValue(sheetName, 34, 1, "Nashville Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 35, 1, "Nashville Sum of Total"),
					new ExpectedCellValue(sheetName, 36, 1, "San Francisco"),
					new ExpectedCellValue(sheetName, 37, 1, "Car Rack"),
					new ExpectedCellValue(sheetName, 38, 1, "20100076"),
					new ExpectedCellValue(sheetName, 39, 1, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 40, 1, "Sum of Total"),
					new ExpectedCellValue(sheetName, 41, 1, "Car Rack Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 42, 1, "Car Rack Sum of Total"),
					new ExpectedCellValue(sheetName, 43, 1, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 44, 1, "20100085"),
					new ExpectedCellValue(sheetName, 45, 1, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 46, 1, "Sum of Total"),
					new ExpectedCellValue(sheetName, 47, 1, "Sleeping Bag Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 48, 1, "Sleeping Bag Sum of Total"),
					new ExpectedCellValue(sheetName, 49, 1, "San Francisco Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 50, 1, "San Francisco Sum of Total"),
					new ExpectedCellValue(sheetName, 51, 1, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 52, 1, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 2, 2, "January"),
					new ExpectedCellValue(sheetName, 6, 2, 2d),
					new ExpectedCellValue(sheetName, 7, 2, 831.5),
					new ExpectedCellValue(sheetName, 8, 2, 2d),
					new ExpectedCellValue(sheetName, 9, 2, 831.5),
					new ExpectedCellValue(sheetName, 16, 2, 2d),
					new ExpectedCellValue(sheetName, 17, 2, 831.5),
					new ExpectedCellValue(sheetName, 24, 2, 2d),
					new ExpectedCellValue(sheetName, 25, 2, 831.5),
					new ExpectedCellValue(sheetName, 26, 2, 2d),
					new ExpectedCellValue(sheetName, 27, 2, 831.5),
					new ExpectedCellValue(sheetName, 34, 2, 2d),
					new ExpectedCellValue(sheetName, 35, 2, 831.5),
					new ExpectedCellValue(sheetName, 39, 2, 1d),
					new ExpectedCellValue(sheetName, 40, 2, 415.75),
					new ExpectedCellValue(sheetName, 41, 2, 1d),
					new ExpectedCellValue(sheetName, 42, 2, 415.75),
					new ExpectedCellValue(sheetName, 49, 2, 1d),
					new ExpectedCellValue(sheetName, 50, 2, 415.75),
					new ExpectedCellValue(sheetName, 51, 2, 5d),
					new ExpectedCellValue(sheetName, 52, 2, 2078.75),
					new ExpectedCellValue(sheetName, 2, 3, "February"),
					new ExpectedCellValue(sheetName, 30, 3, 6d),
					new ExpectedCellValue(sheetName, 31, 3, 1194d),
					new ExpectedCellValue(sheetName, 32, 3, 6d),
					new ExpectedCellValue(sheetName, 33, 3, 1194d),
					new ExpectedCellValue(sheetName, 34, 3, 6d),
					new ExpectedCellValue(sheetName, 35, 3, 1194d),
					new ExpectedCellValue(sheetName, 45, 3, 1d),
					new ExpectedCellValue(sheetName, 46, 3, 99d),
					new ExpectedCellValue(sheetName, 47, 3, 1d),
					new ExpectedCellValue(sheetName, 48, 3, 99d),
					new ExpectedCellValue(sheetName, 49, 3, 1d),
					new ExpectedCellValue(sheetName, 50, 3, 99d),
					new ExpectedCellValue(sheetName, 51, 3, 7d),
					new ExpectedCellValue(sheetName, 52, 3, 1293d),
					new ExpectedCellValue(sheetName, 2, 4, "March"),
					new ExpectedCellValue(sheetName, 12, 4, 1d),
					new ExpectedCellValue(sheetName, 13, 4, 24.99),
					new ExpectedCellValue(sheetName, 14, 4, 1d),
					new ExpectedCellValue(sheetName, 15, 4, 24.99),
					new ExpectedCellValue(sheetName, 16, 4, 1d),
					new ExpectedCellValue(sheetName, 17, 4, 24.99),
					new ExpectedCellValue(sheetName, 21, 4, 2d),
					new ExpectedCellValue(sheetName, 22, 4, 831.5),
					new ExpectedCellValue(sheetName, 26, 4, 2d),
					new ExpectedCellValue(sheetName, 27, 4, 831.5),
					new ExpectedCellValue(sheetName, 34, 4, 2d),
					new ExpectedCellValue(sheetName, 35, 4, 831.5),
					new ExpectedCellValue(sheetName, 51, 4, 3d),
					new ExpectedCellValue(sheetName, 52, 4, 856.49),
					new ExpectedCellValue(sheetName, 2, 5, "Grand Total"),
					new ExpectedCellValue(sheetName, 6, 5, 2d),
					new ExpectedCellValue(sheetName, 7, 5, 831.5),
					new ExpectedCellValue(sheetName, 8, 5, 2d),
					new ExpectedCellValue(sheetName, 9, 5, 831.5),
					new ExpectedCellValue(sheetName, 12, 5, 1d),
					new ExpectedCellValue(sheetName, 13, 5, 24.99),
					new ExpectedCellValue(sheetName, 14, 5, 1d),
					new ExpectedCellValue(sheetName, 15, 5, 24.99),
					new ExpectedCellValue(sheetName, 16, 5, 3d),
					new ExpectedCellValue(sheetName, 17, 5, 856.49),
					new ExpectedCellValue(sheetName, 21, 5, 2d),
					new ExpectedCellValue(sheetName, 22, 5, 831.5),
					new ExpectedCellValue(sheetName, 24, 5, 2d),
					new ExpectedCellValue(sheetName, 25, 5, 831.5),
					new ExpectedCellValue(sheetName, 26, 5, 4d),
					new ExpectedCellValue(sheetName, 27, 5, 1663),
					new ExpectedCellValue(sheetName, 30, 5, 6d),
					new ExpectedCellValue(sheetName, 31, 5, 1194d),
					new ExpectedCellValue(sheetName, 32, 5, 6d),
					new ExpectedCellValue(sheetName, 33, 5, 1194d),
					new ExpectedCellValue(sheetName, 34, 5, 10d),
					new ExpectedCellValue(sheetName, 35, 5, 2857d),
					new ExpectedCellValue(sheetName, 39, 5, 1d),
					new ExpectedCellValue(sheetName, 40, 5, 415.75),
					new ExpectedCellValue(sheetName, 41, 5, 1d),
					new ExpectedCellValue(sheetName, 42, 5, 415.75),
					new ExpectedCellValue(sheetName, 45, 5, 1d),
					new ExpectedCellValue(sheetName, 46, 5, 99d),
					new ExpectedCellValue(sheetName, 47, 5, 1d),
					new ExpectedCellValue(sheetName, 48, 5, 99d),
					new ExpectedCellValue(sheetName, 49, 5, 2d),
					new ExpectedCellValue(sheetName, 50, 5, 514.75),
					new ExpectedCellValue(sheetName, 51, 5, 15d),
					new ExpectedCellValue(sheetName, 52, 5, 4228.24)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithMultipleRowDataFieldsNoColumnField.xlsx")]
		public void PivotTableRefreshMultipleRowDataFieldsOneRowAndNoColumns()
		{
			var file = new FileInfo("PivotTableWithMultipleRowDataFieldsNoColumnField.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["RowDataFields"];
					var pivotTable = worksheet.PivotTables["PivotTable1"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("A1:B12"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "RowDataFields";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 2, 1, "January"),
					new ExpectedCellValue(sheetName, 3, 1, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 4, 1, "Sum of Total"),
					new ExpectedCellValue(sheetName, 5, 1, "February"),
					new ExpectedCellValue(sheetName, 6, 1, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 7, 1, "Sum of Total"),
					new ExpectedCellValue(sheetName, 8, 1, "March"),
					new ExpectedCellValue(sheetName, 9, 1, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 10, 1, "Sum of Total"),
					new ExpectedCellValue(sheetName, 11, 1, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 12, 1, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 3, 2, 5d),
					new ExpectedCellValue(sheetName, 4, 2, 2078.75),
					new ExpectedCellValue(sheetName, 6, 2, 7d),
					new ExpectedCellValue(sheetName, 7, 2, 1293d),
					new ExpectedCellValue(sheetName, 9, 2, 3d),
					new ExpectedCellValue(sheetName, 10, 2, 856.49),
					new ExpectedCellValue(sheetName, 11, 2, 15d),
					new ExpectedCellValue(sheetName, 12, 2, 4228.24)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithMultipleRowDataFieldsNoColumnField.xlsx")]
		public void PivotTableRefreshMultipleRowDataFieldsTwoRowsAndNoColumnsSubtotalsOff()
		{
			var file = new FileInfo("PivotTableWithMultipleRowDataFieldsNoColumnField.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["RowDataFields"];
					var pivotTable = worksheet.PivotTables["PivotTable2"];
					foreach (var field in pivotTable.Fields)
					{
						field.DisableDefaultSubtotal();
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("F1:G26"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(3, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(3, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "RowDataFields";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 2, 6, "January"),
					new ExpectedCellValue(sheetName, 3, 6, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 4, 6, "Chicago"),
					new ExpectedCellValue(sheetName, 5, 6, "Nashville"),
					new ExpectedCellValue(sheetName, 6, 6, "San Francisco"),
					new ExpectedCellValue(sheetName, 7, 6, "Sum of Total"),
					new ExpectedCellValue(sheetName, 8, 6, "Chicago"),
					new ExpectedCellValue(sheetName, 9, 6, "Nashville"),
					new ExpectedCellValue(sheetName, 10, 6, "San Francisco"),
					new ExpectedCellValue(sheetName, 11, 6, "February"),
					new ExpectedCellValue(sheetName, 12, 6, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 13, 6, "Nashville"),
					new ExpectedCellValue(sheetName, 14, 6, "San Francisco"),
					new ExpectedCellValue(sheetName, 15, 6, "Sum of Total"),
					new ExpectedCellValue(sheetName, 16, 6, "Nashville"),
					new ExpectedCellValue(sheetName, 17, 6, "San Francisco"),
					new ExpectedCellValue(sheetName, 18, 6, "March"),
					new ExpectedCellValue(sheetName, 19, 6, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 20, 6, "Chicago"),
					new ExpectedCellValue(sheetName, 21, 6, "Nashville"),
					new ExpectedCellValue(sheetName, 22, 6, "Sum of Total"),
					new ExpectedCellValue(sheetName, 23, 6, "Chicago"),
					new ExpectedCellValue(sheetName, 24, 6, "Nashville"),
					new ExpectedCellValue(sheetName, 25, 6, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 26, 6, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 4, 7, 2d),
					new ExpectedCellValue(sheetName, 5, 7, 2d),
					new ExpectedCellValue(sheetName, 6, 7, 1d),
					new ExpectedCellValue(sheetName, 8, 7, 831.5),
					new ExpectedCellValue(sheetName, 9, 7, 831.5),
					new ExpectedCellValue(sheetName, 10, 7, 415.75),
					new ExpectedCellValue(sheetName, 13, 7, 6d),
					new ExpectedCellValue(sheetName, 14, 7, 1d),
					new ExpectedCellValue(sheetName, 16, 7, 1194d),
					new ExpectedCellValue(sheetName, 17, 7, 99d),
					new ExpectedCellValue(sheetName, 20, 7, 1d),
					new ExpectedCellValue(sheetName, 21, 7, 2d),
					new ExpectedCellValue(sheetName, 23, 7, 24.99),
					new ExpectedCellValue(sheetName, 24, 7, 831.5),
					new ExpectedCellValue(sheetName, 25, 7, 15d),
					new ExpectedCellValue(sheetName, 26, 7, 4228.24)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithMultipleRowDataFieldsNoColumnField.xlsx")]
		public void PivotTableRefreshMultipleRowDataFieldsTwoRowsAndNoColumnsSubtotalsOn()
		{
			var file = new FileInfo("PivotTableWithMultipleRowDataFieldsNoColumnField.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["RowDataFields"];
					var pivotTable = worksheet.PivotTables["PivotTable2"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("F1:G32"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "RowDataFields";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 2, 6, "January"),
					new ExpectedCellValue(sheetName, 3, 6, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 4, 6, "Chicago"),
					new ExpectedCellValue(sheetName, 5, 6, "Nashville"),
					new ExpectedCellValue(sheetName, 6, 6, "San Francisco"),
					new ExpectedCellValue(sheetName, 7, 6, "Sum of Total"),
					new ExpectedCellValue(sheetName, 8, 6, "Chicago"),
					new ExpectedCellValue(sheetName, 9, 6, "Nashville"),
					new ExpectedCellValue(sheetName, 10, 6, "San Francisco"),
					new ExpectedCellValue(sheetName, 11, 6, "January Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 12, 6, "January Sum of Total"),
					new ExpectedCellValue(sheetName, 13, 6, "February"),
					new ExpectedCellValue(sheetName, 14, 6, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 15, 6, "Nashville"),
					new ExpectedCellValue(sheetName, 16, 6, "San Francisco"),
					new ExpectedCellValue(sheetName, 17, 6, "Sum of Total"),
					new ExpectedCellValue(sheetName, 18, 6, "Nashville"),
					new ExpectedCellValue(sheetName, 19, 6, "San Francisco"),
					new ExpectedCellValue(sheetName, 20, 6, "February Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 21, 6, "February Sum of Total"),
					new ExpectedCellValue(sheetName, 22, 6, "March"),
					new ExpectedCellValue(sheetName, 23, 6, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 24, 6, "Chicago"),
					new ExpectedCellValue(sheetName, 25, 6, "Nashville"),
					new ExpectedCellValue(sheetName, 26, 6, "Sum of Total"),
					new ExpectedCellValue(sheetName, 27, 6, "Chicago"),
					new ExpectedCellValue(sheetName, 28, 6, "Nashville"),
					new ExpectedCellValue(sheetName, 29, 6, "March Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 30, 6, "March Sum of Total"),
					new ExpectedCellValue(sheetName, 31, 6, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 32, 6, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 4, 7, 2d),
					new ExpectedCellValue(sheetName, 5, 7, 2d),
					new ExpectedCellValue(sheetName, 6, 7, 1d),
					new ExpectedCellValue(sheetName, 8, 7, 831.5),
					new ExpectedCellValue(sheetName, 9, 7, 831.5),
					new ExpectedCellValue(sheetName, 10, 7, 415.75),
					new ExpectedCellValue(sheetName, 11, 7, 5d),
					new ExpectedCellValue(sheetName, 12, 7, 2078.75),
					new ExpectedCellValue(sheetName, 15, 7, 6d),
					new ExpectedCellValue(sheetName, 16, 7, 1d),
					new ExpectedCellValue(sheetName, 18, 7, 1194d),
					new ExpectedCellValue(sheetName, 19, 7, 99d),
					new ExpectedCellValue(sheetName, 20, 7, 7d),
					new ExpectedCellValue(sheetName, 21, 7, 1293d),
					new ExpectedCellValue(sheetName, 24, 7, 1d),
					new ExpectedCellValue(sheetName, 25, 7, 2d),
					new ExpectedCellValue(sheetName, 27, 7, 24.99),
					new ExpectedCellValue(sheetName, 28, 7, 831.5),
					new ExpectedCellValue(sheetName, 29, 7, 3d),
					new ExpectedCellValue(sheetName, 30, 7, 856.49),
					new ExpectedCellValue(sheetName, 31, 7, 15d),
					new ExpectedCellValue(sheetName, 32, 7, 4228.24)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithMultipleRowDataFieldsNoColumnField.xlsx")]
		public void PivotTableRefreshMultipleRowDataFieldsTwoRowsAndNoColumnsLastColumnDataField()
		{
			var file = new FileInfo("PivotTableWithMultipleRowDataFieldsNoColumnField.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["RowDataFields"];
					var pivotTable = worksheet.PivotTables["PivotTable3"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("K1:L33"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "RowDataFields";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 2, 11, "January"),
					new ExpectedCellValue(sheetName, 3, 11, "Chicago"),
					new ExpectedCellValue(sheetName, 4, 11, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 5, 11, "Sum of Total"),
					new ExpectedCellValue(sheetName, 6, 11, "Nashville"),
					new ExpectedCellValue(sheetName, 7, 11, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 8, 11, "Sum of Total"),
					new ExpectedCellValue(sheetName, 9, 11, "San Francisco"),
					new ExpectedCellValue(sheetName, 10, 11, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 11, 11, "Sum of Total"),
					new ExpectedCellValue(sheetName, 12, 11, "January Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 13, 11, "January Sum of Total"),
					new ExpectedCellValue(sheetName, 14, 11, "February"),
					new ExpectedCellValue(sheetName, 15, 11, "Nashville"),
					new ExpectedCellValue(sheetName, 16, 11, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 17, 11, "Sum of Total"),
					new ExpectedCellValue(sheetName, 18, 11, "San Francisco"),
					new ExpectedCellValue(sheetName, 19, 11, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 20, 11, "Sum of Total"),
					new ExpectedCellValue(sheetName, 21, 11, "February Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 22, 11, "February Sum of Total"),
					new ExpectedCellValue(sheetName, 23, 11, "March"),
					new ExpectedCellValue(sheetName, 24, 11, "Chicago"),
					new ExpectedCellValue(sheetName, 25, 11, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 26, 11, "Sum of Total"),
					new ExpectedCellValue(sheetName, 27, 11, "Nashville"),
					new ExpectedCellValue(sheetName, 28, 11, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 29, 11, "Sum of Total"),
					new ExpectedCellValue(sheetName, 30, 11, "March Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 31, 11, "March Sum of Total"),
					new ExpectedCellValue(sheetName, 32, 11, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 33, 11, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 4, 12, 2d),
					new ExpectedCellValue(sheetName, 5, 12, 831.5d),
					new ExpectedCellValue(sheetName, 7, 12, 2d),
					new ExpectedCellValue(sheetName, 8, 12, 831.5d),
					new ExpectedCellValue(sheetName, 10, 12, 1d),
					new ExpectedCellValue(sheetName, 11, 12, 415.75d),
					new ExpectedCellValue(sheetName, 12, 12, 5d),
					new ExpectedCellValue(sheetName, 13, 12, 2078.75),
					new ExpectedCellValue(sheetName, 16, 12, 6d),
					new ExpectedCellValue(sheetName, 17, 12, 1194d),
					new ExpectedCellValue(sheetName, 19, 12, 1d),
					new ExpectedCellValue(sheetName, 20, 12, 99d),
					new ExpectedCellValue(sheetName, 21, 12, 7d),
					new ExpectedCellValue(sheetName, 22, 12, 1293d),
					new ExpectedCellValue(sheetName, 25, 12, 1d),
					new ExpectedCellValue(sheetName, 26, 12, 24.99),
					new ExpectedCellValue(sheetName, 28, 12, 2d),
					new ExpectedCellValue(sheetName, 29, 12, 831.5),
					new ExpectedCellValue(sheetName, 30, 12, 3d),
					new ExpectedCellValue(sheetName, 31, 12, 856.49),
					new ExpectedCellValue(sheetName, 32, 12, 15),
					new ExpectedCellValue(sheetName, 33, 12, 4228.24)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithMultipleRowDataFieldsNoColumnField.xlsx")]
		public void PivotTableRefreshMultipleRowDataFieldsThreeRowsAndNoColumnsLastColumnDataField()
		{
			var file = new FileInfo("PivotTableWithMultipleRowDataFieldsNoColumnField.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["RowDataFields"];
					var pivotTable = worksheet.PivotTables["PivotTable4"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("P1:Q54"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(8, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "RowDataFields";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 2, 16, "January"),
					new ExpectedCellValue(sheetName, 3, 16, "Chicago"),
					new ExpectedCellValue(sheetName, 4, 16, "20100007"),
					new ExpectedCellValue(sheetName, 5, 16, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 6, 16, "Sum of Total"),
					new ExpectedCellValue(sheetName, 7, 16, "Chicago Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 8, 16, "Chicago Sum of Total"),
					new ExpectedCellValue(sheetName, 9, 16, "Nashville"),
					new ExpectedCellValue(sheetName, 10, 16, "20100090"),
					new ExpectedCellValue(sheetName, 11, 16, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 12, 16, "Sum of Total"),
					new ExpectedCellValue(sheetName, 13, 16, "Nashville Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 14, 16, "Nashville Sum of Total"),
					new ExpectedCellValue(sheetName, 15, 16, "San Francisco"),
					new ExpectedCellValue(sheetName, 16, 16, "20100076"),
					new ExpectedCellValue(sheetName, 17, 16, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 18, 16, "Sum of Total"),
					new ExpectedCellValue(sheetName, 19, 16, "San Francisco Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 20, 16, "San Francisco Sum of Total"),
					new ExpectedCellValue(sheetName, 21, 16, "January Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 22, 16, "January Sum of Total"),
					new ExpectedCellValue(sheetName, 23, 16, "February"),
					new ExpectedCellValue(sheetName, 24, 16, "Nashville"),
					new ExpectedCellValue(sheetName, 25, 16, "20100070"),
					new ExpectedCellValue(sheetName, 26, 16, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 27, 16, "Sum of Total"),
					new ExpectedCellValue(sheetName, 28, 16, "Nashville Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 29, 16, "Nashville Sum of Total"),
					new ExpectedCellValue(sheetName, 30, 16, "San Francisco"),
					new ExpectedCellValue(sheetName, 31, 16, "20100085"),
					new ExpectedCellValue(sheetName, 32, 16, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 33, 16, "Sum of Total"),
					new ExpectedCellValue(sheetName, 34, 16, "San Francisco Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 35, 16, "San Francisco Sum of Total"),
					new ExpectedCellValue(sheetName, 36, 16, "February Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 37, 16, "February Sum of Total"),
					new ExpectedCellValue(sheetName, 38, 16, "March"),
					new ExpectedCellValue(sheetName, 39, 16, "Chicago"),
					new ExpectedCellValue(sheetName, 40, 16, "20100083"),
					new ExpectedCellValue(sheetName, 41, 16, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 42, 16, "Sum of Total"),
					new ExpectedCellValue(sheetName, 43, 16, "Chicago Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 44, 16, "Chicago Sum of Total"),
					new ExpectedCellValue(sheetName, 45, 16, "Nashville"),
					new ExpectedCellValue(sheetName, 46, 16, "20100017"),
					new ExpectedCellValue(sheetName, 47, 16, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 48, 16, "Sum of Total"),
					new ExpectedCellValue(sheetName, 49, 16, "Nashville Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 50, 16, "Nashville Sum of Total"),
					new ExpectedCellValue(sheetName, 51, 16, "March Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 52, 16, "March Sum of Total"),
					new ExpectedCellValue(sheetName, 53, 16, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 54, 16, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 5, 17, 2d),
					new ExpectedCellValue(sheetName, 6, 17, 831.5d),
					new ExpectedCellValue(sheetName, 7, 17, 2d),
					new ExpectedCellValue(sheetName, 8, 17, 831.5d),
					new ExpectedCellValue(sheetName, 11, 17, 2d),
					new ExpectedCellValue(sheetName, 12, 17, 831.5d),
					new ExpectedCellValue(sheetName, 13, 17, 2d),
					new ExpectedCellValue(sheetName, 14, 17, 831.5d),
					new ExpectedCellValue(sheetName, 17, 17, 1d),
					new ExpectedCellValue(sheetName, 18, 17, 415.75d),
					new ExpectedCellValue(sheetName, 19, 17, 1d),
					new ExpectedCellValue(sheetName, 20, 17, 415.75d),
					new ExpectedCellValue(sheetName, 21, 17, 5d),
					new ExpectedCellValue(sheetName, 22, 17, 2078.75),
					new ExpectedCellValue(sheetName, 26, 17, 6d),
					new ExpectedCellValue(sheetName, 27, 17, 1194d),
					new ExpectedCellValue(sheetName, 28, 17, 6d),
					new ExpectedCellValue(sheetName, 29, 17, 1194d),
					new ExpectedCellValue(sheetName, 32, 17, 1d),
					new ExpectedCellValue(sheetName, 33, 17, 99d),
					new ExpectedCellValue(sheetName, 34, 17, 1d),
					new ExpectedCellValue(sheetName, 35, 17, 99d),
					new ExpectedCellValue(sheetName, 36, 17, 7d),
					new ExpectedCellValue(sheetName, 37, 17, 1293d),
					new ExpectedCellValue(sheetName, 41, 17, 1d),
					new ExpectedCellValue(sheetName, 42, 17, 24.99),
					new ExpectedCellValue(sheetName, 43, 17, 1d),
					new ExpectedCellValue(sheetName, 44, 17, 24.99),
					new ExpectedCellValue(sheetName, 47, 17, 2d),
					new ExpectedCellValue(sheetName, 48, 17, 831.5),
					new ExpectedCellValue(sheetName, 49, 17, 2d),
					new ExpectedCellValue(sheetName, 50, 17, 831.5),
					new ExpectedCellValue(sheetName, 51, 17, 3d),
					new ExpectedCellValue(sheetName, 52, 17, 856.49),
					new ExpectedCellValue(sheetName, 53, 17, 15),
					new ExpectedCellValue(sheetName, 54, 17, 4228.24)
				});
			}
		}

		#region Pivot Tables with page fields
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithPageFields.xlsx")]
		public void PivotTableRefreshWithPageFields()
		{
			var file = new FileInfo("PivotTableWithPageFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["Sheet1"];
					var pivotTable = worksheet.PivotTables["PivotTable1"];
					pivotTable.PageFields[0].Item = 1;
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("J9:M18"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(8, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "Sheet1";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 9, 10, null),
					new ExpectedCellValue(sheetName, 9, 11, "Column Labels"),
					new ExpectedCellValue(sheetName, 10, 10, "Row Labels"),
					new ExpectedCellValue(sheetName, 10, 11, "Nashville"),
					new ExpectedCellValue(sheetName, 10, 12, "San Francisco"),
					new ExpectedCellValue(sheetName, 10, 13, "Grand Total"),
					new ExpectedCellValue(sheetName, 11, 10, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 12, 10, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 12, 12, 99d),
					new ExpectedCellValue(sheetName, 12, 13, 99d),
					new ExpectedCellValue(sheetName, 13, 10, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 13, 12, 1d),
					new ExpectedCellValue(sheetName, 13, 13, 1d),
					new ExpectedCellValue(sheetName, 14, 10, "Tent"),
					new ExpectedCellValue(sheetName, 15, 10, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 15, 11, 199d),
					new ExpectedCellValue(sheetName, 15, 13, 199d),
					new ExpectedCellValue(sheetName, 16, 10, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 16, 11, 6d),
					new ExpectedCellValue(sheetName, 16, 13, 6d),
					new ExpectedCellValue(sheetName, 17, 10, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 17, 11, 199d),
					new ExpectedCellValue(sheetName, 17, 12, 99d),
					new ExpectedCellValue(sheetName, 17, 13, 298d),
					new ExpectedCellValue(sheetName, 18, 10, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 18, 11, 6d),
					new ExpectedCellValue(sheetName, 18, 12, 1d),
					new ExpectedCellValue(sheetName, 18, 13, 7d)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithPageFields.xlsx")]
		public void PivotTableRefreshWithPageFieldsMultipleFilters()
		{
			var file = new FileInfo("PivotTableWithPageFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["Sheet1"];
					var pivotTable = worksheet.PivotTables["PivotTable1"];
					pivotTable.PageFields[0].Item = 1;
					pivotTable.PageFields[1].Item = 4;
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("J9:L15"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(8, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "Sheet1";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 9, 10, null),
					new ExpectedCellValue(sheetName, 9, 11, "Column Labels"),
					new ExpectedCellValue(sheetName, 10, 10, "Row Labels"),
					new ExpectedCellValue(sheetName, 10, 11, "Nashville"),
					new ExpectedCellValue(sheetName, 10, 12, "Grand Total"),
					new ExpectedCellValue(sheetName, 11, 10, "Tent"),
					new ExpectedCellValue(sheetName, 12, 10, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 12, 11, 199d),
					new ExpectedCellValue(sheetName, 12, 12, 199d),
					new ExpectedCellValue(sheetName, 13, 10, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 13, 11, 6d),
					new ExpectedCellValue(sheetName, 13, 12, 6d),
					new ExpectedCellValue(sheetName, 14, 10, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 14, 11, 199d),
					new ExpectedCellValue(sheetName, 14, 12, 199d),
					new ExpectedCellValue(sheetName, 15, 10, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 15, 11, 6d),
					new ExpectedCellValue(sheetName, 15, 12, 6d),
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithPageFields.xlsx")]
		public void PivotTableRefreshWithMultiSelectPageFields()
		{
			var file = new FileInfo("PivotTableWithPageFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["Sheet1"];
					var pivotTable = worksheet.PivotTables["PivotTable1"];
					pivotTable.Fields[2].Items[1].Hidden = true;
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("J9:N18"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(8, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "Sheet1";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 9, 10, null),
					new ExpectedCellValue(sheetName, 9, 11, "Column Labels"),
					new ExpectedCellValue(sheetName, 10, 10, "Row Labels"),
					new ExpectedCellValue(sheetName, 10, 11, "Chicago"),
					new ExpectedCellValue(sheetName, 10, 12, "Nashville"),
					new ExpectedCellValue(sheetName, 10, 13, "San Francisco"),
					new ExpectedCellValue(sheetName, 10, 14, "Grand Total"),
					new ExpectedCellValue(sheetName, 11, 10, "Car Rack"),
					new ExpectedCellValue(sheetName, 12, 10, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 12, 11, 415.75),
					new ExpectedCellValue(sheetName, 12, 12, 831.5),
					new ExpectedCellValue(sheetName, 12, 13, 415.75),
					new ExpectedCellValue(sheetName, 12, 14, 1663d),
					new ExpectedCellValue(sheetName, 13, 10, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 13, 11, 2d),
					new ExpectedCellValue(sheetName, 13, 12, 4d),
					new ExpectedCellValue(sheetName, 13, 13, 1d),
					new ExpectedCellValue(sheetName, 13, 14, 7d),
					new ExpectedCellValue(sheetName, 14, 10, "Headlamp"),
					new ExpectedCellValue(sheetName, 15, 10, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 15, 11, 24.99),
					new ExpectedCellValue(sheetName, 15, 14, 24.99),
					new ExpectedCellValue(sheetName, 16, 10, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 16, 11, 1d),
					new ExpectedCellValue(sheetName, 16, 14, 1d),
					new ExpectedCellValue(sheetName, 17, 10, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 17, 11, 440.74),
					new ExpectedCellValue(sheetName, 17, 12, 831.5),
					new ExpectedCellValue(sheetName, 17, 13, 415.75),
					new ExpectedCellValue(sheetName, 17, 14, 1687.99),
					new ExpectedCellValue(sheetName, 18, 10, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 18, 11, 3d),
					new ExpectedCellValue(sheetName, 18, 12, 4d),
					new ExpectedCellValue(sheetName, 18, 13, 1d),
					new ExpectedCellValue(sheetName, 18, 14, 8d)
				});
			}
		}
		#endregion
		#endregion

		#region Multiple Column Data Fields
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithMultipleColumnDataFields.xlsx")]
		public void PivotTableRefreshMultipleColumnDataFieldsAtLeafNode()
		{
			var file = new FileInfo("PivotTableWithMultipleColumnDataFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["ColumnDataFields"];
					var pivotTable = worksheet.PivotTables["PivotTable1"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("A1:I7"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "ColumnDataFields";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 4, 1, "Chicago"),
					new ExpectedCellValue(sheetName, 5, 1, "Nashville"),
					new ExpectedCellValue(sheetName, 6, 1, "San Francisco"),
					new ExpectedCellValue(sheetName, 7, 1, "Grand Total"),
					new ExpectedCellValue(sheetName, 2, 2, "January"),
					new ExpectedCellValue(sheetName, 3, 2, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 4, 2, 2d),
					new ExpectedCellValue(sheetName, 5, 2, 2d),
					new ExpectedCellValue(sheetName, 6, 2, 1d),
					new ExpectedCellValue(sheetName, 7, 2, 5d),
					new ExpectedCellValue(sheetName, 3, 3, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 4, 3, 415.75),
					new ExpectedCellValue(sheetName, 5, 3, 415.75),
					new ExpectedCellValue(sheetName, 6, 3, 415.75),
					new ExpectedCellValue(sheetName, 7, 3, 1247.25),
					new ExpectedCellValue(sheetName, 2, 4, "February"),
					new ExpectedCellValue(sheetName, 3, 4, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 5, 4, 6d),
					new ExpectedCellValue(sheetName, 6, 4, 1d),
					new ExpectedCellValue(sheetName, 7, 4, 7d),
					new ExpectedCellValue(sheetName, 3, 5, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 5, 5, 199d),
					new ExpectedCellValue(sheetName, 6, 5, 99d),
					new ExpectedCellValue(sheetName, 7, 5, 298d),
					new ExpectedCellValue(sheetName, 2, 6, "March"),
					new ExpectedCellValue(sheetName, 3, 6, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 4, 6, 1d),
					new ExpectedCellValue(sheetName, 5, 6, 2d),
					new ExpectedCellValue(sheetName, 7, 6, 3d),
					new ExpectedCellValue(sheetName, 3, 7, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 4, 7, 24.99),
					new ExpectedCellValue(sheetName, 5, 7, 415.75),
					new ExpectedCellValue(sheetName, 7, 7, 440.74),
					new ExpectedCellValue(sheetName, 2, 8, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 4, 8, 3d),
					new ExpectedCellValue(sheetName, 5, 8, 10d),
					new ExpectedCellValue(sheetName, 6, 8, 2d),
					new ExpectedCellValue(sheetName, 7, 8, 15d),
					new ExpectedCellValue(sheetName, 2, 9, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 4, 9, 440.74),
					new ExpectedCellValue(sheetName, 5, 9, 1030.5),
					new ExpectedCellValue(sheetName, 6, 9, 514.75),
					new ExpectedCellValue(sheetName, 7, 9, 1985.99),
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithMultipleColumnDataFields.xlsx")]
		public void PivotTableRefreshMultipleColumnDataFieldsAsParentRowDepthTwo()
		{
			var file = new FileInfo("PivotTableWithMultipleColumnDataFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["ColumnDataFields"];
					var pivotTable = worksheet.PivotTables["PivotTable2"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("A11:I17"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "ColumnDataFields";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 14, 1, "Chicago"),
					new ExpectedCellValue(sheetName, 15, 1, "Nashville"),
					new ExpectedCellValue(sheetName, 16, 1, "San Francisco"),
					new ExpectedCellValue(sheetName, 17, 1, "Grand Total"),
					new ExpectedCellValue(sheetName, 12, 2, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 13, 2, "January"),
					new ExpectedCellValue(sheetName, 14, 2, 2d),
					new ExpectedCellValue(sheetName, 15, 2, 2d),
					new ExpectedCellValue(sheetName, 16, 2, 1d),
					new ExpectedCellValue(sheetName, 17, 2, 5d),
					new ExpectedCellValue(sheetName, 13, 3, "February"),
					new ExpectedCellValue(sheetName, 15, 3, 6d),
					new ExpectedCellValue(sheetName, 16, 3, 1d),
					new ExpectedCellValue(sheetName, 17, 3, 7d),
					new ExpectedCellValue(sheetName, 13, 4, "March"),
					new ExpectedCellValue(sheetName, 14, 4, 1d),
					new ExpectedCellValue(sheetName, 15, 4, 2d),
					new ExpectedCellValue(sheetName, 17, 4, 3d),
					new ExpectedCellValue(sheetName, 12, 5, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 13, 5, "January"),
					new ExpectedCellValue(sheetName, 14, 5, 415.75),
					new ExpectedCellValue(sheetName, 15, 5, 415.75),
					new ExpectedCellValue(sheetName, 16, 5, 415.75),
					new ExpectedCellValue(sheetName, 17, 5, 1247.25),
					new ExpectedCellValue(sheetName, 13, 6, "February"),
					new ExpectedCellValue(sheetName, 15, 6, 199d),
					new ExpectedCellValue(sheetName, 16, 6, 99d),
					new ExpectedCellValue(sheetName, 17, 6, 298d),
					new ExpectedCellValue(sheetName, 13, 7, "March"),
					new ExpectedCellValue(sheetName, 14, 7, 24.99),
					new ExpectedCellValue(sheetName, 15, 7, 415.75),
					new ExpectedCellValue(sheetName, 17, 7, 440.74),
					new ExpectedCellValue(sheetName, 12, 8, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 14, 8, 3d),
					new ExpectedCellValue(sheetName, 15, 8, 10d),
					new ExpectedCellValue(sheetName, 16, 8, 2d),
					new ExpectedCellValue(sheetName, 17, 8, 15d),
					new ExpectedCellValue(sheetName, 12, 9, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 14, 9, 440.74),
					new ExpectedCellValue(sheetName, 15, 9, 1030.5),
					new ExpectedCellValue(sheetName, 16, 9, 514.75),
					new ExpectedCellValue(sheetName, 17, 9, 1985.99),
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithMultipleColumnDataFields.xlsx")]
		public void PivotTableRefreshMultipleColumnDataFieldsAsParentNodeColumnDepthThreeSubtotalsOn()
		{
			var file = new FileInfo("PivotTableWithMultipleColumnDataFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["ColumnDataFields"];
					var pivotTable = worksheet.PivotTables["PivotTable3"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("A21:W29"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "ColumnDataFields";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 25, 1, "Car Rack"),
					new ExpectedCellValue(sheetName, 26, 1, "Headlamp"),
					new ExpectedCellValue(sheetName, 27, 1, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 28, 1, "Tent"),
					new ExpectedCellValue(sheetName, 29, 1, "Grand Total"),
					new ExpectedCellValue(sheetName, 22, 2, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 23, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 24, 2, "January"),
					new ExpectedCellValue(sheetName, 25, 2, 2d),
					new ExpectedCellValue(sheetName, 29, 2, 2d),
					new ExpectedCellValue(sheetName, 24, 3, "March"),
					new ExpectedCellValue(sheetName, 26, 3, 1d),
					new ExpectedCellValue(sheetName, 29, 3, 1d),
					new ExpectedCellValue(sheetName, 23, 4, "Chicago Total"),
					new ExpectedCellValue(sheetName, 25, 4, 2d),
					new ExpectedCellValue(sheetName, 26, 4, 1d),
					new ExpectedCellValue(sheetName, 29, 4, 3d),
					new ExpectedCellValue(sheetName, 23, 5, "Nashville"),
					new ExpectedCellValue(sheetName, 24, 5, "January"),
					new ExpectedCellValue(sheetName, 25, 5, 2d),
					new ExpectedCellValue(sheetName, 29, 5, 2d),
					new ExpectedCellValue(sheetName, 24, 6, "February"),
					new ExpectedCellValue(sheetName, 28, 6, 6d),
					new ExpectedCellValue(sheetName, 29, 6, 6d),
					new ExpectedCellValue(sheetName, 24, 7, "March"),
					new ExpectedCellValue(sheetName, 25, 7, 2d),
					new ExpectedCellValue(sheetName, 29, 7, 2d),
					new ExpectedCellValue(sheetName, 23, 8, "Nashville Total"),
					new ExpectedCellValue(sheetName, 25, 8, 4d),
					new ExpectedCellValue(sheetName, 28, 8, 6d),
					new ExpectedCellValue(sheetName, 29, 8, 10d),
					new ExpectedCellValue(sheetName, 23, 9, "San Francisco"),
					new ExpectedCellValue(sheetName, 24, 9, "January"),
					new ExpectedCellValue(sheetName, 25, 9, 1d),
					new ExpectedCellValue(sheetName, 29, 9, 1d),
					new ExpectedCellValue(sheetName, 24, 10, "February"),
					new ExpectedCellValue(sheetName, 27, 10, 1d),
					new ExpectedCellValue(sheetName, 29, 10, 1d),
					new ExpectedCellValue(sheetName, 23, 11, "San Francisco Total"),
					new ExpectedCellValue(sheetName, 25, 11, 1d),
					new ExpectedCellValue(sheetName, 27, 11, 1d),
					new ExpectedCellValue(sheetName, 29, 11, 2d),
					new ExpectedCellValue(sheetName, 22, 12, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 23, 12, "Chicago"),
					new ExpectedCellValue(sheetName, 24, 12, "January"),
					new ExpectedCellValue(sheetName, 25, 12, 415.75),
					new ExpectedCellValue(sheetName, 29, 12, 415.75),
					new ExpectedCellValue(sheetName, 24, 13, "March"),
					new ExpectedCellValue(sheetName, 26, 13, 24.99),
					new ExpectedCellValue(sheetName, 29, 13, 24.99),
					new ExpectedCellValue(sheetName, 23, 14, "Chicago Total"),
					new ExpectedCellValue(sheetName, 25, 14, 415.75),
					new ExpectedCellValue(sheetName, 26, 14, 24.99),
					new ExpectedCellValue(sheetName, 29, 14, 440.74),
					new ExpectedCellValue(sheetName, 23, 15, "Nashville"),
					new ExpectedCellValue(sheetName, 24, 15, "January"),
					new ExpectedCellValue(sheetName, 25, 15, 415.75),
					new ExpectedCellValue(sheetName, 29, 15, 415.75),
					new ExpectedCellValue(sheetName, 24, 16, "February"),
					new ExpectedCellValue(sheetName, 28, 16, 199d),
					new ExpectedCellValue(sheetName, 29, 16, 199d),
					new ExpectedCellValue(sheetName, 24, 17, "March"),
					new ExpectedCellValue(sheetName, 25, 17, 415.75),
					new ExpectedCellValue(sheetName, 29, 17, 415.75),
					new ExpectedCellValue(sheetName, 23, 18, "Nashville Total"),
					new ExpectedCellValue(sheetName, 25, 18, 831.5),
					new ExpectedCellValue(sheetName, 28, 18, 199d),
					new ExpectedCellValue(sheetName, 29, 18, 1030.5),
					new ExpectedCellValue(sheetName, 23, 19, "San Francisco"),
					new ExpectedCellValue(sheetName, 24, 19, "January"),
					new ExpectedCellValue(sheetName, 25, 19, 415.75),
					new ExpectedCellValue(sheetName, 29, 19, 415.75),
					new ExpectedCellValue(sheetName, 24, 20, "February"),
					new ExpectedCellValue(sheetName, 27, 20, 99d),
					new ExpectedCellValue(sheetName, 29, 20, 99d),
					new ExpectedCellValue(sheetName, 23, 21, "San Francisco Total"),
					new ExpectedCellValue(sheetName, 25, 21, 415.75),
					new ExpectedCellValue(sheetName, 27, 21, 99d),
					new ExpectedCellValue(sheetName, 29, 21, 514.75),
					new ExpectedCellValue(sheetName, 22, 22, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 25, 22, 7d),
					new ExpectedCellValue(sheetName, 26, 22, 1d),
					new ExpectedCellValue(sheetName, 27, 22, 1d),
					new ExpectedCellValue(sheetName, 28, 22, 6d),
					new ExpectedCellValue(sheetName, 29, 22, 15d),
					new ExpectedCellValue(sheetName, 22, 23, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 25, 23, 1663d),
					new ExpectedCellValue(sheetName, 26, 23, 24.99),
					new ExpectedCellValue(sheetName, 27, 23, 99d),
					new ExpectedCellValue(sheetName, 28, 23, 199d),
					new ExpectedCellValue(sheetName, 29, 23, 1985.99)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithMultipleColumnDataFields.xlsx")]
		public void PivotTableRefreshMultipleColumnDataFieldsAsParentNodeColumnDepthThreeSubtotalsOff()
		{
			var file = new FileInfo("PivotTableWithMultipleColumnDataFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["ColumnDataFields"];
					var pivotTable = worksheet.PivotTables["PivotTable3"];
					foreach (var field in pivotTable.Fields)
					{
						field.DisableDefaultSubtotal();
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("A21:Q29"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(3, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(3, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "ColumnDataFields";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 25, 1, "Car Rack"),
					new ExpectedCellValue(sheetName, 26, 1, "Headlamp"),
					new ExpectedCellValue(sheetName, 27, 1, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 28, 1, "Tent"),
					new ExpectedCellValue(sheetName, 29, 1, "Grand Total"),
					new ExpectedCellValue(sheetName, 22, 2, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 23, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 24, 2, "January"),
					new ExpectedCellValue(sheetName, 25, 2, 2d),
					new ExpectedCellValue(sheetName, 29, 2, 2d),
					new ExpectedCellValue(sheetName, 24, 3, "March"),
					new ExpectedCellValue(sheetName, 26, 3, 1d),
					new ExpectedCellValue(sheetName, 29, 3, 1d),
					new ExpectedCellValue(sheetName, 23, 4, "Nashville"),
					new ExpectedCellValue(sheetName, 24, 4, "January"),
					new ExpectedCellValue(sheetName, 25, 4, 2d),
					new ExpectedCellValue(sheetName, 29, 4, 2d),
					new ExpectedCellValue(sheetName, 24, 5, "February"),
					new ExpectedCellValue(sheetName, 28, 5, 6d),
					new ExpectedCellValue(sheetName, 29, 5, 6d),
					new ExpectedCellValue(sheetName, 24, 6, "March"),
					new ExpectedCellValue(sheetName, 25, 6, 2d),
					new ExpectedCellValue(sheetName, 29, 6, 2d),
					new ExpectedCellValue(sheetName, 23, 7, "San Francisco"),
					new ExpectedCellValue(sheetName, 24, 7, "January"),
					new ExpectedCellValue(sheetName, 25, 7, 1d),
					new ExpectedCellValue(sheetName, 29, 7, 1d),
					new ExpectedCellValue(sheetName, 24, 8, "February"),
					new ExpectedCellValue(sheetName, 27, 8, 1d),
					new ExpectedCellValue(sheetName, 29, 8, 1d),
					new ExpectedCellValue(sheetName, 22, 9, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 23, 9, "Chicago"),
					new ExpectedCellValue(sheetName, 24, 9, "January"),
					new ExpectedCellValue(sheetName, 25, 9, 415.75),
					new ExpectedCellValue(sheetName, 29, 9, 415.75),
					new ExpectedCellValue(sheetName, 24, 10, "March"),
					new ExpectedCellValue(sheetName, 26, 10, 24.99),
					new ExpectedCellValue(sheetName, 29, 10, 24.99),
					new ExpectedCellValue(sheetName, 23, 11, "Nashville"),
					new ExpectedCellValue(sheetName, 24, 11, "January"),
					new ExpectedCellValue(sheetName, 25, 11, 415.75),
					new ExpectedCellValue(sheetName, 29, 11, 415.75),
					new ExpectedCellValue(sheetName, 24, 12, "February"),
					new ExpectedCellValue(sheetName, 28, 12, 199d),
					new ExpectedCellValue(sheetName, 29, 12, 199d),
					new ExpectedCellValue(sheetName, 24, 13, "March"),
					new ExpectedCellValue(sheetName, 25, 13, 415.75),
					new ExpectedCellValue(sheetName, 29, 13, 415.75),
					new ExpectedCellValue(sheetName, 23, 14, "San Francisco"),
					new ExpectedCellValue(sheetName, 24, 14, "January"),
					new ExpectedCellValue(sheetName, 25, 14, 415.75),
					new ExpectedCellValue(sheetName, 29, 14, 415.75),
					new ExpectedCellValue(sheetName, 24, 15, "February"),
					new ExpectedCellValue(sheetName, 27, 15, 99d),
					new ExpectedCellValue(sheetName, 29, 15, 99d),
					new ExpectedCellValue(sheetName, 22, 16, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 25, 16, 7d),
					new ExpectedCellValue(sheetName, 26, 16, 1d),
					new ExpectedCellValue(sheetName, 27, 16, 1d),
					new ExpectedCellValue(sheetName, 28, 16, 6d),
					new ExpectedCellValue(sheetName, 29, 16, 15d),
					new ExpectedCellValue(sheetName, 22, 17, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 25, 17, 1663d),
					new ExpectedCellValue(sheetName, 26, 17, 24.99),
					new ExpectedCellValue(sheetName, 27, 17, 99d),
					new ExpectedCellValue(sheetName, 28, 17, 199d),
					new ExpectedCellValue(sheetName, 29, 17, 1985.99),
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithMultipleColumnDataFields.xlsx")]
		public void PivotTableRefreshMultipleColumnDataFieldsAsInnerChildSubtotalsOn()
		{
			var file = new FileInfo("PivotTableWithMultipleColumnDataFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["ColumnDataFields"];
					var pivotTable = worksheet.PivotTables["PivotTable4"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("A33:S40"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "ColumnDataFields";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 37, 1, "Chicago"),
					new ExpectedCellValue(sheetName, 38, 1, "Nashville"),
					new ExpectedCellValue(sheetName, 39, 1, "San Francisco"),
					new ExpectedCellValue(sheetName, 40, 1, "Grand Total"),
					new ExpectedCellValue(sheetName, 34, 2, "January"),
					new ExpectedCellValue(sheetName, 35, 2, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 36, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 37, 2, 2d),
					new ExpectedCellValue(sheetName, 38, 2, 2d),
					new ExpectedCellValue(sheetName, 39, 2, 1d),
					new ExpectedCellValue(sheetName, 40, 2, 5d),
					new ExpectedCellValue(sheetName, 35, 3, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 36, 3, "Car Rack"),
					new ExpectedCellValue(sheetName, 37, 3, 415.75),
					new ExpectedCellValue(sheetName, 38, 3, 415.75),
					new ExpectedCellValue(sheetName, 39, 3, 415.75),
					new ExpectedCellValue(sheetName, 40, 3, 1247.25),
					new ExpectedCellValue(sheetName, 34, 4, "January Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 37, 4, 2d),
					new ExpectedCellValue(sheetName, 38, 4, 2d),
					new ExpectedCellValue(sheetName, 39, 4, 1d),
					new ExpectedCellValue(sheetName, 40, 4, 5d),
					new ExpectedCellValue(sheetName, 34, 5, "January Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 37, 5, 415.75),
					new ExpectedCellValue(sheetName, 38, 5, 415.75),
					new ExpectedCellValue(sheetName, 39, 5, 415.75),
					new ExpectedCellValue(sheetName, 40, 5, 1247.25),
					new ExpectedCellValue(sheetName, 34, 6, "February"),
					new ExpectedCellValue(sheetName, 35, 6, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 36, 6, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 39, 6, 1d),
					new ExpectedCellValue(sheetName, 40, 6, 1d),
					new ExpectedCellValue(sheetName, 36, 7, "Tent"),
					new ExpectedCellValue(sheetName, 38, 7, 6d),
					new ExpectedCellValue(sheetName, 40, 7, 6d),
					new ExpectedCellValue(sheetName, 35, 8, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 36, 8, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 39, 8, 99d),
					new ExpectedCellValue(sheetName, 40, 8, 99d),
					new ExpectedCellValue(sheetName, 36, 9, "Tent"),
					new ExpectedCellValue(sheetName, 38, 9, 199d),
					new ExpectedCellValue(sheetName, 40, 9, 199d),
					new ExpectedCellValue(sheetName, 34, 10, "February Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 38, 10, 6d),
					new ExpectedCellValue(sheetName, 39, 10, 1d),
					new ExpectedCellValue(sheetName, 40, 10, 7d),
					new ExpectedCellValue(sheetName, 34, 11, "February Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 38, 11, 199d),
					new ExpectedCellValue(sheetName, 39, 11, 99d),
					new ExpectedCellValue(sheetName, 40, 11, 298d),
					new ExpectedCellValue(sheetName, 34, 12, "March"),
					new ExpectedCellValue(sheetName, 35, 12, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 36, 12, "Car Rack"),
					new ExpectedCellValue(sheetName, 38, 12, 2d),
					new ExpectedCellValue(sheetName, 40, 12, 2d),
					new ExpectedCellValue(sheetName, 36, 13, "Headlamp"),
					new ExpectedCellValue(sheetName, 37, 13, 1d),
					new ExpectedCellValue(sheetName, 40, 13, 1d),
					new ExpectedCellValue(sheetName, 35, 14, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 36, 14, "Car Rack"),
					new ExpectedCellValue(sheetName, 38, 14, 415.75),
					new ExpectedCellValue(sheetName, 40, 14, 415.75),
					new ExpectedCellValue(sheetName, 36, 15, "Headlamp"),
					new ExpectedCellValue(sheetName, 37, 15, 24.99),
					new ExpectedCellValue(sheetName, 40, 15, 24.99),
					new ExpectedCellValue(sheetName, 34, 16, "March Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 37, 16, 1d),
					new ExpectedCellValue(sheetName, 38, 16, 2d),
					new ExpectedCellValue(sheetName, 40, 16, 3d),
					new ExpectedCellValue(sheetName, 34, 17, "March Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 37, 17, 24.99),
					new ExpectedCellValue(sheetName, 38, 17, 415.75),
					new ExpectedCellValue(sheetName, 40, 17, 440.74),
					new ExpectedCellValue(sheetName, 34, 18, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 37, 18, 3d),
					new ExpectedCellValue(sheetName, 38, 18, 10d),
					new ExpectedCellValue(sheetName, 39, 18, 2d),
					new ExpectedCellValue(sheetName, 40, 18, 15d),
					new ExpectedCellValue(sheetName, 34, 19, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 37, 19, 440.74),
					new ExpectedCellValue(sheetName, 38, 19, 1030.5),
					new ExpectedCellValue(sheetName, 39, 19, 514.75),
					new ExpectedCellValue(sheetName, 40, 19, 1985.99)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithMultipleColumnDataFields.xlsx")]
		public void PivotTableRefreshMultipleColumnDataFieldsAsInnerChildSubtotalsOff()
		{
			var file = new FileInfo("PivotTableWithMultipleColumnDataFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["ColumnDataFields"];
					var pivotTable = worksheet.PivotTables["PivotTable4"];
					foreach (var field in pivotTable.Fields)
					{
						field.DisableDefaultSubtotal();
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("A33:M40"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(3, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(3, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "ColumnDataFields";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 37, 1, "Chicago"),
					new ExpectedCellValue(sheetName, 38, 1, "Nashville"),
					new ExpectedCellValue(sheetName, 39, 1, "San Francisco"),
					new ExpectedCellValue(sheetName, 40, 1, "Grand Total"),
					new ExpectedCellValue(sheetName, 34, 2, "January"),
					new ExpectedCellValue(sheetName, 35, 2, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 36, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 37, 2, 2d),
					new ExpectedCellValue(sheetName, 38, 2, 2d),
					new ExpectedCellValue(sheetName, 39, 2, 1d),
					new ExpectedCellValue(sheetName, 40, 2, 5d),
					new ExpectedCellValue(sheetName, 35, 3, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 36, 3, "Car Rack"),
					new ExpectedCellValue(sheetName, 37, 3, 415.75),
					new ExpectedCellValue(sheetName, 38, 3, 415.75),
					new ExpectedCellValue(sheetName, 39, 3, 415.75),
					new ExpectedCellValue(sheetName, 40, 3, 1247.25),
					new ExpectedCellValue(sheetName, 34, 4, "February"),
					new ExpectedCellValue(sheetName, 35, 4, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 36, 4, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 39, 4, 1d),
					new ExpectedCellValue(sheetName, 40, 4, 1d),
					new ExpectedCellValue(sheetName, 36, 5, "Tent"),
					new ExpectedCellValue(sheetName, 38, 5, 6d),
					new ExpectedCellValue(sheetName, 40, 5, 6d),
					new ExpectedCellValue(sheetName, 35, 6, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 36, 6, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 39, 6, 99d),
					new ExpectedCellValue(sheetName, 40, 6, 99d),
					new ExpectedCellValue(sheetName, 36, 7, "Tent"),
					new ExpectedCellValue(sheetName, 38, 7, 199d),
					new ExpectedCellValue(sheetName, 40, 7, 199d),
					new ExpectedCellValue(sheetName, 34, 8, "March"),
					new ExpectedCellValue(sheetName, 35, 8, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 36, 8, "Car Rack"),
					new ExpectedCellValue(sheetName, 38, 8, 2d),
					new ExpectedCellValue(sheetName, 40, 8, 2d),
					new ExpectedCellValue(sheetName, 36, 9, "Headlamp"),
					new ExpectedCellValue(sheetName, 37, 9, 1d),
					new ExpectedCellValue(sheetName, 40, 9, 1d),
					new ExpectedCellValue(sheetName, 35, 10, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 36, 10, "Car Rack"),
					new ExpectedCellValue(sheetName, 38, 10, 415.75),
					new ExpectedCellValue(sheetName, 40, 10, 415.75),
					new ExpectedCellValue(sheetName, 36, 11, "Headlamp"),
					new ExpectedCellValue(sheetName, 37, 11, 24.99),
					new ExpectedCellValue(sheetName, 40, 11, 24.99),
					new ExpectedCellValue(sheetName, 34, 12, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 37, 12, 3d),
					new ExpectedCellValue(sheetName, 38, 12, 10d),
					new ExpectedCellValue(sheetName, 39, 12, 2d),
					new ExpectedCellValue(sheetName, 40, 12, 15d),
					new ExpectedCellValue(sheetName, 34, 13, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 37, 13, 440.74),
					new ExpectedCellValue(sheetName, 38, 13, 1030.5),
					new ExpectedCellValue(sheetName, 39, 13, 514.75),
					new ExpectedCellValue(sheetName, 40, 13, 1985.99)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithMultipleColumnDataFields.xlsx")]
		public void PivotTableRefreshMultipleColumnDataFieldsAsFirstInnerChildSubtotalsOn()
		{
			var file = new FileInfo("PivotTableWithMultipleColumnDataFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["ColumnDataFields"];
					var pivotTable = worksheet.PivotTables["PivotTable5"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("A44:AK56"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(8, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "ColumnDataFields";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 49, 1, 20100007),
					new ExpectedCellValue(sheetName, 50, 1, 20100017),
					new ExpectedCellValue(sheetName, 51, 1, 20100070),
					new ExpectedCellValue(sheetName, 52, 1, 20100076),
					new ExpectedCellValue(sheetName, 53, 1, 20100083),
					new ExpectedCellValue(sheetName, 54, 1, 20100085),
					new ExpectedCellValue(sheetName, 55, 1, 20100090),
					new ExpectedCellValue(sheetName, 56, 1, "Grand Total"),
					new ExpectedCellValue(sheetName, 45, 2, "January"),
					new ExpectedCellValue(sheetName, 46, 2, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 47, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 48, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 49, 2, 2d),
					new ExpectedCellValue(sheetName, 56, 2, 2d),
					new ExpectedCellValue(sheetName, 47, 3, "Chicago Total"),
					new ExpectedCellValue(sheetName, 49, 3, 2d),
					new ExpectedCellValue(sheetName, 56, 3, 2d),
					new ExpectedCellValue(sheetName, 47, 4, "Nashville"),
					new ExpectedCellValue(sheetName, 48, 4, "Car Rack"),
					new ExpectedCellValue(sheetName, 55, 4, 2d),
					new ExpectedCellValue(sheetName, 56, 4, 2d),
					new ExpectedCellValue(sheetName, 47, 5, "Nashville Total"),
					new ExpectedCellValue(sheetName, 55, 5, 2d),
					new ExpectedCellValue(sheetName, 56, 5, 2d),
					new ExpectedCellValue(sheetName, 47, 6, "San Francisco"),
					new ExpectedCellValue(sheetName, 48, 6, "Car Rack"),
					new ExpectedCellValue(sheetName, 52, 6, 1d),
					new ExpectedCellValue(sheetName, 56, 6, 1d),
					new ExpectedCellValue(sheetName, 47, 7, "San Francisco Total"),
					new ExpectedCellValue(sheetName, 52, 7, 1d),
					new ExpectedCellValue(sheetName, 56, 7, 1d),
					new ExpectedCellValue(sheetName, 46, 8, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 47, 8, "Chicago"),
					new ExpectedCellValue(sheetName, 48, 8, "Car Rack"),
					new ExpectedCellValue(sheetName, 49, 8, 415.75),
					new ExpectedCellValue(sheetName, 56, 8, 415.75),
					new ExpectedCellValue(sheetName, 47, 9, "Chicago Total"),
					new ExpectedCellValue(sheetName, 49, 9, 415.75),
					new ExpectedCellValue(sheetName, 56, 9, 415.75),
					new ExpectedCellValue(sheetName, 47, 10, "Nashville"),
					new ExpectedCellValue(sheetName, 48, 10, "Car Rack"),
					new ExpectedCellValue(sheetName, 55, 10, 415.75),
					new ExpectedCellValue(sheetName, 56, 10, 415.75),
					new ExpectedCellValue(sheetName, 47, 11, "Nashville Total"),
					new ExpectedCellValue(sheetName, 55, 11, 415.75),
					new ExpectedCellValue(sheetName, 56, 11, 415.75),
					new ExpectedCellValue(sheetName, 47, 12, "San Francisco"),
					new ExpectedCellValue(sheetName, 48, 12, "Car Rack"),
					new ExpectedCellValue(sheetName, 52, 12, 415.75),
					new ExpectedCellValue(sheetName, 56, 12, 415.75),
					new ExpectedCellValue(sheetName, 47, 13, "San Francisco Total"),
					new ExpectedCellValue(sheetName, 52, 13, 415.75),
					new ExpectedCellValue(sheetName, 56, 13, 415.75),
					new ExpectedCellValue(sheetName, 45, 14, "January Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 49, 14, 2d),
					new ExpectedCellValue(sheetName, 52, 14, 1d),
					new ExpectedCellValue(sheetName, 55, 14, 2d),
					new ExpectedCellValue(sheetName, 56, 14, 5d),
					new ExpectedCellValue(sheetName, 45, 15, "January Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 49, 15, 415.75),
					new ExpectedCellValue(sheetName, 52, 15, 415.75),
					new ExpectedCellValue(sheetName, 55, 15, 415.75),
					new ExpectedCellValue(sheetName, 56, 15, 1247.25),
					new ExpectedCellValue(sheetName, 45, 16, "February"),
					new ExpectedCellValue(sheetName, 46, 16, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 47, 16, "Nashville"),
					new ExpectedCellValue(sheetName, 48, 16, "Tent"),
					new ExpectedCellValue(sheetName, 51, 16, 6d),
					new ExpectedCellValue(sheetName, 56, 16, 6d),
					new ExpectedCellValue(sheetName, 47, 17, "Nashville Total"),
					new ExpectedCellValue(sheetName, 51, 17, 6d),
					new ExpectedCellValue(sheetName, 56, 17, 6d),
					new ExpectedCellValue(sheetName, 47, 18, "San Francisco"),
					new ExpectedCellValue(sheetName, 48, 18, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 54, 18, 1d),
					new ExpectedCellValue(sheetName, 56, 18, 1d),
					new ExpectedCellValue(sheetName, 47, 19, "San Francisco Total"),
					new ExpectedCellValue(sheetName, 54, 19, 1d),
					new ExpectedCellValue(sheetName, 56, 19, 1d),
					new ExpectedCellValue(sheetName, 46, 20, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 47, 20, "Nashville"),
					new ExpectedCellValue(sheetName, 48, 20, "Tent"),
					new ExpectedCellValue(sheetName, 51, 20, 199d),
					new ExpectedCellValue(sheetName, 56, 20, 199d),
					new ExpectedCellValue(sheetName, 47, 21, "Nashville Total"),
					new ExpectedCellValue(sheetName, 51, 21, 199d),
					new ExpectedCellValue(sheetName, 56, 21, 199d),
					new ExpectedCellValue(sheetName, 47, 22, "San Francisco"),
					new ExpectedCellValue(sheetName, 48, 22, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 54, 22, 99d),
					new ExpectedCellValue(sheetName, 56, 22, 99d),
					new ExpectedCellValue(sheetName, 47, 23, "San Francisco Total"),
					new ExpectedCellValue(sheetName, 54, 23, 99d),
					new ExpectedCellValue(sheetName, 56, 23, 99d),
					new ExpectedCellValue(sheetName, 45, 24, "February Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 51, 24, 6d),
					new ExpectedCellValue(sheetName, 54, 24, 1d),
					new ExpectedCellValue(sheetName, 56, 24, 7d),
					new ExpectedCellValue(sheetName, 45, 25, "February Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 51, 25, 199d),
					new ExpectedCellValue(sheetName, 54, 25, 99d),
					new ExpectedCellValue(sheetName, 56, 25, 298d),
					new ExpectedCellValue(sheetName, 45, 26, "March"),
					new ExpectedCellValue(sheetName, 46, 26, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 47, 26, "Chicago"),
					new ExpectedCellValue(sheetName, 48, 26, "Headlamp"),
					new ExpectedCellValue(sheetName, 53, 26, 1d),
					new ExpectedCellValue(sheetName, 56, 26, 1d),
					new ExpectedCellValue(sheetName, 47, 27, "Chicago Total"),
					new ExpectedCellValue(sheetName, 53, 27, 1d),
					new ExpectedCellValue(sheetName, 56, 27, 1d),
					new ExpectedCellValue(sheetName, 47, 28, "Nashville"),
					new ExpectedCellValue(sheetName, 48, 28, "Car Rack"),
					new ExpectedCellValue(sheetName, 50, 28, 2d),
					new ExpectedCellValue(sheetName, 56, 28, 2d),
					new ExpectedCellValue(sheetName, 47, 29, "Nashville Total"),
					new ExpectedCellValue(sheetName, 50, 29, 2d),
					new ExpectedCellValue(sheetName, 56, 29, 2d),
					new ExpectedCellValue(sheetName, 46, 30, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 47, 30, "Chicago"),
					new ExpectedCellValue(sheetName, 48, 30, "Headlamp"),
					new ExpectedCellValue(sheetName, 53, 30, 24.99),
					new ExpectedCellValue(sheetName, 56, 30, 24.99),
					new ExpectedCellValue(sheetName, 47, 31, "Chicago Total"),
					new ExpectedCellValue(sheetName, 53, 31, 24.99),
					new ExpectedCellValue(sheetName, 56, 31, 24.99),
					new ExpectedCellValue(sheetName, 47, 32, "Nashville"),
					new ExpectedCellValue(sheetName, 48, 32, "Car Rack"),
					new ExpectedCellValue(sheetName, 50, 32, 415.75),
					new ExpectedCellValue(sheetName, 56, 32, 415.75),
					new ExpectedCellValue(sheetName, 47, 33, "Nashville Total"),
					new ExpectedCellValue(sheetName, 50, 33, 415.75),
					new ExpectedCellValue(sheetName, 56, 33, 415.75),
					new ExpectedCellValue(sheetName, 45, 34, "March Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 50, 34, 2d),
					new ExpectedCellValue(sheetName, 53, 34, 1d),
					new ExpectedCellValue(sheetName, 56, 34, 3d),
					new ExpectedCellValue(sheetName, 45, 35, "March Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 50, 35, 415.75),
					new ExpectedCellValue(sheetName, 53, 35, 24.99),
					new ExpectedCellValue(sheetName, 56, 35, 440.74),
					new ExpectedCellValue(sheetName, 45, 36, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 49, 36, 2d),
					new ExpectedCellValue(sheetName, 50, 36, 2d),
					new ExpectedCellValue(sheetName, 51, 36, 6d),
					new ExpectedCellValue(sheetName, 52, 36, 1d),
					new ExpectedCellValue(sheetName, 53, 36, 1d),
					new ExpectedCellValue(sheetName, 54, 36, 1d),
					new ExpectedCellValue(sheetName, 55, 36, 2d),
					new ExpectedCellValue(sheetName, 56, 36, 15d),
					new ExpectedCellValue(sheetName, 45, 37, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 49, 37, 415.75),
					new ExpectedCellValue(sheetName, 50, 37, 415.75),
					new ExpectedCellValue(sheetName, 51, 37, 199d),
					new ExpectedCellValue(sheetName, 52, 37, 415.75),
					new ExpectedCellValue(sheetName, 53, 37, 24.99),
					new ExpectedCellValue(sheetName, 54, 37, 99d),
					new ExpectedCellValue(sheetName, 55, 37, 415.75),
					new ExpectedCellValue(sheetName, 56, 37, 1985.99)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithMultipleColumnDataFields.xlsx")]
		public void PivotTableRefreshMultipleColumnDataFieldsAsFirstInnerChildSubtotalsOff()
		{
			var file = new FileInfo("PivotTableWithMultipleColumnDataFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["ColumnDataFields"];
					var pivotTable = worksheet.PivotTables["PivotTable5"];
					foreach (var field in pivotTable.Fields)
					{
						field.DisableDefaultSubtotal();
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("A44:Q56"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(7, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(3, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(3, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "ColumnDataFields";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 49, 1, 20100007),
					new ExpectedCellValue(sheetName, 50, 1, 20100017),
					new ExpectedCellValue(sheetName, 51, 1, 20100070),
					new ExpectedCellValue(sheetName, 52, 1, 20100076),
					new ExpectedCellValue(sheetName, 53, 1, 20100083),
					new ExpectedCellValue(sheetName, 54, 1, 20100085),
					new ExpectedCellValue(sheetName, 55, 1, 20100090),
					new ExpectedCellValue(sheetName, 56, 1, "Grand Total"),
					new ExpectedCellValue(sheetName, 45, 2, "January"),
					new ExpectedCellValue(sheetName, 46, 2, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 47, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 48, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 49, 2, 2d),
					new ExpectedCellValue(sheetName, 56, 2, 2d),
					new ExpectedCellValue(sheetName, 47, 3, "Nashville"),
					new ExpectedCellValue(sheetName, 48, 3, "Car Rack"),
					new ExpectedCellValue(sheetName, 55, 3, 2d),
					new ExpectedCellValue(sheetName, 56, 3, 2d),
					new ExpectedCellValue(sheetName, 47, 4, "San Francisco"),
					new ExpectedCellValue(sheetName, 48, 4, "Car Rack"),
					new ExpectedCellValue(sheetName, 52, 4, 1d),
					new ExpectedCellValue(sheetName, 56, 4, 1d),
					new ExpectedCellValue(sheetName, 46, 5, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 47, 5, "Chicago"),
					new ExpectedCellValue(sheetName, 48, 5, "Car Rack"),
					new ExpectedCellValue(sheetName, 49, 5, 415.75),
					new ExpectedCellValue(sheetName, 56, 5, 415.75),
					new ExpectedCellValue(sheetName, 47, 6, "Nashville"),
					new ExpectedCellValue(sheetName, 48, 6, "Car Rack"),
					new ExpectedCellValue(sheetName, 55, 6, 415.75),
					new ExpectedCellValue(sheetName, 56, 6, 415.75),
					new ExpectedCellValue(sheetName, 47, 7, "San Francisco"),
					new ExpectedCellValue(sheetName, 48, 7, "Car Rack"),
					new ExpectedCellValue(sheetName, 52, 7, 415.75),
					new ExpectedCellValue(sheetName, 56, 7, 415.75),
					new ExpectedCellValue(sheetName, 45, 8, "February"),
					new ExpectedCellValue(sheetName, 46, 8, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 47, 8, "Nashville"),
					new ExpectedCellValue(sheetName, 48, 8, "Tent"),
					new ExpectedCellValue(sheetName, 51, 8, 6d),
					new ExpectedCellValue(sheetName, 56, 8, 6d),
					new ExpectedCellValue(sheetName, 47, 9, "San Francisco"),
					new ExpectedCellValue(sheetName, 48, 9, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 54, 9, 1d),
					new ExpectedCellValue(sheetName, 56, 9, 1d),
					new ExpectedCellValue(sheetName, 46, 10, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 47, 10, "Nashville"),
					new ExpectedCellValue(sheetName, 48, 10, "Tent"),
					new ExpectedCellValue(sheetName, 51, 10, 199d),
					new ExpectedCellValue(sheetName, 56, 10, 199d),
					new ExpectedCellValue(sheetName, 47, 11, "San Francisco"),
					new ExpectedCellValue(sheetName, 48, 11, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 54, 11, 99d),
					new ExpectedCellValue(sheetName, 56, 11, 99d),
					new ExpectedCellValue(sheetName, 45, 12, "March"),
					new ExpectedCellValue(sheetName, 46, 12, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 47, 12, "Chicago"),
					new ExpectedCellValue(sheetName, 48, 12, "Headlamp"),
					new ExpectedCellValue(sheetName, 53, 12, 1d),
					new ExpectedCellValue(sheetName, 56, 12, 1d),
					new ExpectedCellValue(sheetName, 47, 13, "Nashville"),
					new ExpectedCellValue(sheetName, 48, 13, "Car Rack"),
					new ExpectedCellValue(sheetName, 50, 13, 2d),
					new ExpectedCellValue(sheetName, 56, 13, 2d),
					new ExpectedCellValue(sheetName, 46, 14, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 47, 14, "Chicago"),
					new ExpectedCellValue(sheetName, 48, 14, "Headlamp"),
					new ExpectedCellValue(sheetName, 53, 14, 24.99),
					new ExpectedCellValue(sheetName, 56, 14, 24.99),
					new ExpectedCellValue(sheetName, 47, 15, "Nashville"),
					new ExpectedCellValue(sheetName, 48, 15, "Car Rack"),
					new ExpectedCellValue(sheetName, 50, 15, 415.75),
					new ExpectedCellValue(sheetName, 56, 15, 415.75),
					new ExpectedCellValue(sheetName, 45, 16, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 49, 16, 2d),
					new ExpectedCellValue(sheetName, 50, 16, 2d),
					new ExpectedCellValue(sheetName, 51, 16, 6d),
					new ExpectedCellValue(sheetName, 52, 16, 1d),
					new ExpectedCellValue(sheetName, 53, 16, 1d),
					new ExpectedCellValue(sheetName, 54, 16, 1d),
					new ExpectedCellValue(sheetName, 55, 16, 2d),
					new ExpectedCellValue(sheetName, 56, 16, 15d),
					new ExpectedCellValue(sheetName, 45, 17, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 49, 17, 415.75),
					new ExpectedCellValue(sheetName, 50, 17, 415.75),
					new ExpectedCellValue(sheetName, 51, 17, 199d),
					new ExpectedCellValue(sheetName, 52, 17, 415.75),
					new ExpectedCellValue(sheetName, 53, 17, 24.99),
					new ExpectedCellValue(sheetName, 54, 17, 99d),
					new ExpectedCellValue(sheetName, 55, 17, 415.75),
					new ExpectedCellValue(sheetName, 56, 17, 1985.99)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithMultipleColumnDataFields.xlsx")]
		public void PivotTableRefreshMultipleColumnDataFieldsAsLastInnerChildSubtotalsOn()
		{
			var file = new FileInfo("PivotTableWithMultipleColumnDataFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["ColumnDataFields"];
					var pivotTable = worksheet.PivotTables["PivotTable6"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("A60:AK72"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(8, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "ColumnDataFields";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 65, 1, 20100007),
					new ExpectedCellValue(sheetName, 66, 1, 20100017),
					new ExpectedCellValue(sheetName, 67, 1, 20100070),
					new ExpectedCellValue(sheetName, 68, 1, 20100076),
					new ExpectedCellValue(sheetName, 69, 1, 20100083),
					new ExpectedCellValue(sheetName, 70, 1, 20100085),
					new ExpectedCellValue(sheetName, 71, 1, 20100090),
					new ExpectedCellValue(sheetName, 72, 1, "Grand Total"),
					new ExpectedCellValue(sheetName, 61, 2, "January"),
					new ExpectedCellValue(sheetName, 62, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 63, 2, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 64, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 65, 2, 2d),
					new ExpectedCellValue(sheetName, 72, 2, 2d),
					new ExpectedCellValue(sheetName, 63, 3, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 64, 3, "Car Rack"),
					new ExpectedCellValue(sheetName, 65, 3, 415.75),
					new ExpectedCellValue(sheetName, 72, 3, 415.75),
					new ExpectedCellValue(sheetName, 62, 4, "Chicago Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 65, 4, 2d),
					new ExpectedCellValue(sheetName, 72, 4, 2d),
					new ExpectedCellValue(sheetName, 62, 5, "Chicago Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 65, 5, 415.75),
					new ExpectedCellValue(sheetName, 72, 5, 415.75),
					new ExpectedCellValue(sheetName, 62, 6, "Nashville"),
					new ExpectedCellValue(sheetName, 63, 6, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 64, 6, "Car Rack"),
					new ExpectedCellValue(sheetName, 71, 6, 2d),
					new ExpectedCellValue(sheetName, 72, 6, 2d),
					new ExpectedCellValue(sheetName, 63, 7, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 64, 7, "Car Rack"),
					new ExpectedCellValue(sheetName, 71, 7, 415.75),
					new ExpectedCellValue(sheetName, 72, 7, 415.75),
					new ExpectedCellValue(sheetName, 62, 8, "Nashville Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 71, 8, 2d),
					new ExpectedCellValue(sheetName, 72, 8, 2d),
					new ExpectedCellValue(sheetName, 62, 9, "Nashville Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 71, 9, 415.75),
					new ExpectedCellValue(sheetName, 72, 9, 415.75),
					new ExpectedCellValue(sheetName, 62, 10, "San Francisco"),
					new ExpectedCellValue(sheetName, 63, 10, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 64, 10, "Car Rack"),
					new ExpectedCellValue(sheetName, 68, 10, 1d),
					new ExpectedCellValue(sheetName, 72, 10, 1d),
					new ExpectedCellValue(sheetName, 63, 11, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 64, 11, "Car Rack"),
					new ExpectedCellValue(sheetName, 68, 11, 415.75),
					new ExpectedCellValue(sheetName, 72, 11, 415.75),
					new ExpectedCellValue(sheetName, 62, 12, "San Francisco Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 68, 12, 1d),
					new ExpectedCellValue(sheetName, 72, 12, 1d),
					new ExpectedCellValue(sheetName, 62, 13, "San Francisco Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 68, 13, 415.75),
					new ExpectedCellValue(sheetName, 72, 13, 415.75),
					new ExpectedCellValue(sheetName, 61, 14, "January Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 65, 14, 2d),
					new ExpectedCellValue(sheetName, 68, 14, 1d),
					new ExpectedCellValue(sheetName, 71, 14, 2d),
					new ExpectedCellValue(sheetName, 72, 14, 5d),
					new ExpectedCellValue(sheetName, 61, 15, "January Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 65, 15, 415.75),
					new ExpectedCellValue(sheetName, 68, 15, 415.75),
					new ExpectedCellValue(sheetName, 71, 15, 415.75),
					new ExpectedCellValue(sheetName, 72, 15, 1247.25),
					new ExpectedCellValue(sheetName, 61, 16, "February"),
					new ExpectedCellValue(sheetName, 62, 16, "Nashville"),
					new ExpectedCellValue(sheetName, 63, 16, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 64, 16, "Tent"),
					new ExpectedCellValue(sheetName, 67, 16, 6d),
					new ExpectedCellValue(sheetName, 72, 16, 6d),
					new ExpectedCellValue(sheetName, 63, 17, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 64, 17, "Tent"),
					new ExpectedCellValue(sheetName, 67, 17, 199d),
					new ExpectedCellValue(sheetName, 72, 17, 199d),
					new ExpectedCellValue(sheetName, 62, 18, "Nashville Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 67, 18, 6d),
					new ExpectedCellValue(sheetName, 72, 18, 6d),
					new ExpectedCellValue(sheetName, 62, 19, "Nashville Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 67, 19, 199d),
					new ExpectedCellValue(sheetName, 72, 19, 199d),
					new ExpectedCellValue(sheetName, 62, 20, "San Francisco"),
					new ExpectedCellValue(sheetName, 63, 20, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 64, 20, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 70, 20, 1d),
					new ExpectedCellValue(sheetName, 72, 20, 1d),
					new ExpectedCellValue(sheetName, 63, 21, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 64, 21, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 70, 21, 99d),
					new ExpectedCellValue(sheetName, 72, 21, 99d),
					new ExpectedCellValue(sheetName, 62, 22, "San Francisco Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 70, 22, 1d),
					new ExpectedCellValue(sheetName, 72, 22, 1d),
					new ExpectedCellValue(sheetName, 62, 23, "San Francisco Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 70, 23, 99d),
					new ExpectedCellValue(sheetName, 72, 23, 99d),
					new ExpectedCellValue(sheetName, 61, 24, "February Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 67, 24, 6d),
					new ExpectedCellValue(sheetName, 70, 24, 1d),
					new ExpectedCellValue(sheetName, 72, 24, 7d),
					new ExpectedCellValue(sheetName, 61, 25, "February Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 67, 25, 199d),
					new ExpectedCellValue(sheetName, 70, 25, 99d),
					new ExpectedCellValue(sheetName, 72, 25, 298d),
					new ExpectedCellValue(sheetName, 61, 26, "March"),
					new ExpectedCellValue(sheetName, 62, 26, "Chicago"),
					new ExpectedCellValue(sheetName, 63, 26, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 64, 26, "Headlamp"),
					new ExpectedCellValue(sheetName, 69, 26, 1d),
					new ExpectedCellValue(sheetName, 72, 26, 1d),
					new ExpectedCellValue(sheetName, 63, 27, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 64, 27, "Headlamp"),
					new ExpectedCellValue(sheetName, 69, 27, 24.99),
					new ExpectedCellValue(sheetName, 72, 27, 24.99),
					new ExpectedCellValue(sheetName, 62, 28, "Chicago Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 69, 28, 1d),
					new ExpectedCellValue(sheetName, 72, 28, 1d),
					new ExpectedCellValue(sheetName, 62, 29, "Chicago Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 69, 29, 24.99),
					new ExpectedCellValue(sheetName, 72, 29, 24.99),
					new ExpectedCellValue(sheetName, 62, 30, "Nashville"),
					new ExpectedCellValue(sheetName, 63, 30, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 64, 30, "Car Rack"),
					new ExpectedCellValue(sheetName, 66, 30, 2d),
					new ExpectedCellValue(sheetName, 72, 30, 2d),
					new ExpectedCellValue(sheetName, 63, 31, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 64, 31, "Car Rack"),
					new ExpectedCellValue(sheetName, 66, 31, 415.75),
					new ExpectedCellValue(sheetName, 72, 31, 415.75),
					new ExpectedCellValue(sheetName, 62, 32, "Nashville Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 66, 32, 2d),
					new ExpectedCellValue(sheetName, 72, 32, 2d),
					new ExpectedCellValue(sheetName, 62, 33, "Nashville Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 66, 33, 415.75),
					new ExpectedCellValue(sheetName, 72, 33, 415.75),
					new ExpectedCellValue(sheetName, 61, 34, "March Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 66, 34, 2d),
					new ExpectedCellValue(sheetName, 69, 34, 1d),
					new ExpectedCellValue(sheetName, 72, 34, 3d),
					new ExpectedCellValue(sheetName, 61, 35, "March Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 66, 35, 415.75),
					new ExpectedCellValue(sheetName, 69, 35, 24.99),
					new ExpectedCellValue(sheetName, 72, 35, 440.74),
					new ExpectedCellValue(sheetName, 61, 36, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 65, 36, 2d),
					new ExpectedCellValue(sheetName, 66, 36, 2d),
					new ExpectedCellValue(sheetName, 67, 36, 6d),
					new ExpectedCellValue(sheetName, 68, 36, 1d),
					new ExpectedCellValue(sheetName, 69, 36, 1d),
					new ExpectedCellValue(sheetName, 70, 36, 1d),
					new ExpectedCellValue(sheetName, 71, 36, 2d),
					new ExpectedCellValue(sheetName, 72, 36, 15d),
					new ExpectedCellValue(sheetName, 61, 37, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 65, 37, 415.75),
					new ExpectedCellValue(sheetName, 66, 37, 415.75),
					new ExpectedCellValue(sheetName, 67, 37, 199d),
					new ExpectedCellValue(sheetName, 68, 37, 415.75),
					new ExpectedCellValue(sheetName, 69, 37, 24.99),
					new ExpectedCellValue(sheetName, 70, 37, 99d),
					new ExpectedCellValue(sheetName, 71, 37, 415.75),
					new ExpectedCellValue(sheetName, 72, 37, 1985.99)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithMultipleColumnDataFields.xlsx")]
		public void PivotTableRefreshMultipleColumnDataFieldsAsLastInnerChildSubtotalsOff()
		{
			var file = new FileInfo("PivotTableWithMultipleColumnDataFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["ColumnDataFields"];
					var pivotTable = worksheet.PivotTables["PivotTable6"];
					foreach (var field in pivotTable.Fields)
					{
						field.DisableDefaultSubtotal();
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("A60:Q72"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(7, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(3, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(3, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "ColumnDataFields";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 65, 1, 20100007),
					new ExpectedCellValue(sheetName, 66, 1, 20100017),
					new ExpectedCellValue(sheetName, 67, 1, 20100070),
					new ExpectedCellValue(sheetName, 68, 1, 20100076),
					new ExpectedCellValue(sheetName, 69, 1, 20100083),
					new ExpectedCellValue(sheetName, 70, 1, 20100085),
					new ExpectedCellValue(sheetName, 71, 1, 20100090),
					new ExpectedCellValue(sheetName, 72, 1, "Grand Total"),
					new ExpectedCellValue(sheetName, 61, 2, "January"),
					new ExpectedCellValue(sheetName, 62, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 63, 2, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 64, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 65, 2, 2d),
					new ExpectedCellValue(sheetName, 72, 2, 2d),
					new ExpectedCellValue(sheetName, 63, 3, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 64, 3, "Car Rack"),
					new ExpectedCellValue(sheetName, 65, 3, 415.75),
					new ExpectedCellValue(sheetName, 72, 3, 415.75),
					new ExpectedCellValue(sheetName, 62, 4, "Nashville"),
					new ExpectedCellValue(sheetName, 63, 4, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 64, 4, "Car Rack"),
					new ExpectedCellValue(sheetName, 71, 4, 2d),
					new ExpectedCellValue(sheetName, 72, 4, 2d),
					new ExpectedCellValue(sheetName, 63, 5, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 64, 5, "Car Rack"),
					new ExpectedCellValue(sheetName, 71, 5, 415.75),
					new ExpectedCellValue(sheetName, 72, 5, 415.75),
					new ExpectedCellValue(sheetName, 62, 6, "San Francisco"),
					new ExpectedCellValue(sheetName, 63, 6, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 64, 6, "Car Rack"),
					new ExpectedCellValue(sheetName, 68, 6, 1d),
					new ExpectedCellValue(sheetName, 72, 6, 1d),
					new ExpectedCellValue(sheetName, 63, 7, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 64, 7, "Car Rack"),
					new ExpectedCellValue(sheetName, 68, 7, 415.75),
					new ExpectedCellValue(sheetName, 72, 7, 415.75),
					new ExpectedCellValue(sheetName, 61, 8, "February"),
					new ExpectedCellValue(sheetName, 62, 8, "Nashville"),
					new ExpectedCellValue(sheetName, 63, 8, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 64, 8, "Tent"),
					new ExpectedCellValue(sheetName, 67, 8, 6d),
					new ExpectedCellValue(sheetName, 72, 8, 6d),
					new ExpectedCellValue(sheetName, 63, 9, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 64, 9, "Tent"),
					new ExpectedCellValue(sheetName, 67, 9, 199d),
					new ExpectedCellValue(sheetName, 72, 9, 199d),
					new ExpectedCellValue(sheetName, 62, 10, "San Francisco"),
					new ExpectedCellValue(sheetName, 63, 10, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 64, 10, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 70, 10, 1d),
					new ExpectedCellValue(sheetName, 72, 10, 1d),
					new ExpectedCellValue(sheetName, 63, 11, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 64, 11, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 70, 11, 99d),
					new ExpectedCellValue(sheetName, 72, 11, 99d),
					new ExpectedCellValue(sheetName, 61, 12, "March"),
					new ExpectedCellValue(sheetName, 62, 12, "Chicago"),
					new ExpectedCellValue(sheetName, 63, 12, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 64, 12, "Headlamp"),
					new ExpectedCellValue(sheetName, 69, 12, 1d),
					new ExpectedCellValue(sheetName, 72, 12, 1d),
					new ExpectedCellValue(sheetName, 63, 13, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 64, 13, "Headlamp"),
					new ExpectedCellValue(sheetName, 69, 13, 24.99),
					new ExpectedCellValue(sheetName, 72, 13, 24.99),
					new ExpectedCellValue(sheetName, 62, 14, "Nashville"),
					new ExpectedCellValue(sheetName, 63, 14, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 64, 14, "Car Rack"),
					new ExpectedCellValue(sheetName, 66, 14, 2d),
					new ExpectedCellValue(sheetName, 72, 14, 2d),
					new ExpectedCellValue(sheetName, 63, 15, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 64, 15, "Car Rack"),
					new ExpectedCellValue(sheetName, 66, 15, 415.75),
					new ExpectedCellValue(sheetName, 72, 15, 415.75),
					new ExpectedCellValue(sheetName, 61, 16, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 65, 16, 2d),
					new ExpectedCellValue(sheetName, 66, 16, 2d),
					new ExpectedCellValue(sheetName, 67, 16, 6d),
					new ExpectedCellValue(sheetName, 68, 16, 1d),
					new ExpectedCellValue(sheetName, 69, 16, 1d),
					new ExpectedCellValue(sheetName, 70, 16, 1d),
					new ExpectedCellValue(sheetName, 71, 16, 2d),
					new ExpectedCellValue(sheetName, 72, 16, 15d),
					new ExpectedCellValue(sheetName, 61, 17, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 65, 17, 415.75),
					new ExpectedCellValue(sheetName, 66, 17, 415.75),
					new ExpectedCellValue(sheetName, 67, 17, 199d),
					new ExpectedCellValue(sheetName, 68, 17, 415.75),
					new ExpectedCellValue(sheetName, 69, 17, 24.99),
					new ExpectedCellValue(sheetName, 70, 17, 99d),
					new ExpectedCellValue(sheetName, 71, 17, 415.75),
					new ExpectedCellValue(sheetName, 72, 17, 1985.99)
				});
			}
		}
		#endregion
		#endregion

		#region Sorting Tests
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableSorting.xlsx")]
		public void PivotTableRefreshCustomSortingRowFieldsWithNoReferencesToDataFields()
		{
			var file = new FileInfo("PivotTableSorting.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["SortingWithoutReferences"];
					var pivotTable = worksheet.PivotTables["PivotTable1"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("A1:C12"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(8, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "SortingWithoutReferences";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 2, 1, "San Francisco"),
					new ExpectedCellValue(sheetName, 3, 1, 20100085),
					new ExpectedCellValue(sheetName, 4, 1, 20100076),
					new ExpectedCellValue(sheetName, 5, 1, "Nashville"),
					new ExpectedCellValue(sheetName, 6, 1, 20100090),
					new ExpectedCellValue(sheetName, 7, 1, 20100070),
					new ExpectedCellValue(sheetName, 8, 1, 20100017),
					new ExpectedCellValue(sheetName, 9, 1, "Chicago"),
					new ExpectedCellValue(sheetName, 10, 1, 20100083),
					new ExpectedCellValue(sheetName, 11, 1, 20100007),
					new ExpectedCellValue(sheetName, 12, 1, "Grand Total"),
					new ExpectedCellValue(sheetName, 1, 2, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 2, 2, 514.75),
					new ExpectedCellValue(sheetName, 3, 2, 99d),
					new ExpectedCellValue(sheetName, 4, 2, 415.75),
					new ExpectedCellValue(sheetName, 5, 2, 1030.5),
					new ExpectedCellValue(sheetName, 6, 2, 415.75),
					new ExpectedCellValue(sheetName, 7, 2, 199d),
					new ExpectedCellValue(sheetName, 8, 2, 415.75),
					new ExpectedCellValue(sheetName, 9, 2, 440.74),
					new ExpectedCellValue(sheetName, 10, 2, 24.99),
					new ExpectedCellValue(sheetName, 11, 2, 415.75),
					new ExpectedCellValue(sheetName, 12, 2, 1985.99),
					new ExpectedCellValue(sheetName, 1, 3, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 2, 3, 2d),
					new ExpectedCellValue(sheetName, 3, 3, 1d),
					new ExpectedCellValue(sheetName, 4, 3, 1d),
					new ExpectedCellValue(sheetName, 5, 3, 10d),
					new ExpectedCellValue(sheetName, 6, 3, 2d),
					new ExpectedCellValue(sheetName, 7, 3, 6d),
					new ExpectedCellValue(sheetName, 8, 3, 2d),
					new ExpectedCellValue(sheetName, 9, 3, 3d),
					new ExpectedCellValue(sheetName, 10, 3, 1d),
					new ExpectedCellValue(sheetName, 11, 3, 2d),
					new ExpectedCellValue(sheetName, 12, 3, 15d)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableSorting.xlsx")]
		public void PivotTableRefreshCustomSortingColumnFieldsWithNoReferencesToDataFields()
		{
			var file = new FileInfo("PivotTableSorting.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["SortingWithoutReferences"];
					var pivotTable = worksheet.PivotTables["PivotTable2"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("A16:W23"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(8, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "SortingWithoutReferences";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 20, 1, "January"),
					new ExpectedCellValue(sheetName, 21, 1, "February"),
					new ExpectedCellValue(sheetName, 22, 1, "March"),
					new ExpectedCellValue(sheetName, 23, 1, "Grand Total"),
					new ExpectedCellValue(sheetName, 17, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 18, 2, 20100085),
					new ExpectedCellValue(sheetName, 19, 2, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 21, 2, 99d),
					new ExpectedCellValue(sheetName, 23, 2, 99d),
					new ExpectedCellValue(sheetName, 19, 3, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 21, 3, 1d),
					new ExpectedCellValue(sheetName, 23, 3, 1d),
					new ExpectedCellValue(sheetName, 18, 4, 20100076),
					new ExpectedCellValue(sheetName, 19, 4, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 20, 4, 415.75),
					new ExpectedCellValue(sheetName, 23, 4, 415.75),
					new ExpectedCellValue(sheetName, 19, 5, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 20, 5, 1d),
					new ExpectedCellValue(sheetName, 23, 5, 1d),
					new ExpectedCellValue(sheetName, 17, 6, "San Francisco Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 20, 6, 415.75),
					new ExpectedCellValue(sheetName, 21, 6, 99d),
					new ExpectedCellValue(sheetName, 23, 6, 514.75),
					new ExpectedCellValue(sheetName, 17, 7, "San Francisco Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 20, 7, 1d),
					new ExpectedCellValue(sheetName, 21, 7, 1d),
					new ExpectedCellValue(sheetName, 23, 7, 2d),
					new ExpectedCellValue(sheetName, 17, 8, "Nashville"),
					new ExpectedCellValue(sheetName, 18, 8, 20100090),
					new ExpectedCellValue(sheetName, 19, 8, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 20, 8, 415.75),
					new ExpectedCellValue(sheetName, 23, 8, 415.75),
					new ExpectedCellValue(sheetName, 19, 9, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 20, 9, 2d),
					new ExpectedCellValue(sheetName, 23, 9, 2d),
					new ExpectedCellValue(sheetName, 18, 10, 20100070),
					new ExpectedCellValue(sheetName, 19, 10, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 21, 10, 199d),
					new ExpectedCellValue(sheetName, 23, 10, 199d),
					new ExpectedCellValue(sheetName, 19, 11, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 21, 11, 6d),
					new ExpectedCellValue(sheetName, 23, 11, 6d),
					new ExpectedCellValue(sheetName, 18, 12, 20100017),
					new ExpectedCellValue(sheetName, 19, 12, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 22, 12, 415.75),
					new ExpectedCellValue(sheetName, 23, 12, 415.75),
					new ExpectedCellValue(sheetName, 19, 13, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 22, 13, 2d),
					new ExpectedCellValue(sheetName, 23, 13, 2d),
					new ExpectedCellValue(sheetName, 17, 14, "Nashville Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 20, 14, 415.75),
					new ExpectedCellValue(sheetName, 21, 14, 199d),
					new ExpectedCellValue(sheetName, 22, 14, 415.75),
					new ExpectedCellValue(sheetName, 23, 14, 1030.5),
					new ExpectedCellValue(sheetName, 17, 15, "Nashville Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 20, 15, 2d),
					new ExpectedCellValue(sheetName, 21, 15, 6d),
					new ExpectedCellValue(sheetName, 22, 15, 2d),
					new ExpectedCellValue(sheetName, 23, 15, 10d),
					new ExpectedCellValue(sheetName, 17, 16, "Chicago"),
					new ExpectedCellValue(sheetName, 18, 16, 20100083),
					new ExpectedCellValue(sheetName, 19, 16, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 22, 16, 24.99),
					new ExpectedCellValue(sheetName, 23, 16, 24.99),
					new ExpectedCellValue(sheetName, 19, 17, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 22, 17, 1d),
					new ExpectedCellValue(sheetName, 23, 17, 1d),
					new ExpectedCellValue(sheetName, 18, 18, 20100007),
					new ExpectedCellValue(sheetName, 19, 18, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 20, 18, 415.75),
					new ExpectedCellValue(sheetName, 23, 18, 415.75),
					new ExpectedCellValue(sheetName, 19, 19, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 20, 19, 2d),
					new ExpectedCellValue(sheetName, 23, 19, 2d),
					new ExpectedCellValue(sheetName, 17, 20, "Chicago Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 20, 20, 415.75),
					new ExpectedCellValue(sheetName, 22, 20, 24.99),
					new ExpectedCellValue(sheetName, 23, 20, 440.74),
					new ExpectedCellValue(sheetName, 17, 21, "Chicago Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 20, 21, 2d),
					new ExpectedCellValue(sheetName, 22, 21, 1d),
					new ExpectedCellValue(sheetName, 23, 21, 3d),
					new ExpectedCellValue(sheetName, 17, 22, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 20, 22, 1247.25),
					new ExpectedCellValue(sheetName, 21, 22, 298d),
					new ExpectedCellValue(sheetName, 22, 22, 440.74),
					new ExpectedCellValue(sheetName, 23, 22, 1985.99),
					new ExpectedCellValue(sheetName, 17, 23, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 20, 23, 5d),
					new ExpectedCellValue(sheetName, 21, 23, 7d),
					new ExpectedCellValue(sheetName, 22, 23, 3d),
					new ExpectedCellValue(sheetName, 23, 23, 15d)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableSorting.xlsx")]
		public void PivotTableRefreshCustomSortingRowFieldsAndColumnFieldsWithNoReferencesToDataFields()
		{
			var file = new FileInfo("PivotTableSorting.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["SortingWithoutReferences"];
					var pivotTable = worksheet.PivotTables["PivotTable3"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("A27:I40"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(8, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "SortingWithoutReferences";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 30, 1, "San Francisco"),
					new ExpectedCellValue(sheetName, 31, 1, 20100085),
					new ExpectedCellValue(sheetName, 32, 1, 20100076),
					new ExpectedCellValue(sheetName, 33, 1, "Nashville"),
					new ExpectedCellValue(sheetName, 34, 1, 20100090),
					new ExpectedCellValue(sheetName, 35, 1, 20100070),
					new ExpectedCellValue(sheetName, 36, 1, 20100017),
					new ExpectedCellValue(sheetName, 37, 1, "Chicago"),
					new ExpectedCellValue(sheetName, 38, 1, 20100083),
					new ExpectedCellValue(sheetName, 39, 1, 20100007),
					new ExpectedCellValue(sheetName, 40, 1, "Grand Total"),
					new ExpectedCellValue(sheetName, 28, 2, "March"),
					new ExpectedCellValue(sheetName, 29, 2, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 33, 2, 2d),
					new ExpectedCellValue(sheetName, 36, 2, 2d),
					new ExpectedCellValue(sheetName, 37, 2, 1d),
					new ExpectedCellValue(sheetName, 38, 2, 1d),
					new ExpectedCellValue(sheetName, 40, 2, 3d),
					new ExpectedCellValue(sheetName, 29, 3, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 33, 3, 415.75),
					new ExpectedCellValue(sheetName, 36, 3, 415.75),
					new ExpectedCellValue(sheetName, 37, 3, 24.99),
					new ExpectedCellValue(sheetName, 38, 3, 24.99),
					new ExpectedCellValue(sheetName, 40, 3, 440.74),
					new ExpectedCellValue(sheetName, 28, 4, "February"),
					new ExpectedCellValue(sheetName, 29, 4, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 30, 4, 1d),
					new ExpectedCellValue(sheetName, 31, 4, 1d),
					new ExpectedCellValue(sheetName, 33, 4, 6d),
					new ExpectedCellValue(sheetName, 35, 4, 6d),
					new ExpectedCellValue(sheetName, 40, 4, 7d),
					new ExpectedCellValue(sheetName, 29, 5, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 30, 5, 99d),
					new ExpectedCellValue(sheetName, 31, 5, 99d),
					new ExpectedCellValue(sheetName, 33, 5, 199d),
					new ExpectedCellValue(sheetName, 35, 5, 199d),
					new ExpectedCellValue(sheetName, 40, 5, 298d),
					new ExpectedCellValue(sheetName, 28, 6, "January"),
					new ExpectedCellValue(sheetName, 29, 6, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 30, 6, 1d),
					new ExpectedCellValue(sheetName, 32, 6, 1d),
					new ExpectedCellValue(sheetName, 33, 6, 2d),
					new ExpectedCellValue(sheetName, 34, 6, 2d),
					new ExpectedCellValue(sheetName, 37, 6, 2d),
					new ExpectedCellValue(sheetName, 39, 6, 2d),
					new ExpectedCellValue(sheetName, 40, 6, 5d),
					new ExpectedCellValue(sheetName, 29, 7, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 30, 7, 415.75),
					new ExpectedCellValue(sheetName, 32, 7, 415.75),
					new ExpectedCellValue(sheetName, 33, 7, 415.75),
					new ExpectedCellValue(sheetName, 34, 7, 415.75),
					new ExpectedCellValue(sheetName, 37, 7, 415.75),
					new ExpectedCellValue(sheetName, 39, 7, 415.75),
					new ExpectedCellValue(sheetName, 40, 7, 1247.25),
					new ExpectedCellValue(sheetName, 28, 8, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 30, 8, 2d),
					new ExpectedCellValue(sheetName, 31, 8, 1d),
					new ExpectedCellValue(sheetName, 32, 8, 1d),
					new ExpectedCellValue(sheetName, 33, 8, 10d),
					new ExpectedCellValue(sheetName, 34, 8, 2d),
					new ExpectedCellValue(sheetName, 35, 8, 6d),
					new ExpectedCellValue(sheetName, 36, 8, 2d),
					new ExpectedCellValue(sheetName, 37, 8, 3d),
					new ExpectedCellValue(sheetName, 38, 8, 1d),
					new ExpectedCellValue(sheetName, 39, 8, 2d),
					new ExpectedCellValue(sheetName, 40, 8, 15d),
					new ExpectedCellValue(sheetName, 28, 9, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 30, 9, 514.75),
					new ExpectedCellValue(sheetName, 31, 9, 99d),
					new ExpectedCellValue(sheetName, 32, 9, 415.75),
					new ExpectedCellValue(sheetName, 33, 9, 1030.5),
					new ExpectedCellValue(sheetName, 34, 9, 415.75),
					new ExpectedCellValue(sheetName, 35, 9, 199d),
					new ExpectedCellValue(sheetName, 36, 9, 415.75),
					new ExpectedCellValue(sheetName, 37, 9, 440.74),
					new ExpectedCellValue(sheetName, 38, 9, 24.99),
					new ExpectedCellValue(sheetName, 39, 9, 415.75),
					new ExpectedCellValue(sheetName, 40, 9, 1985.99)

				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableSorting.xlsx")]
		public void PivotTableRefreshCustomSortingWithReferencesToDataFieldsOneRowFieldAndNoColumnFields()
		{
			var file = new FileInfo("PivotTableSorting.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["SortingWithReferencesRowFields"];
					var pivotTable = worksheet.PivotTables["PivotTable1"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("A1:C5"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "SortingWithReferencesRowFields";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 2, 1, "Nashville"),
					new ExpectedCellValue(sheetName, 3, 1, "San Francisco"),
					new ExpectedCellValue(sheetName, 4, 1, "Chicago"),
					new ExpectedCellValue(sheetName, 5, 1, "Grand Total"),
					new ExpectedCellValue(sheetName, 1, 2, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 2, 2, 1030.5),
					new ExpectedCellValue(sheetName, 3, 2, 514.75),
					new ExpectedCellValue(sheetName, 4, 2, 440.74),
					new ExpectedCellValue(sheetName, 5, 2, 1985.99),
					new ExpectedCellValue(sheetName, 1, 3, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 2, 3, 10d),
					new ExpectedCellValue(sheetName, 3, 3, 2d),
					new ExpectedCellValue(sheetName, 4, 3, 3d),
					new ExpectedCellValue(sheetName, 5, 3, 15d)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableSorting.xlsx")]
		public void PivotTableRefreshCustomSortingWithReferencesToDataFieldsTwoRowFieldsAndNoColumnFields()
		{
			var file = new FileInfo("PivotTableSorting.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["SortingWithReferencesRowFields"];
					var pivotTable = worksheet.PivotTables["PivotTable2"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("A9:C19"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "SortingWithReferencesRowFields";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 10, 1, "Nashville"),
					new ExpectedCellValue(sheetName, 11, 1, "Car Rack"),
					new ExpectedCellValue(sheetName, 12, 1, "Tent"),
					new ExpectedCellValue(sheetName, 13, 1, "San Francisco"),
					new ExpectedCellValue(sheetName, 14, 1, "Car Rack"),
					new ExpectedCellValue(sheetName, 15, 1, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 16, 1, "Chicago"),
					new ExpectedCellValue(sheetName, 17, 1, "Headlamp"),
					new ExpectedCellValue(sheetName, 18, 1, "Car Rack"),
					new ExpectedCellValue(sheetName, 19, 1, "Grand Total"),
					new ExpectedCellValue(sheetName, 9, 2, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 10, 2, 1030.5),
					new ExpectedCellValue(sheetName, 11, 2, 831.5),
					new ExpectedCellValue(sheetName, 12, 2, 199d),
					new ExpectedCellValue(sheetName, 13, 2, 514.75),
					new ExpectedCellValue(sheetName, 14, 2, 415.75),
					new ExpectedCellValue(sheetName, 15, 2, 99d),
					new ExpectedCellValue(sheetName, 16, 2, 440.74),
					new ExpectedCellValue(sheetName, 17, 2, 24.99),
					new ExpectedCellValue(sheetName, 18, 2, 415.75),
					new ExpectedCellValue(sheetName, 19, 2, 1985.99),
					new ExpectedCellValue(sheetName, 9, 3, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 10, 3, 10d),
					new ExpectedCellValue(sheetName, 11, 3, 4d),
					new ExpectedCellValue(sheetName, 12, 3, 6d),
					new ExpectedCellValue(sheetName, 13, 3, 2d),
					new ExpectedCellValue(sheetName, 14, 3, 1d),
					new ExpectedCellValue(sheetName, 15, 3, 1d),
					new ExpectedCellValue(sheetName, 16, 3, 3d),
					new ExpectedCellValue(sheetName, 17, 3, 1d),
					new ExpectedCellValue(sheetName, 18, 3, 2d),
					new ExpectedCellValue(sheetName, 19, 3, 15d)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableSorting.xlsx")]
		public void PivotTableRefreshCustomSortingWithReferencesToDataFieldsTwoRowFieldsAndNoColumnFieldsCorrectSortWithSameValuesDescending()
		{
			var file = new FileInfo("PivotTableSorting.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["SortingWithReferencesRowFields"];
					var pivotTable = worksheet.PivotTables["PivotTable3"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("A23:C34"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "SortingWithReferencesRowFields";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 24, 1, "Nashville"),
					new ExpectedCellValue(sheetName, 25, 1, "February"),
					new ExpectedCellValue(sheetName, 26, 1, "March"),
					new ExpectedCellValue(sheetName, 27, 1, "January"),
					new ExpectedCellValue(sheetName, 28, 1, "San Francisco"),
					new ExpectedCellValue(sheetName, 29, 1, "February"),
					new ExpectedCellValue(sheetName, 30, 1, "January"),
					new ExpectedCellValue(sheetName, 31, 1, "Chicago"),
					new ExpectedCellValue(sheetName, 32, 1, "January"),
					new ExpectedCellValue(sheetName, 33, 1, "March"),
					new ExpectedCellValue(sheetName, 34, 1, "Grand Total"),
					new ExpectedCellValue(sheetName, 23, 2, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 24, 2, 1030.5),
					new ExpectedCellValue(sheetName, 25, 2, 199d),
					new ExpectedCellValue(sheetName, 26, 2, 415.75),
					new ExpectedCellValue(sheetName, 27, 2, 415.75),
					new ExpectedCellValue(sheetName, 28, 2, 514.75),
					new ExpectedCellValue(sheetName, 29, 2, 99d),
					new ExpectedCellValue(sheetName, 30, 2, 415.75),
					new ExpectedCellValue(sheetName, 31, 2, 440.74),
					new ExpectedCellValue(sheetName, 32, 2, 415.75),
					new ExpectedCellValue(sheetName, 33, 2, 24.99),
					new ExpectedCellValue(sheetName, 34, 2, 1985.99),
					new ExpectedCellValue(sheetName, 23, 3, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 24, 3, 10d),
					new ExpectedCellValue(sheetName, 25, 3, 6d),
					new ExpectedCellValue(sheetName, 26, 3, 2d),
					new ExpectedCellValue(sheetName, 27, 3, 2d),
					new ExpectedCellValue(sheetName, 28, 3, 2d),
					new ExpectedCellValue(sheetName, 29, 3, 1d),
					new ExpectedCellValue(sheetName, 30, 3, 1d),
					new ExpectedCellValue(sheetName, 31, 3, 3d),
					new ExpectedCellValue(sheetName, 32, 3, 2d),
					new ExpectedCellValue(sheetName, 33, 3, 1d),
					new ExpectedCellValue(sheetName, 34, 3, 15d)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableSorting.xlsx")]
		public void PivotTableRefreshCustomSortingWithReferencesToDataFieldsTwoRowFieldsAndNoColumnFieldsCorrectSortWithSameValuesAscending()
		{
			var file = new FileInfo("PivotTableSorting.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["SortingWithReferencesRowFields"];
					var pivotTable = worksheet.PivotTables["PivotTable4"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("A38:C49"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "SortingWithReferencesRowFields";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 39, 1, "Nashville"),
					new ExpectedCellValue(sheetName, 40, 1, "February"),
					new ExpectedCellValue(sheetName, 41, 1, "January"),
					new ExpectedCellValue(sheetName, 42, 1, "March"),
					new ExpectedCellValue(sheetName, 43, 1, "San Francisco"),
					new ExpectedCellValue(sheetName, 44, 1, "February"),
					new ExpectedCellValue(sheetName, 45, 1, "January"),
					new ExpectedCellValue(sheetName, 46, 1, "Chicago"),
					new ExpectedCellValue(sheetName, 47, 1, "March"),
					new ExpectedCellValue(sheetName, 48, 1, "January"),
					new ExpectedCellValue(sheetName, 49, 1, "Grand Total"),
					new ExpectedCellValue(sheetName, 38, 2, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 39, 2, 1030.5),
					new ExpectedCellValue(sheetName, 40, 2, 199d),
					new ExpectedCellValue(sheetName, 41, 2, 415.75),
					new ExpectedCellValue(sheetName, 42, 2, 415.75),
					new ExpectedCellValue(sheetName, 43, 2, 514.75),
					new ExpectedCellValue(sheetName, 44, 2, 99d),
					new ExpectedCellValue(sheetName, 45, 2, 415.75),
					new ExpectedCellValue(sheetName, 46, 2, 440.74),
					new ExpectedCellValue(sheetName, 47, 2, 24.99),
					new ExpectedCellValue(sheetName, 48, 2, 415.75),
					new ExpectedCellValue(sheetName, 49, 2, 1985.99),
					new ExpectedCellValue(sheetName, 38, 3, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 39, 3, 10d),
					new ExpectedCellValue(sheetName, 40, 3, 6d),
					new ExpectedCellValue(sheetName, 41, 3, 2d),
					new ExpectedCellValue(sheetName, 42, 3, 2d),
					new ExpectedCellValue(sheetName, 43, 3, 2d),
					new ExpectedCellValue(sheetName, 44, 3, 1d),
					new ExpectedCellValue(sheetName, 45, 3, 1d),
					new ExpectedCellValue(sheetName, 46, 3, 3d),
					new ExpectedCellValue(sheetName, 47, 3, 1d),
					new ExpectedCellValue(sheetName, 48, 3, 2d),
					new ExpectedCellValue(sheetName, 49, 3, 15d)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableSorting.xlsx")]
		public void PivotTableRefreshCustomSortingWithReferencesToDataFieldsTwoRowFieldsAndOneColumnFields()
		{
			var file = new FileInfo("PivotTableSorting.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["RowColumnFieldsSorting"];
					var pivotTable = worksheet.PivotTables["PivotTable1"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("A1:K14"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "RowColumnFieldsSorting";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 4, 1, "Nashville"),
					new ExpectedCellValue(sheetName, 5, 1, "February"),
					new ExpectedCellValue(sheetName, 6, 1, "January"),
					new ExpectedCellValue(sheetName, 7, 1, "March"),
					new ExpectedCellValue(sheetName, 8, 1, "San Francisco"),
					new ExpectedCellValue(sheetName, 9, 1, "February"),
					new ExpectedCellValue(sheetName, 10, 1, "January"),
					new ExpectedCellValue(sheetName, 11, 1, "Chicago"),
					new ExpectedCellValue(sheetName, 12, 1, "March"),
					new ExpectedCellValue(sheetName, 13, 1, "January"),
					new ExpectedCellValue(sheetName, 14, 1, "Grand Total"),
					new ExpectedCellValue(sheetName, 2, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 3, 2, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 4, 2, 831.5),
					new ExpectedCellValue(sheetName, 6, 2, 415.75),
					new ExpectedCellValue(sheetName, 7, 2, 415.75),
					new ExpectedCellValue(sheetName, 8, 2, 415.75),
					new ExpectedCellValue(sheetName, 10, 2, 415.75),
					new ExpectedCellValue(sheetName, 11, 2, 415.75),
					new ExpectedCellValue(sheetName, 13, 2, 415.75),
					new ExpectedCellValue(sheetName, 14, 2, 1663d),
					new ExpectedCellValue(sheetName, 3, 3, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 4, 3, 4d),
					new ExpectedCellValue(sheetName, 6, 3, 2d),
					new ExpectedCellValue(sheetName, 7, 3, 2d),
					new ExpectedCellValue(sheetName, 8, 3, 1d),
					new ExpectedCellValue(sheetName, 10, 3, 1d),
					new ExpectedCellValue(sheetName, 11, 3, 2d),
					new ExpectedCellValue(sheetName, 13, 3, 2d),
					new ExpectedCellValue(sheetName, 14, 3, 7d),
					new ExpectedCellValue(sheetName, 2, 4, "Headlamp"),
					new ExpectedCellValue(sheetName, 3, 4, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 11, 4, 24.99),
					new ExpectedCellValue(sheetName, 12, 4, 24.99),
					new ExpectedCellValue(sheetName, 14, 4, 24.99),
					new ExpectedCellValue(sheetName, 3, 5, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 11, 5, 1d),
					new ExpectedCellValue(sheetName, 12, 5, 1d),
					new ExpectedCellValue(sheetName, 14, 5, 1d),
					new ExpectedCellValue(sheetName, 2, 6, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 3, 6, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 8, 6, 99d),
					new ExpectedCellValue(sheetName, 9, 6, 99d),
					new ExpectedCellValue(sheetName, 14, 6, 99d),
					new ExpectedCellValue(sheetName, 3, 7, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 8, 7, 1d),
					new ExpectedCellValue(sheetName, 9, 7, 1d),
					new ExpectedCellValue(sheetName, 14, 7, 1d),
					new ExpectedCellValue(sheetName, 2, 8, "Tent"),
					new ExpectedCellValue(sheetName, 3, 8, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 4, 8, 199d),
					new ExpectedCellValue(sheetName, 5, 8, 199d),
					new ExpectedCellValue(sheetName, 14, 8, 199d),
					new ExpectedCellValue(sheetName, 3, 9, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 4, 9, 6d),
					new ExpectedCellValue(sheetName, 5, 9, 6d),
					new ExpectedCellValue(sheetName, 14, 9, 6d),
					new ExpectedCellValue(sheetName, 2, 10, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 4, 10, 1030.5),
					new ExpectedCellValue(sheetName, 5, 10, 199d),
					new ExpectedCellValue(sheetName, 6, 10, 415.75),
					new ExpectedCellValue(sheetName, 7, 10, 415.75),
					new ExpectedCellValue(sheetName, 8, 10, 514.75),
					new ExpectedCellValue(sheetName, 9, 10, 99d),
					new ExpectedCellValue(sheetName, 10, 10, 415.75),
					new ExpectedCellValue(sheetName, 11, 10, 440.74),
					new ExpectedCellValue(sheetName, 12, 10, 24.99),
					new ExpectedCellValue(sheetName, 13, 10, 415.75),
					new ExpectedCellValue(sheetName, 14, 10, 1985.99),
					new ExpectedCellValue(sheetName, 2, 11, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 4, 11, 10d),
					new ExpectedCellValue(sheetName, 5, 11, 6d),
					new ExpectedCellValue(sheetName, 6, 11, 2d),
					new ExpectedCellValue(sheetName, 7, 11, 2d),
					new ExpectedCellValue(sheetName, 8, 11, 2d),
					new ExpectedCellValue(sheetName, 9, 11, 1d),
					new ExpectedCellValue(sheetName, 10, 11, 1d),
					new ExpectedCellValue(sheetName, 11, 11, 3d),
					new ExpectedCellValue(sheetName, 12, 11, 1d),
					new ExpectedCellValue(sheetName, 13, 11, 2d),
					new ExpectedCellValue(sheetName, 14, 11, 15d)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableSorting.xlsx")]
		public void PivotTableRefreshCustomSortingWithReferencesToDataFieldsOneRowFieldAndTwoColumnFields()
		{
			var file = new FileInfo("PivotTableSorting.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["RowColumnFieldsSorting"];
					var pivotTable = worksheet.PivotTables["PivotTable2"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("A18:S25"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "RowColumnFieldsSorting";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 22, 1, "Nashville"),
					new ExpectedCellValue(sheetName, 23, 1, "San Francisco"),
					new ExpectedCellValue(sheetName, 24, 1, "Chicago"),
					new ExpectedCellValue(sheetName, 25, 1, "Grand Total"),
					new ExpectedCellValue(sheetName, 19, 2, "March"),
					new ExpectedCellValue(sheetName, 20, 2, "Headlamp"),
					new ExpectedCellValue(sheetName, 21, 2, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 24, 2, 24.99),
					new ExpectedCellValue(sheetName, 25, 2, 24.99),
					new ExpectedCellValue(sheetName, 21, 3, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 24, 3, 1d),
					new ExpectedCellValue(sheetName, 25, 3, 1d),
					new ExpectedCellValue(sheetName, 20, 4, "Car Rack"),
					new ExpectedCellValue(sheetName, 21, 4, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 22, 4, 415.75),
					new ExpectedCellValue(sheetName, 25, 4, 415.75),
					new ExpectedCellValue(sheetName, 21, 5, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 22, 5, 2d),
					new ExpectedCellValue(sheetName, 25, 5, 2d),
					new ExpectedCellValue(sheetName, 19, 6, "March Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 22, 6, 415.75),
					new ExpectedCellValue(sheetName, 24, 6, 24.99),
					new ExpectedCellValue(sheetName, 25, 6, 440.74),
					new ExpectedCellValue(sheetName, 19, 7, "March Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 22, 7, 2d),
					new ExpectedCellValue(sheetName, 24, 7, 1d),
					new ExpectedCellValue(sheetName, 25, 7, 3d),
					new ExpectedCellValue(sheetName, 19, 8, "January"),
					new ExpectedCellValue(sheetName, 20, 8, "Car Rack"),
					new ExpectedCellValue(sheetName, 21, 8, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 22, 8, 415.75),
					new ExpectedCellValue(sheetName, 23, 8, 415.75),
					new ExpectedCellValue(sheetName, 24, 8, 415.75),
					new ExpectedCellValue(sheetName, 25, 8, 1247.25),
					new ExpectedCellValue(sheetName, 21, 9, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 22, 9, 2d),
					new ExpectedCellValue(sheetName, 23, 9, 1d),
					new ExpectedCellValue(sheetName, 24, 9, 2d),
					new ExpectedCellValue(sheetName, 25, 9, 5d),
					new ExpectedCellValue(sheetName, 19, 10, "January Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 22, 10, 415.75),
					new ExpectedCellValue(sheetName, 23, 10, 415.75),
					new ExpectedCellValue(sheetName, 24, 10, 415.75),
					new ExpectedCellValue(sheetName, 25, 10, 1247.25),
					new ExpectedCellValue(sheetName, 19, 11, "January Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 22, 11, 2d),
					new ExpectedCellValue(sheetName, 23, 11, 1d),
					new ExpectedCellValue(sheetName, 24, 11, 2d),
					new ExpectedCellValue(sheetName, 25, 11, 5d),
					new ExpectedCellValue(sheetName, 19, 12, "February"),
					new ExpectedCellValue(sheetName, 20, 12, "Tent"),
					new ExpectedCellValue(sheetName, 21, 12, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 22, 12, 199d),
					new ExpectedCellValue(sheetName, 25, 12, 199d),
					new ExpectedCellValue(sheetName, 21, 13, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 22, 13, 6d),
					new ExpectedCellValue(sheetName, 25, 13, 6d),
					new ExpectedCellValue(sheetName, 20, 14, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 21, 14, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 23, 14, 99d),
					new ExpectedCellValue(sheetName, 25, 14, 99d),
					new ExpectedCellValue(sheetName, 21, 15, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 23, 15, 1d),
					new ExpectedCellValue(sheetName, 25, 15, 1d),
					new ExpectedCellValue(sheetName, 19, 16, "February Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 22, 16, 199d),
					new ExpectedCellValue(sheetName, 23, 16, 99d),
					new ExpectedCellValue(sheetName, 25, 16, 298d),
					new ExpectedCellValue(sheetName, 19, 17, "February Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 22, 17, 6d),
					new ExpectedCellValue(sheetName, 23, 17, 1d),
					new ExpectedCellValue(sheetName, 25, 17, 7d),
					new ExpectedCellValue(sheetName, 19, 18, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 22, 18, 1030.5),
					new ExpectedCellValue(sheetName, 23, 18, 514.75),
					new ExpectedCellValue(sheetName, 24, 18, 440.74),
					new ExpectedCellValue(sheetName, 25, 18, 1985.99),
					new ExpectedCellValue(sheetName, 19, 19, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 22, 19, 10d),
					new ExpectedCellValue(sheetName, 23, 19, 2d),
					new ExpectedCellValue(sheetName, 24, 19, 3d),
					new ExpectedCellValue(sheetName, 25, 19, 15d),
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableSorting.xlsx")]
		public void PivotTableRefreshCustomSortingWithReferencesToDataFieldsTwoRowFieldAndTwoColumnFields()
		{
			var file = new FileInfo("PivotTableSorting.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["RowColumnFieldsSorting"];
					var pivotTable = worksheet.PivotTables["PivotTable3"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("A29:AH44"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(8, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "RowColumnFieldsSorting";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					// These values will need to be changed when Bug #10033 is completed.
					new ExpectedCellValue(sheetName, 33, 1, "Car Rack"),
					new ExpectedCellValue(sheetName, 34, 1, 20100076),
					new ExpectedCellValue(sheetName, 35, 1, 20100007),
					new ExpectedCellValue(sheetName, 36, 1, 20100017),
					new ExpectedCellValue(sheetName, 37, 1, 20100090),
					new ExpectedCellValue(sheetName, 38, 1, "Tent"),
					new ExpectedCellValue(sheetName, 39, 1, 20100070),
					new ExpectedCellValue(sheetName, 40, 1, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 41, 1, 20100085),
					new ExpectedCellValue(sheetName, 42, 1, "Headlamp"),
					new ExpectedCellValue(sheetName, 43, 1, 20100083),
					new ExpectedCellValue(sheetName, 44, 1, "Grand Total"),
					new ExpectedCellValue(sheetName, 30, 2, "March"),
					new ExpectedCellValue(sheetName, 31, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 32, 2, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 42, 2, 24.99),
					new ExpectedCellValue(sheetName, 43, 2, 24.99),
					new ExpectedCellValue(sheetName, 44, 2, 24.99),
					new ExpectedCellValue(sheetName, 32, 3, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 42, 3, 1d),
					new ExpectedCellValue(sheetName, 43, 3, 1d),
					new ExpectedCellValue(sheetName, 44, 3, 1d),
					new ExpectedCellValue(sheetName, 32, 4, "Sum of Total"),
					new ExpectedCellValue(sheetName, 42, 4, 24.99),
					new ExpectedCellValue(sheetName, 43, 4, 24.99),
					new ExpectedCellValue(sheetName, 44, 4, 24.99),
					new ExpectedCellValue(sheetName, 31, 5, "Nashville"),
					new ExpectedCellValue(sheetName, 32, 5, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 33, 5, 415.75),
					new ExpectedCellValue(sheetName, 36, 5, 415.75),
					new ExpectedCellValue(sheetName, 44, 5, 415.75),
					new ExpectedCellValue(sheetName, 32, 6, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 33, 6, 2d),
					new ExpectedCellValue(sheetName, 36, 6, 2d),
					new ExpectedCellValue(sheetName, 44, 6, 2d),
					new ExpectedCellValue(sheetName, 32, 7, "Sum of Total"),
					new ExpectedCellValue(sheetName, 33, 7, 831.5),
					new ExpectedCellValue(sheetName, 36, 7, 831.5),
					new ExpectedCellValue(sheetName, 44, 7, 831.5),
					new ExpectedCellValue(sheetName, 30, 8, "March Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 33, 8, 415.75),
					new ExpectedCellValue(sheetName, 36, 8, 415.75),
					new ExpectedCellValue(sheetName, 42, 8, 24.99),
					new ExpectedCellValue(sheetName, 43, 8, 24.99),
					new ExpectedCellValue(sheetName, 44, 8, 440.74),
					new ExpectedCellValue(sheetName, 30, 9, "March Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 33, 9, 2d),
					new ExpectedCellValue(sheetName, 36, 9, 2d),
					new ExpectedCellValue(sheetName, 42, 9, 1d),
					new ExpectedCellValue(sheetName, 43, 9, 1d),
					new ExpectedCellValue(sheetName, 44, 9, 3d),
					new ExpectedCellValue(sheetName, 30, 10, "March Sum of Total"),
					new ExpectedCellValue(sheetName, 33, 10, 831.5),
					new ExpectedCellValue(sheetName, 36, 10, 831.5),
					new ExpectedCellValue(sheetName, 42, 10, 24.99),
					new ExpectedCellValue(sheetName, 43, 10, 24.99),
					new ExpectedCellValue(sheetName, 44, 10, 856.49),
					new ExpectedCellValue(sheetName, 30, 11, "February"),
					new ExpectedCellValue(sheetName, 31, 11, "San Francisco"),
					new ExpectedCellValue(sheetName, 32, 11, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 40, 11, 99d),
					new ExpectedCellValue(sheetName, 41, 11, 99d),
					new ExpectedCellValue(sheetName, 44, 11, 99d),
					new ExpectedCellValue(sheetName, 32, 12, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 40, 12, 1d),
					new ExpectedCellValue(sheetName, 41, 12, 1d),
					new ExpectedCellValue(sheetName, 44, 12, 1d),
					new ExpectedCellValue(sheetName, 32, 13, "Sum of Total"),
					new ExpectedCellValue(sheetName, 40, 13, 99d),
					new ExpectedCellValue(sheetName, 41, 13, 99d),
					new ExpectedCellValue(sheetName, 44, 13, 99d),
					new ExpectedCellValue(sheetName, 31, 14, "Nashville"),
					new ExpectedCellValue(sheetName, 32, 14, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 38, 14, 199d),
					new ExpectedCellValue(sheetName, 39, 14, 199d),
					new ExpectedCellValue(sheetName, 44, 14, 199d),
					new ExpectedCellValue(sheetName, 32, 15, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 38, 15, 6d),
					new ExpectedCellValue(sheetName, 39, 15, 6d),
					new ExpectedCellValue(sheetName, 44, 15, 6d),
					new ExpectedCellValue(sheetName, 32, 16, "Sum of Total"),
					new ExpectedCellValue(sheetName, 38, 16, 1194d),
					new ExpectedCellValue(sheetName, 39, 16, 1194d),
					new ExpectedCellValue(sheetName, 44, 16, 1194d),
					new ExpectedCellValue(sheetName, 30, 17, "February Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 38, 17, 199d),
					new ExpectedCellValue(sheetName, 39, 17, 199d),
					new ExpectedCellValue(sheetName, 40, 17, 99d),
					new ExpectedCellValue(sheetName, 41, 17, 99d),
					new ExpectedCellValue(sheetName, 44, 17, 298d),
					new ExpectedCellValue(sheetName, 30, 18, "February Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 38, 18, 6d),
					new ExpectedCellValue(sheetName, 39, 18, 6d),
					new ExpectedCellValue(sheetName, 40, 18, 1d),
					new ExpectedCellValue(sheetName, 41, 18, 1d),
					new ExpectedCellValue(sheetName, 44, 18, 7d),
					new ExpectedCellValue(sheetName, 30, 19, "February Sum of Total"),
					new ExpectedCellValue(sheetName, 38, 19, 1194d),
					new ExpectedCellValue(sheetName, 39, 19, 1194d),
					new ExpectedCellValue(sheetName, 40, 19, 99d),
					new ExpectedCellValue(sheetName, 41, 19, 99d),
					new ExpectedCellValue(sheetName, 44, 19, 1293d),
					new ExpectedCellValue(sheetName, 30, 20, "January"),
					new ExpectedCellValue(sheetName, 31, 20, "San Francisco"),
					new ExpectedCellValue(sheetName, 32, 20, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 33, 20, 415.75),
					new ExpectedCellValue(sheetName, 34, 20, 415.75),
					new ExpectedCellValue(sheetName, 44, 20, 415.75),
					new ExpectedCellValue(sheetName, 32, 21, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 33, 21, 1d),
					new ExpectedCellValue(sheetName, 34, 21, 1d),
					new ExpectedCellValue(sheetName, 44, 21, 1d),
					new ExpectedCellValue(sheetName, 32, 22, "Sum of Total"),
					new ExpectedCellValue(sheetName, 33, 22, 415.75),
					new ExpectedCellValue(sheetName, 34, 22, 415.75),
					new ExpectedCellValue(sheetName, 44, 22, 415.75),
					new ExpectedCellValue(sheetName, 31, 23, "Chicago"),
					new ExpectedCellValue(sheetName, 32, 23, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 33, 23, 415.75),
					new ExpectedCellValue(sheetName, 35, 23, 415.75),
					new ExpectedCellValue(sheetName, 44, 23, 415.75),
					new ExpectedCellValue(sheetName, 32, 24, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 33, 24, 2d),
					new ExpectedCellValue(sheetName, 35, 24, 2d),
					new ExpectedCellValue(sheetName, 44, 24, 2d),
					new ExpectedCellValue(sheetName, 32, 25, "Sum of Total"),
					new ExpectedCellValue(sheetName, 33, 25, 831.5),
					new ExpectedCellValue(sheetName, 35, 25, 831.5),
					new ExpectedCellValue(sheetName, 44, 25, 831.5),
					new ExpectedCellValue(sheetName, 31, 26, "Nashville"),
					new ExpectedCellValue(sheetName, 32, 26, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 33, 26, 415.75),
					new ExpectedCellValue(sheetName, 37, 26, 415.75),
					new ExpectedCellValue(sheetName, 44, 26, 415.75),
					new ExpectedCellValue(sheetName, 32, 27, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 33, 27, 2d),
					new ExpectedCellValue(sheetName, 37, 27, 2d),
					new ExpectedCellValue(sheetName, 44, 27, 2d),
					new ExpectedCellValue(sheetName, 32, 28, "Sum of Total"),
					new ExpectedCellValue(sheetName, 33, 28, 831.5),
					new ExpectedCellValue(sheetName, 37, 28, 831.5),
					new ExpectedCellValue(sheetName, 44, 28, 831.5),
					new ExpectedCellValue(sheetName, 30, 29, "January Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 33, 29, 1247.25),
					new ExpectedCellValue(sheetName, 34, 29, 415.75),
					new ExpectedCellValue(sheetName, 35, 29, 415.75),
					new ExpectedCellValue(sheetName, 37, 29, 415.75),
					new ExpectedCellValue(sheetName, 44, 29, 1247.25),
					new ExpectedCellValue(sheetName, 30, 30, "January Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 33, 30, 5d),
					new ExpectedCellValue(sheetName, 34, 30, 1d),
					new ExpectedCellValue(sheetName, 35, 30, 2d),
					new ExpectedCellValue(sheetName, 37, 30, 2d),
					new ExpectedCellValue(sheetName, 44, 30, 5d),
					new ExpectedCellValue(sheetName, 30, 31, "January Sum of Total"),
					new ExpectedCellValue(sheetName, 33, 31, 2078.75),
					new ExpectedCellValue(sheetName, 34, 31, 415.75),
					new ExpectedCellValue(sheetName, 35, 31, 831.5),
					new ExpectedCellValue(sheetName, 37, 31, 831.5),
					new ExpectedCellValue(sheetName, 44, 31, 2078.75),
					new ExpectedCellValue(sheetName, 30, 32, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 33, 32, 1663d),
					new ExpectedCellValue(sheetName, 34, 32, 415.75),
					new ExpectedCellValue(sheetName, 35, 32, 415.75),
					new ExpectedCellValue(sheetName, 36, 32, 415.75),
					new ExpectedCellValue(sheetName, 37, 32, 415.75),
					new ExpectedCellValue(sheetName, 38, 32, 199d),
					new ExpectedCellValue(sheetName, 39, 32, 199d),
					new ExpectedCellValue(sheetName, 40, 32, 99d),
					new ExpectedCellValue(sheetName, 41, 32, 99d),
					new ExpectedCellValue(sheetName, 42, 32, 24.99),
					new ExpectedCellValue(sheetName, 43, 32, 24.99),
					new ExpectedCellValue(sheetName, 44, 32, 1985.99),
					new ExpectedCellValue(sheetName, 30, 33, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 33, 33, 7d),
					new ExpectedCellValue(sheetName, 34, 33, 1d),
					new ExpectedCellValue(sheetName, 35, 33, 2d),
					new ExpectedCellValue(sheetName, 36, 33, 2d),
					new ExpectedCellValue(sheetName, 37, 33, 2d),
					new ExpectedCellValue(sheetName, 38, 33, 6d),
					new ExpectedCellValue(sheetName, 39, 33, 6d),
					new ExpectedCellValue(sheetName, 40, 33, 1d),
					new ExpectedCellValue(sheetName, 41, 33, 1d),
					new ExpectedCellValue(sheetName, 42, 33, 1d),
					new ExpectedCellValue(sheetName, 43, 33, 1d),
					new ExpectedCellValue(sheetName, 44, 33, 15d),
					new ExpectedCellValue(sheetName, 30, 34, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 33, 34, 2910.25),
					new ExpectedCellValue(sheetName, 34, 34, 415.75),
					new ExpectedCellValue(sheetName, 35, 34, 831.5),
					new ExpectedCellValue(sheetName, 36, 34, 831.5),
					new ExpectedCellValue(sheetName, 37, 34, 831.5),
					new ExpectedCellValue(sheetName, 38, 34, 1194d),
					new ExpectedCellValue(sheetName, 39, 34, 1194d),
					new ExpectedCellValue(sheetName, 40, 34, 99d),
					new ExpectedCellValue(sheetName, 41, 34, 99d),
					new ExpectedCellValue(sheetName, 42, 34, 24.99),
					new ExpectedCellValue(sheetName, 43, 34, 24.99),
					new ExpectedCellValue(sheetName, 44, 34, 4228.24)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableSorting.xlsx")]
		public void PivotTableRefreshCustomSortingWithReferencesToDataFieldsRowFieldsWithDataFieldsSortingRowsAndColumns()
		{
			var file = new FileInfo("PivotTableSorting.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["DataFieldsInRows"];
					var pivotTable = worksheet.PivotTables["PivotTable1"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("A1:E13"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "DataFieldsInRows";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 3, 1, "San Francisco"),
					new ExpectedCellValue(sheetName, 4, 1, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 5, 1, "Sum of Total"),
					new ExpectedCellValue(sheetName, 6, 1, "Chicago"),
					new ExpectedCellValue(sheetName, 7, 1, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 8, 1, "Sum of Total"),
					new ExpectedCellValue(sheetName, 9, 1, "Nashville"),
					new ExpectedCellValue(sheetName, 10, 1, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 11, 1, "Sum of Total"),
					new ExpectedCellValue(sheetName, 12, 1, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 13, 1, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 2, 2, "February"),
					new ExpectedCellValue(sheetName, 4, 2, 1d),
					new ExpectedCellValue(sheetName, 5, 2, 99d),
					new ExpectedCellValue(sheetName, 10, 2, 6d),
					new ExpectedCellValue(sheetName, 11, 2, 1194d),
					new ExpectedCellValue(sheetName, 12, 2, 7d),
					new ExpectedCellValue(sheetName, 13, 2, 1293d),
					new ExpectedCellValue(sheetName, 2, 3, "January"),
					new ExpectedCellValue(sheetName, 4, 3, 1d),
					new ExpectedCellValue(sheetName, 5, 3, 415.75),
					new ExpectedCellValue(sheetName, 7, 3, 2d),
					new ExpectedCellValue(sheetName, 8, 3, 831.5),
					new ExpectedCellValue(sheetName, 10, 3, 2d),
					new ExpectedCellValue(sheetName, 11, 3, 831.5),
					new ExpectedCellValue(sheetName, 12, 3, 5d),
					new ExpectedCellValue(sheetName, 13, 3, 2078.75),
					new ExpectedCellValue(sheetName, 2, 4, "March"),
					new ExpectedCellValue(sheetName, 7, 4, 1d),
					new ExpectedCellValue(sheetName, 8, 4, 24.99),
					new ExpectedCellValue(sheetName, 10, 4, 2d),
					new ExpectedCellValue(sheetName, 11, 4, 831.5),
					new ExpectedCellValue(sheetName, 12, 4, 3d),
					new ExpectedCellValue(sheetName, 13, 4, 856.49),
					new ExpectedCellValue(sheetName, 2, 5, "Grand Total"),
					new ExpectedCellValue(sheetName, 4, 5, 2d),
					new ExpectedCellValue(sheetName, 5, 5, 514.75),
					new ExpectedCellValue(sheetName, 7, 5, 3d),
					new ExpectedCellValue(sheetName, 8, 5, 856.49),
					new ExpectedCellValue(sheetName, 10, 5, 10d),
					new ExpectedCellValue(sheetName, 11, 5, 2857d),
					new ExpectedCellValue(sheetName, 12, 5, 15d),
					new ExpectedCellValue(sheetName, 13, 5, 4228.24)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableSorting.xlsx")]
		public void PivotTableRefreshCustomSortingWithReferencesToDataFieldsRowFieldsWithDataFieldsSortingRowsColumnsDataFields()
		{
			var file = new FileInfo("PivotTableSorting.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["DataFieldsInRows"];
					var pivotTable = worksheet.PivotTables["PivotTable2"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("A17:E29"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "DataFieldsInRows";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 19, 1, "Nashville"),
					new ExpectedCellValue(sheetName, 20, 1, "Sum of Total"),
					new ExpectedCellValue(sheetName, 21, 1, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 22, 1, "Chicago"),
					new ExpectedCellValue(sheetName, 23, 1, "Sum of Total"),
					new ExpectedCellValue(sheetName, 24, 1, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 25, 1, "San Francisco"),
					new ExpectedCellValue(sheetName, 26, 1, "Sum of Total"),
					new ExpectedCellValue(sheetName, 27, 1, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 28, 1, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 29, 1, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 18, 2, "March"),
					new ExpectedCellValue(sheetName, 20, 2, 831.5),
					new ExpectedCellValue(sheetName, 21, 2, 2d),
					new ExpectedCellValue(sheetName, 23, 2, 24.99),
					new ExpectedCellValue(sheetName, 24, 2, 1d),
					new ExpectedCellValue(sheetName, 28, 2, 856.49),
					new ExpectedCellValue(sheetName, 29, 2, 3d),
					new ExpectedCellValue(sheetName, 18, 3, "February"),
					new ExpectedCellValue(sheetName, 20, 3, 1194d),
					new ExpectedCellValue(sheetName, 21, 3, 6d),
					new ExpectedCellValue(sheetName, 26, 3, 99d),
					new ExpectedCellValue(sheetName, 27, 3, 1d),
					new ExpectedCellValue(sheetName, 28, 3, 1293d),
					new ExpectedCellValue(sheetName, 29, 3, 7d),
					new ExpectedCellValue(sheetName, 18, 4, "January"),
					new ExpectedCellValue(sheetName, 20, 4, 831.5),
					new ExpectedCellValue(sheetName, 21, 4, 2d),
					new ExpectedCellValue(sheetName, 23, 4, 831.5),
					new ExpectedCellValue(sheetName, 24, 4, 2d),
					new ExpectedCellValue(sheetName, 26, 4, 415.75),
					new ExpectedCellValue(sheetName, 27, 4, 1d),
					new ExpectedCellValue(sheetName, 28, 4, 2078.75),
					new ExpectedCellValue(sheetName, 29, 4, 5d),
					new ExpectedCellValue(sheetName, 18, 5, "Grand Total"),
					new ExpectedCellValue(sheetName, 20, 5, 2857d),
					new ExpectedCellValue(sheetName, 21, 5, 10d),
					new ExpectedCellValue(sheetName, 23, 5, 856.49),
					new ExpectedCellValue(sheetName, 24, 5, 3d),
					new ExpectedCellValue(sheetName, 26, 5, 514.75),
					new ExpectedCellValue(sheetName, 27, 5, 2d),
					new ExpectedCellValue(sheetName, 28, 5, 4228.24),
					new ExpectedCellValue(sheetName, 29, 5, 15d),
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableSorting.xlsx")]
		public void PivotTableRefreshCustomSortingWithReferencesToDataFieldsRowFieldsWithDataFieldsAndTwoColumnFieldsSortingRowsColumnsDataFields()
		{
			var file = new FileInfo("PivotTableSorting.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["DataFieldsInRows"];
					var pivotTable = worksheet.PivotTables["PivotTable3"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("A33:J46"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "DataFieldsInRows";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 36, 1, "San Francisco"),
					new ExpectedCellValue(sheetName, 37, 1, "Sum of Total"),
					new ExpectedCellValue(sheetName, 38, 1, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 39, 1, "Nashville"),
					new ExpectedCellValue(sheetName, 40, 1, "Sum of Total"),
					new ExpectedCellValue(sheetName, 41, 1, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 42, 1, "Chicago"),
					new ExpectedCellValue(sheetName, 43, 1, "Sum of Total"),
					new ExpectedCellValue(sheetName, 44, 1, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 45, 1, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 46, 1, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 34, 2, "February"),
					new ExpectedCellValue(sheetName, 35, 2, "Tent"),
					new ExpectedCellValue(sheetName, 36, 2, null),
					new ExpectedCellValue(sheetName, 37, 2, null),
					new ExpectedCellValue(sheetName, 38, 2, null),
					new ExpectedCellValue(sheetName, 39, 2, null),
					new ExpectedCellValue(sheetName, 40, 2, 1194d),
					new ExpectedCellValue(sheetName, 41, 2, 6d),
					new ExpectedCellValue(sheetName, 42, 2, null),
					new ExpectedCellValue(sheetName, 43, 2, null),
					new ExpectedCellValue(sheetName, 44, 2, null),
					new ExpectedCellValue(sheetName, 45, 2, 1194d),
					new ExpectedCellValue(sheetName, 46, 2, 6d),
					new ExpectedCellValue(sheetName, 35, 3, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 36, 3, null),
					new ExpectedCellValue(sheetName, 37, 3, 99d),
					new ExpectedCellValue(sheetName, 38, 3, 1d),
					new ExpectedCellValue(sheetName, 39, 3, null),
					new ExpectedCellValue(sheetName, 40, 3, null),
					new ExpectedCellValue(sheetName, 41, 3, null),
					new ExpectedCellValue(sheetName, 42, 3, null),
					new ExpectedCellValue(sheetName, 43, 3, null),
					new ExpectedCellValue(sheetName, 44, 3, null),
					new ExpectedCellValue(sheetName, 45, 3, 99d),
					new ExpectedCellValue(sheetName, 46, 3, 1d),
					new ExpectedCellValue(sheetName, 34, 4, "February Total"),
					new ExpectedCellValue(sheetName, 35, 4, null),
					new ExpectedCellValue(sheetName, 36, 4, null),
					new ExpectedCellValue(sheetName, 37, 4, 99d),
					new ExpectedCellValue(sheetName, 38, 4, 1d),
					new ExpectedCellValue(sheetName, 39, 4, null),
					new ExpectedCellValue(sheetName, 40, 4, 1194d),
					new ExpectedCellValue(sheetName, 41, 4, 6d),
					new ExpectedCellValue(sheetName, 42, 4, null),
					new ExpectedCellValue(sheetName, 43, 4, null),
					new ExpectedCellValue(sheetName, 44, 4, null),
					new ExpectedCellValue(sheetName, 45, 4, 1293d),
					new ExpectedCellValue(sheetName, 46, 4, 7d),
					new ExpectedCellValue(sheetName, 34, 5, "January"),
					new ExpectedCellValue(sheetName, 35, 5, "Car Rack"),
					new ExpectedCellValue(sheetName, 36, 5, null),
					new ExpectedCellValue(sheetName, 37, 5, 415.75),
					new ExpectedCellValue(sheetName, 38, 5, 1d),
					new ExpectedCellValue(sheetName, 39, 5, null),
					new ExpectedCellValue(sheetName, 40, 5, 831.5),
					new ExpectedCellValue(sheetName, 41, 5, 2d),
					new ExpectedCellValue(sheetName, 42, 5, null),
					new ExpectedCellValue(sheetName, 43, 5, 831.5),
					new ExpectedCellValue(sheetName, 44, 5, 2d),
					new ExpectedCellValue(sheetName, 45, 5, 2078.75),
					new ExpectedCellValue(sheetName, 46, 5, 5d),
					new ExpectedCellValue(sheetName, 34, 6, "January Total"),
					new ExpectedCellValue(sheetName, 35, 6, null),
					new ExpectedCellValue(sheetName, 36, 6, null),
					new ExpectedCellValue(sheetName, 37, 6, 415.75),
					new ExpectedCellValue(sheetName, 38, 6, 1d),
					new ExpectedCellValue(sheetName, 39, 6, null),
					new ExpectedCellValue(sheetName, 40, 6, 831.5),
					new ExpectedCellValue(sheetName, 41, 6, 2d),
					new ExpectedCellValue(sheetName, 42, 6, null),
					new ExpectedCellValue(sheetName, 43, 6, 831.5),
					new ExpectedCellValue(sheetName, 44, 6, 2d),
					new ExpectedCellValue(sheetName, 45, 6, 2078.75),
					new ExpectedCellValue(sheetName, 46, 6, 5d),
					new ExpectedCellValue(sheetName, 34, 7, "March"),
					new ExpectedCellValue(sheetName, 35, 7, "Headlamp"),
					new ExpectedCellValue(sheetName, 36, 7, null),
					new ExpectedCellValue(sheetName, 37, 7, null),
					new ExpectedCellValue(sheetName, 38, 7, null),
					new ExpectedCellValue(sheetName, 39, 7, null),
					new ExpectedCellValue(sheetName, 40, 7, null),
					new ExpectedCellValue(sheetName, 41, 7, null),
					new ExpectedCellValue(sheetName, 42, 7, null),
					new ExpectedCellValue(sheetName, 43, 7, 24.99),
					new ExpectedCellValue(sheetName, 44, 7, 1d),
					new ExpectedCellValue(sheetName, 45, 7, 24.99),
					new ExpectedCellValue(sheetName, 46, 7, 1d),
					new ExpectedCellValue(sheetName, 35, 8, "Car Rack"),
					new ExpectedCellValue(sheetName, 36, 8, null),
					new ExpectedCellValue(sheetName, 37, 8, null),
					new ExpectedCellValue(sheetName, 38, 8, null),
					new ExpectedCellValue(sheetName, 39, 8, null),
					new ExpectedCellValue(sheetName, 40, 8, 831.5),
					new ExpectedCellValue(sheetName, 41, 8, 2d),
					new ExpectedCellValue(sheetName, 42, 8, null),
					new ExpectedCellValue(sheetName, 43, 8, null),
					new ExpectedCellValue(sheetName, 44, 8, null),
					new ExpectedCellValue(sheetName, 45, 8, 831.5),
					new ExpectedCellValue(sheetName, 46, 8, 2d),
					new ExpectedCellValue(sheetName, 34, 9, "March Total"),
					new ExpectedCellValue(sheetName, 35, 9, null),
					new ExpectedCellValue(sheetName, 36, 9, null),
					new ExpectedCellValue(sheetName, 37, 9, null),
					new ExpectedCellValue(sheetName, 38, 9, null),
					new ExpectedCellValue(sheetName, 39, 9, null),
					new ExpectedCellValue(sheetName, 40, 9, 831.5),
					new ExpectedCellValue(sheetName, 41, 9, 2d),
					new ExpectedCellValue(sheetName, 42, 9, null),
					new ExpectedCellValue(sheetName, 43, 9, 24.99),
					new ExpectedCellValue(sheetName, 44, 9, 1d),
					new ExpectedCellValue(sheetName, 45, 9, 856.49),
					new ExpectedCellValue(sheetName, 46, 9, 3d),
					new ExpectedCellValue(sheetName, 34, 10, "Grand Total"),
					new ExpectedCellValue(sheetName, 35, 10, null),
					new ExpectedCellValue(sheetName, 36, 10, null),
					new ExpectedCellValue(sheetName, 37, 10, 514.75),
					new ExpectedCellValue(sheetName, 38, 10, 2d),
					new ExpectedCellValue(sheetName, 39, 10, null),
					new ExpectedCellValue(sheetName, 40, 10, 2857d),
					new ExpectedCellValue(sheetName, 41, 10, 10d),
					new ExpectedCellValue(sheetName, 42, 10, null),
					new ExpectedCellValue(sheetName, 43, 10, 856.49),
					new ExpectedCellValue(sheetName, 44, 10, 3d),
					new ExpectedCellValue(sheetName, 45, 10, 4228.24),
					new ExpectedCellValue(sheetName, 46, 10, 15d)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableSorting.xlsx")]
		public void PivotTableRefreshCustomSortingWithReferencesToDataFieldsTwoRowFieldsWithDataFieldsAndTwoColumnFieldsSortingRowsColumnsDataFields()
		{
			var file = new FileInfo("PivotTableSorting.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["DataFieldsInRows"];
					var pivotTable = worksheet.PivotTables["PivotTable4"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					this.CheckPivotTableAddress(new ExcelAddress("A50:J84"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(8, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					package.SaveAs(newFile.File);
				}
				string sheetName = "DataFieldsInRows";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 53, 1, "Nashville"),
					new ExpectedCellValue(sheetName, 54, 1, 20100090),
					new ExpectedCellValue(sheetName, 55, 1, "Sum of Total"),
					new ExpectedCellValue(sheetName, 56, 1, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 57, 1, 20100070),
					new ExpectedCellValue(sheetName, 58, 1, "Sum of Total"),
					new ExpectedCellValue(sheetName, 59, 1, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 60, 1, 20100017),
					new ExpectedCellValue(sheetName, 61, 1, "Sum of Total"),
					new ExpectedCellValue(sheetName, 62, 1, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 63, 1, "Nashville Sum of Total"),
					new ExpectedCellValue(sheetName, 64, 1, "Nashville Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 65, 1, "Chicago"),
					new ExpectedCellValue(sheetName, 66, 1, 20100083),
					new ExpectedCellValue(sheetName, 67, 1, "Sum of Total"),
					new ExpectedCellValue(sheetName, 68, 1, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 69, 1, 20100007),
					new ExpectedCellValue(sheetName, 70, 1, "Sum of Total"),
					new ExpectedCellValue(sheetName, 71, 1, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 72, 1, "Chicago Sum of Total"),
					new ExpectedCellValue(sheetName, 73, 1, "Chicago Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 74, 1, "San Francisco"),
					new ExpectedCellValue(sheetName, 75, 1, 20100085),
					new ExpectedCellValue(sheetName, 76, 1, "Sum of Total"),
					new ExpectedCellValue(sheetName, 77, 1, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 78, 1, 20100076),
					new ExpectedCellValue(sheetName, 79, 1, "Sum of Total"),
					new ExpectedCellValue(sheetName, 80, 1, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 81, 1, "San Francisco Sum of Total"),
					new ExpectedCellValue(sheetName, 82, 1, "San Francisco Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 83, 1, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 84, 1, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 51, 2, "February"),
					new ExpectedCellValue(sheetName, 52, 2, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 53, 2, null),
					new ExpectedCellValue(sheetName, 54, 2, null),
					new ExpectedCellValue(sheetName, 55, 2, null),
					new ExpectedCellValue(sheetName, 56, 2, null),
					new ExpectedCellValue(sheetName, 57, 2, null),
					new ExpectedCellValue(sheetName, 58, 2, null),
					new ExpectedCellValue(sheetName, 59, 2, null),
					new ExpectedCellValue(sheetName, 60, 2, null),
					new ExpectedCellValue(sheetName, 61, 2, null),
					new ExpectedCellValue(sheetName, 62, 2, null),
					new ExpectedCellValue(sheetName, 63, 2, null),
					new ExpectedCellValue(sheetName, 64, 2, null),
					new ExpectedCellValue(sheetName, 65, 2, null),
					new ExpectedCellValue(sheetName, 66, 2, null),
					new ExpectedCellValue(sheetName, 67, 2, null),
					new ExpectedCellValue(sheetName, 68, 2, null),
					new ExpectedCellValue(sheetName, 69, 2, null),
					new ExpectedCellValue(sheetName, 70, 2, null),
					new ExpectedCellValue(sheetName, 71, 2, null),
					new ExpectedCellValue(sheetName, 72, 2, null),
					new ExpectedCellValue(sheetName, 73, 2, null),
					new ExpectedCellValue(sheetName, 74, 2, null),
					new ExpectedCellValue(sheetName, 75, 2, null),
					new ExpectedCellValue(sheetName, 76, 2, 99d),
					new ExpectedCellValue(sheetName, 77, 2, 1d),
					new ExpectedCellValue(sheetName, 78, 2, null),
					new ExpectedCellValue(sheetName, 79, 2, null),
					new ExpectedCellValue(sheetName, 80, 2, null),
					new ExpectedCellValue(sheetName, 81, 2, 99d),
					new ExpectedCellValue(sheetName, 82, 2, 1d),
					new ExpectedCellValue(sheetName, 83, 2, 99d),
					new ExpectedCellValue(sheetName, 84, 2, 1d),
					new ExpectedCellValue(sheetName, 52, 3, "Tent"),
					new ExpectedCellValue(sheetName, 53, 3, null),
					new ExpectedCellValue(sheetName, 54, 3, null),
					new ExpectedCellValue(sheetName, 55, 3, null),
					new ExpectedCellValue(sheetName, 56, 3, null),
					new ExpectedCellValue(sheetName, 57, 3, null),
					new ExpectedCellValue(sheetName, 58, 3, 1194d),
					new ExpectedCellValue(sheetName, 59, 3, 6d),
					new ExpectedCellValue(sheetName, 60, 3, null),
					new ExpectedCellValue(sheetName, 61, 3, null),
					new ExpectedCellValue(sheetName, 62, 3, null),
					new ExpectedCellValue(sheetName, 63, 3, 1194d),
					new ExpectedCellValue(sheetName, 64, 3, 6d),
					new ExpectedCellValue(sheetName, 65, 3, null),
					new ExpectedCellValue(sheetName, 66, 3, null),
					new ExpectedCellValue(sheetName, 67, 3, null),
					new ExpectedCellValue(sheetName, 68, 3, null),
					new ExpectedCellValue(sheetName, 69, 3, null),
					new ExpectedCellValue(sheetName, 70, 3, null),
					new ExpectedCellValue(sheetName, 71, 3, null),
					new ExpectedCellValue(sheetName, 72, 3, null),
					new ExpectedCellValue(sheetName, 73, 3, null),
					new ExpectedCellValue(sheetName, 74, 3, null),
					new ExpectedCellValue(sheetName, 75, 3, null),
					new ExpectedCellValue(sheetName, 76, 3, null),
					new ExpectedCellValue(sheetName, 77, 3, null),
					new ExpectedCellValue(sheetName, 78, 3, null),
					new ExpectedCellValue(sheetName, 79, 3, null),
					new ExpectedCellValue(sheetName, 80, 3, null),
					new ExpectedCellValue(sheetName, 81, 3, null),
					new ExpectedCellValue(sheetName, 82, 3, null),
					new ExpectedCellValue(sheetName, 83, 3, 1194d),
					new ExpectedCellValue(sheetName, 84, 3, 6d),
					new ExpectedCellValue(sheetName, 51, 4, "February Total"),
					new ExpectedCellValue(sheetName, 52, 4, null),
					new ExpectedCellValue(sheetName, 53, 4, null),
					new ExpectedCellValue(sheetName, 54, 4, null),
					new ExpectedCellValue(sheetName, 55, 4, null),
					new ExpectedCellValue(sheetName, 56, 4, null),
					new ExpectedCellValue(sheetName, 57, 4, null),
					new ExpectedCellValue(sheetName, 58, 4, 1194d),
					new ExpectedCellValue(sheetName, 59, 4, 6d),
					new ExpectedCellValue(sheetName, 60, 4, null),
					new ExpectedCellValue(sheetName, 61, 4, null),
					new ExpectedCellValue(sheetName, 62, 4, null),
					new ExpectedCellValue(sheetName, 63, 4, 1194d),
					new ExpectedCellValue(sheetName, 64, 4, 6d),
					new ExpectedCellValue(sheetName, 65, 4, null),
					new ExpectedCellValue(sheetName, 66, 4, null),
					new ExpectedCellValue(sheetName, 67, 4, null),
					new ExpectedCellValue(sheetName, 68, 4, null),
					new ExpectedCellValue(sheetName, 69, 4, null),
					new ExpectedCellValue(sheetName, 70, 4, null),
					new ExpectedCellValue(sheetName, 71, 4, null),
					new ExpectedCellValue(sheetName, 72, 4, null),
					new ExpectedCellValue(sheetName, 73, 4, null),
					new ExpectedCellValue(sheetName, 74, 4, null),
					new ExpectedCellValue(sheetName, 75, 4, null),
					new ExpectedCellValue(sheetName, 76, 4, 99d),
					new ExpectedCellValue(sheetName, 77, 4, 1d),
					new ExpectedCellValue(sheetName, 78, 4, null),
					new ExpectedCellValue(sheetName, 79, 4, null),
					new ExpectedCellValue(sheetName, 80, 4, null),
					new ExpectedCellValue(sheetName, 81, 4, 99d),
					new ExpectedCellValue(sheetName, 82, 4, 1d),
					new ExpectedCellValue(sheetName, 83, 4, 1293d),
					new ExpectedCellValue(sheetName, 84, 4, 7d),
					new ExpectedCellValue(sheetName, 51, 5, "January"),
					new ExpectedCellValue(sheetName, 52, 5, "Car Rack"),
					new ExpectedCellValue(sheetName, 53, 5, null),
					new ExpectedCellValue(sheetName, 54, 5, null),
					new ExpectedCellValue(sheetName, 55, 5, 831.5),
					new ExpectedCellValue(sheetName, 56, 5, 2d),
					new ExpectedCellValue(sheetName, 57, 5, null),
					new ExpectedCellValue(sheetName, 58, 5, null),
					new ExpectedCellValue(sheetName, 59, 5, null),
					new ExpectedCellValue(sheetName, 60, 5, null),
					new ExpectedCellValue(sheetName, 61, 5, null),
					new ExpectedCellValue(sheetName, 62, 5, null),
					new ExpectedCellValue(sheetName, 63, 5, 831.5),
					new ExpectedCellValue(sheetName, 64, 5, 2d),
					new ExpectedCellValue(sheetName, 65, 5, null),
					new ExpectedCellValue(sheetName, 66, 5, null),
					new ExpectedCellValue(sheetName, 67, 5, null),
					new ExpectedCellValue(sheetName, 68, 5, null),
					new ExpectedCellValue(sheetName, 69, 5, null),
					new ExpectedCellValue(sheetName, 70, 5, 831.5),
					new ExpectedCellValue(sheetName, 71, 5, 2d),
					new ExpectedCellValue(sheetName, 72, 5, 831.5),
					new ExpectedCellValue(sheetName, 73, 5, 2d),
					new ExpectedCellValue(sheetName, 74, 5, null),
					new ExpectedCellValue(sheetName, 75, 5, null),
					new ExpectedCellValue(sheetName, 76, 5, null),
					new ExpectedCellValue(sheetName, 77, 5, null),
					new ExpectedCellValue(sheetName, 78, 5, null),
					new ExpectedCellValue(sheetName, 79, 5, 415.75),
					new ExpectedCellValue(sheetName, 80, 5, 1d),
					new ExpectedCellValue(sheetName, 81, 5, 415.75),
					new ExpectedCellValue(sheetName, 82, 5, 1d),
					new ExpectedCellValue(sheetName, 83, 5, 2078.75),
					new ExpectedCellValue(sheetName, 84, 5, 5d),
					new ExpectedCellValue(sheetName, 51, 6, "January Total"),
					new ExpectedCellValue(sheetName, 52, 6, null),
					new ExpectedCellValue(sheetName, 53, 6, null),
					new ExpectedCellValue(sheetName, 54, 6, null),
					new ExpectedCellValue(sheetName, 55, 6, 831.5),
					new ExpectedCellValue(sheetName, 56, 6, 2d),
					new ExpectedCellValue(sheetName, 57, 6, null),
					new ExpectedCellValue(sheetName, 58, 6, null),
					new ExpectedCellValue(sheetName, 59, 6, null),
					new ExpectedCellValue(sheetName, 60, 6, null),
					new ExpectedCellValue(sheetName, 61, 6, null),
					new ExpectedCellValue(sheetName, 62, 6, null),
					new ExpectedCellValue(sheetName, 63, 6, 831.5),
					new ExpectedCellValue(sheetName, 64, 6, 2d),
					new ExpectedCellValue(sheetName, 65, 6, null),
					new ExpectedCellValue(sheetName, 66, 6, null),
					new ExpectedCellValue(sheetName, 67, 6, null),
					new ExpectedCellValue(sheetName, 68, 6, null),
					new ExpectedCellValue(sheetName, 69, 6, null),
					new ExpectedCellValue(sheetName, 70, 6, 831.5),
					new ExpectedCellValue(sheetName, 71, 6, 2d),
					new ExpectedCellValue(sheetName, 72, 6, 831.5),
					new ExpectedCellValue(sheetName, 73, 6, 2d),
					new ExpectedCellValue(sheetName, 74, 6, null),
					new ExpectedCellValue(sheetName, 75, 6, null),
					new ExpectedCellValue(sheetName, 76, 6, null),
					new ExpectedCellValue(sheetName, 77, 6, null),
					new ExpectedCellValue(sheetName, 78, 6, null),
					new ExpectedCellValue(sheetName, 79, 6, 415.75),
					new ExpectedCellValue(sheetName, 80, 6, 1d),
					new ExpectedCellValue(sheetName, 81, 6, 415.75),
					new ExpectedCellValue(sheetName, 82, 6, 1d),
					new ExpectedCellValue(sheetName, 83, 6, 2078.75),
					new ExpectedCellValue(sheetName, 84, 6, 5d),
					new ExpectedCellValue(sheetName, 51, 7, "March"),
					new ExpectedCellValue(sheetName, 52, 7, "Headlamp"),
					new ExpectedCellValue(sheetName, 53, 7, null),
					new ExpectedCellValue(sheetName, 54, 7, null),
					new ExpectedCellValue(sheetName, 55, 7, null),
					new ExpectedCellValue(sheetName, 56, 7, null),
					new ExpectedCellValue(sheetName, 57, 7, null),
					new ExpectedCellValue(sheetName, 58, 7, null),
					new ExpectedCellValue(sheetName, 59, 7, null),
					new ExpectedCellValue(sheetName, 60, 7, null),
					new ExpectedCellValue(sheetName, 61, 7, null),
					new ExpectedCellValue(sheetName, 62, 7, null),
					new ExpectedCellValue(sheetName, 63, 7, null),
					new ExpectedCellValue(sheetName, 64, 7, null),
					new ExpectedCellValue(sheetName, 65, 7, null),
					new ExpectedCellValue(sheetName, 66, 7, null),
					new ExpectedCellValue(sheetName, 67, 7, 24.99),
					new ExpectedCellValue(sheetName, 68, 7, 1d),
					new ExpectedCellValue(sheetName, 69, 7, null),
					new ExpectedCellValue(sheetName, 70, 7, null),
					new ExpectedCellValue(sheetName, 71, 7, null),
					new ExpectedCellValue(sheetName, 72, 7, 24.99),
					new ExpectedCellValue(sheetName, 73, 7, 1d),
					new ExpectedCellValue(sheetName, 74, 7, null),
					new ExpectedCellValue(sheetName, 75, 7, null),
					new ExpectedCellValue(sheetName, 76, 7, null),
					new ExpectedCellValue(sheetName, 77, 7, null),
					new ExpectedCellValue(sheetName, 78, 7, null),
					new ExpectedCellValue(sheetName, 79, 7, null),
					new ExpectedCellValue(sheetName, 80, 7, null),
					new ExpectedCellValue(sheetName, 81, 7, null),
					new ExpectedCellValue(sheetName, 82, 7, null),
					new ExpectedCellValue(sheetName, 83, 7, 24.99),
					new ExpectedCellValue(sheetName, 84, 7, 1d),
					new ExpectedCellValue(sheetName, 52, 8, "Car Rack"),
					new ExpectedCellValue(sheetName, 53, 8, null),
					new ExpectedCellValue(sheetName, 54, 8, null),
					new ExpectedCellValue(sheetName, 55, 8, null),
					new ExpectedCellValue(sheetName, 56, 8, null),
					new ExpectedCellValue(sheetName, 57, 8, null),
					new ExpectedCellValue(sheetName, 58, 8, null),
					new ExpectedCellValue(sheetName, 59, 8, null),
					new ExpectedCellValue(sheetName, 60, 8, null),
					new ExpectedCellValue(sheetName, 61, 8, 831.5),
					new ExpectedCellValue(sheetName, 62, 8, 2d),
					new ExpectedCellValue(sheetName, 63, 8, 831.5),
					new ExpectedCellValue(sheetName, 64, 8, 2d),
					new ExpectedCellValue(sheetName, 65, 8, null),
					new ExpectedCellValue(sheetName, 66, 8, null),
					new ExpectedCellValue(sheetName, 67, 8, null),
					new ExpectedCellValue(sheetName, 68, 8, null),
					new ExpectedCellValue(sheetName, 69, 8, null),
					new ExpectedCellValue(sheetName, 70, 8, null),
					new ExpectedCellValue(sheetName, 71, 8, null),
					new ExpectedCellValue(sheetName, 72, 8, null),
					new ExpectedCellValue(sheetName, 73, 8, null),
					new ExpectedCellValue(sheetName, 74, 8, null),
					new ExpectedCellValue(sheetName, 75, 8, null),
					new ExpectedCellValue(sheetName, 76, 8, null),
					new ExpectedCellValue(sheetName, 77, 8, null),
					new ExpectedCellValue(sheetName, 78, 8, null),
					new ExpectedCellValue(sheetName, 79, 8, null),
					new ExpectedCellValue(sheetName, 80, 8, null),
					new ExpectedCellValue(sheetName, 81, 8, null),
					new ExpectedCellValue(sheetName, 82, 8, null),
					new ExpectedCellValue(sheetName, 83, 8, 831.5),
					new ExpectedCellValue(sheetName, 84, 8, 2d),
					new ExpectedCellValue(sheetName, 51, 9, "March Total"),
					new ExpectedCellValue(sheetName, 52, 9, null),
					new ExpectedCellValue(sheetName, 53, 9, null),
					new ExpectedCellValue(sheetName, 54, 9, null),
					new ExpectedCellValue(sheetName, 55, 9, null),
					new ExpectedCellValue(sheetName, 56, 9, null),
					new ExpectedCellValue(sheetName, 57, 9, null),
					new ExpectedCellValue(sheetName, 58, 9, null),
					new ExpectedCellValue(sheetName, 59, 9, null),
					new ExpectedCellValue(sheetName, 60, 9, null),
					new ExpectedCellValue(sheetName, 61, 9, 831.5),
					new ExpectedCellValue(sheetName, 62, 9, 2d),
					new ExpectedCellValue(sheetName, 63, 9, 831.5),
					new ExpectedCellValue(sheetName, 64, 9, 2d),
					new ExpectedCellValue(sheetName, 65, 9, null),
					new ExpectedCellValue(sheetName, 66, 9, null),
					new ExpectedCellValue(sheetName, 67, 9, 24.99),
					new ExpectedCellValue(sheetName, 68, 9, 1d),
					new ExpectedCellValue(sheetName, 69, 9, null),
					new ExpectedCellValue(sheetName, 70, 9, null),
					new ExpectedCellValue(sheetName, 71, 9, null),
					new ExpectedCellValue(sheetName, 72, 9, 24.99),
					new ExpectedCellValue(sheetName, 73, 9, 1d),
					new ExpectedCellValue(sheetName, 74, 9, null),
					new ExpectedCellValue(sheetName, 75, 9, null),
					new ExpectedCellValue(sheetName, 76, 9, null),
					new ExpectedCellValue(sheetName, 77, 9, null),
					new ExpectedCellValue(sheetName, 78, 9, null),
					new ExpectedCellValue(sheetName, 79, 9, null),
					new ExpectedCellValue(sheetName, 80, 9, null),
					new ExpectedCellValue(sheetName, 81, 9, null),
					new ExpectedCellValue(sheetName, 82, 9, null),
					new ExpectedCellValue(sheetName, 83, 9, 856.49),
					new ExpectedCellValue(sheetName, 84, 9, 3d),
					new ExpectedCellValue(sheetName, 51, 10, "Grand Total"),
					new ExpectedCellValue(sheetName, 52, 10, null),
					new ExpectedCellValue(sheetName, 53, 10, null),
					new ExpectedCellValue(sheetName, 54, 10, null),
					new ExpectedCellValue(sheetName, 55, 10, 831.5),
					new ExpectedCellValue(sheetName, 56, 10, 2d),
					new ExpectedCellValue(sheetName, 57, 10, null),
					new ExpectedCellValue(sheetName, 58, 10, 1194d),
					new ExpectedCellValue(sheetName, 59, 10, 6d),
					new ExpectedCellValue(sheetName, 60, 10, null),
					new ExpectedCellValue(sheetName, 61, 10, 831.5),
					new ExpectedCellValue(sheetName, 62, 10, 2d),
					new ExpectedCellValue(sheetName, 63, 10, 2857d),
					new ExpectedCellValue(sheetName, 64, 10, 10d),
					new ExpectedCellValue(sheetName, 65, 10, null),
					new ExpectedCellValue(sheetName, 66, 10, null),
					new ExpectedCellValue(sheetName, 67, 10, 24.99),
					new ExpectedCellValue(sheetName, 68, 10, 1d),
					new ExpectedCellValue(sheetName, 69, 10, null),
					new ExpectedCellValue(sheetName, 70, 10, 831.5),
					new ExpectedCellValue(sheetName, 71, 10, 2d),
					new ExpectedCellValue(sheetName, 72, 10, 856.49),
					new ExpectedCellValue(sheetName, 73, 10, 3d),
					new ExpectedCellValue(sheetName, 74, 10, null),
					new ExpectedCellValue(sheetName, 75, 10, null),
					new ExpectedCellValue(sheetName, 76, 10, 99d),
					new ExpectedCellValue(sheetName, 77, 10, 1d),
					new ExpectedCellValue(sheetName, 78, 10, null),
					new ExpectedCellValue(sheetName, 79, 10, 415.75),
					new ExpectedCellValue(sheetName, 80, 10, 1d),
					new ExpectedCellValue(sheetName, 81, 10, 514.75),
					new ExpectedCellValue(sheetName, 82, 10, 2d),
					new ExpectedCellValue(sheetName, 83, 10, 4228.24),
					new ExpectedCellValue(sheetName, 84, 10, 15d)
				});
			}
		}
		#endregion

		#region Helper Methods
		private void CheckPivotTableAddress(ExcelAddress expectedAddress, ExcelAddress pivotTableAddress)
		{
			Assert.AreEqual(expectedAddress.Start.Row, pivotTableAddress.Start.Row);
			Assert.AreEqual(expectedAddress.Start.Column, pivotTableAddress.Start.Column);
			Assert.AreEqual(expectedAddress.End.Row, pivotTableAddress.End.Row);
			Assert.AreEqual(expectedAddress.End.Column, pivotTableAddress.End.Column);
		}
		#endregion
	}
}