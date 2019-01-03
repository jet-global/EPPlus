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
using OfficeOpenXml.Table.PivotTable;

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
		[DeploymentItem(@"..\..\Workbooks\PivotTableDataSourceTypeWorksheet.xlsx")]
		public void PivotTableRefreshFromCacheWithChangedData()
		{
			var file = new FileInfo("PivotTableDataSourceTypeWorksheet.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var worksheet = package.Workbook.Worksheets.First();
				var pivotTable = worksheet.PivotTables.First();
				var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
				worksheet.Cells[4, 5].Value = "Blue";
				worksheet.Cells[5, 5].Value = "Green";
				worksheet.Cells[6, 5].Value = "Purple";
				cacheDefinition.UpdateData();
				Assert.AreEqual(4, pivotTable.Fields.Count);
				Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
				Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
				Assert.AreEqual(6, pivotTable.Fields[2].Items.Count);
				Assert.AreEqual(0, pivotTable.Fields[3].Items.Count);
				foreach (var field in pivotTable.Fields)
				{
					if (field.Items.Count > 0)
						this.CheckFieldItems(field);
				}
				Assert.AreEqual(7, pivotTable.RowItems.Count);
				Assert.AreEqual("Blue", worksheet.Cells[11, 9].Value);
				Assert.AreEqual(100d, worksheet.Cells[11, 10].Value);
				Assert.AreEqual("Bike", worksheet.Cells[12, 9].Value);
				Assert.AreEqual(100d, worksheet.Cells[12, 10].Value);
				Assert.AreEqual("Green", worksheet.Cells[13, 9].Value);
				Assert.AreEqual(90000d, worksheet.Cells[13, 10].Value);
				Assert.AreEqual("Car", worksheet.Cells[14, 9].Value);
				Assert.AreEqual(90000d, worksheet.Cells[14, 10].Value);
				Assert.AreEqual("Purple", worksheet.Cells[15, 9].Value);
				Assert.AreEqual(10d, worksheet.Cells[15, 10].Value);
				Assert.AreEqual("Skateboard", worksheet.Cells[16, 9].Value);
				Assert.AreEqual(10d, worksheet.Cells[16, 10].Value);
				Assert.AreEqual("Grand Total", worksheet.Cells[17, 9].Value);
				Assert.AreEqual(90110d, worksheet.Cells[17, 10].Value);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableDataSourceTypeWorksheet.xlsx")]
		public void PivotTableRefreshFromCacheWithAddedData()
		{
			var file = new FileInfo("PivotTableDataSourceTypeWorksheet.xlsx");
			Assert.IsTrue(file.Exists);
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
				Assert.AreEqual(4, pivotTable.Fields.Count);
				var pivotField1 = pivotTable.Fields[0];
				Assert.AreEqual(0, pivotField1.Items.Count);
				var pivotField2 = pivotTable.Fields[1];
				Assert.AreEqual(5, pivotField2.Items.Count);
				var pivotField3 = pivotTable.Fields[2];
				Assert.AreEqual(4, pivotField3.Items.Count);
				this.CheckFieldItems(pivotField3);
				var pivotField4 = pivotTable.Fields[3];
				Assert.AreEqual(0, pivotField4.Items.Count);
				Assert.AreEqual(8, pivotTable.RowItems.Count);
				Assert.AreEqual("Black", worksheet.Cells[11, 9].Value);
				Assert.AreEqual(110d, worksheet.Cells[11, 10].Value);
				Assert.AreEqual("Bike", worksheet.Cells[12, 9].Value);
				Assert.AreEqual(100d, worksheet.Cells[12, 10].Value);
				Assert.AreEqual("Skateboard", worksheet.Cells[13, 9].Value);
				Assert.AreEqual(10d, worksheet.Cells[13, 10].Value);
				Assert.AreEqual("Red", worksheet.Cells[14, 9].Value);
				Assert.AreEqual(90000d, worksheet.Cells[14, 10].Value);
				Assert.AreEqual("Car", worksheet.Cells[15, 9].Value);
				Assert.AreEqual(90000d, worksheet.Cells[15, 10].Value);
				Assert.AreEqual("Purple", worksheet.Cells[16, 9].Value);
				Assert.AreEqual(28d, worksheet.Cells[16, 10].Value);
				Assert.AreEqual("Scooter", worksheet.Cells[17, 9].Value);
				Assert.AreEqual(28d, worksheet.Cells[17, 10].Value);
				Assert.AreEqual("Grand Total", worksheet.Cells[18, 9].Value);
				Assert.AreEqual(90138d, worksheet.Cells[18, 10].Value);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableDataSourceTypeWorksheet.xlsx")]
		public void PivotTableRefreshFromCacheRemoveRow()
		{
			var file = new FileInfo("PivotTableDataSourceTypeWorksheet.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var worksheet = package.Workbook.Worksheets.First();
				var pivotTable = worksheet.PivotTables.First();
				var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
				cacheDefinition.SetSourceRangeAddress(worksheet, worksheet.Cells["C3:F5"]);
				cacheDefinition.UpdateData();
				Assert.AreEqual(4, pivotTable.Fields.Count);
				var pivotField1 = pivotTable.Fields[0];
				Assert.AreEqual(0, pivotField1.Items.Count);
				var pivotField2 = pivotTable.Fields[1];
				Assert.AreEqual(4, pivotField2.Items.Count);
				var pivotField3 = pivotTable.Fields[2];
				Assert.AreEqual(3, pivotField3.Items.Count);
				this.CheckFieldItems(pivotField3);
				var pivotField4 = pivotTable.Fields[3];
				Assert.AreEqual(0, pivotField4.Items.Count);
				Assert.AreEqual(5, pivotTable.RowItems.Count);
				Assert.AreEqual("Black", worksheet.Cells[11, 9].Value);
				Assert.AreEqual(100d, worksheet.Cells[11, 10].Value);
				Assert.AreEqual("Bike", worksheet.Cells[12, 9].Value);
				Assert.AreEqual(100d, worksheet.Cells[12, 10].Value);
				Assert.AreEqual("Red", worksheet.Cells[13, 9].Value);
				Assert.AreEqual(90000d, worksheet.Cells[13, 10].Value);
				Assert.AreEqual("Car", worksheet.Cells[14, 9].Value);
				Assert.AreEqual(90000d, worksheet.Cells[14, 10].Value);
				Assert.AreEqual("Grand Total", worksheet.Cells[15, 9].Value);
				Assert.AreEqual(90100d, worksheet.Cells[15, 10].Value);
				Assert.IsNull(worksheet.Cells[16, 9].Value);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableColumnFields.xlsx")]
		public void PivotTableRefreshColumnItemsWithChangedData()
		{
			var file = new FileInfo("PivotTableColumnFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var worksheet = package.Workbook.Worksheets.First();
				var pivotTable = worksheet.PivotTables.First();
				var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
				worksheet.Cells[4, 3].Value = "January";
				worksheet.Cells[7, 3].Value = "January";
				cacheDefinition.UpdateData();
				Assert.AreEqual(7, pivotTable.Fields.Count);
				Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
				Assert.AreEqual(0, pivotTable.Fields[1].Items.Count);
				Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
				Assert.AreEqual(0, pivotTable.Fields[3].Items.Count);
				Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
				Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
				Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
				foreach (var field in pivotTable.Fields)
				{
					if (field.Items.Count > 0)
						this.CheckFieldItems(field);
				}
				Assert.AreEqual("January", worksheet.Cells[13, 3].Value);
				Assert.AreEqual("Car Rack", worksheet.Cells[14, 3].Value);
				Assert.AreEqual("San Francisco", worksheet.Cells[15, 3].Value);
				Assert.AreEqual("Chicago", worksheet.Cells[15, 4].Value);
				Assert.AreEqual("Nashville", worksheet.Cells[15, 5].Value);
				Assert.AreEqual("Car Rack Total", worksheet.Cells[14, 6].Value);
				Assert.AreEqual("Headlamp", worksheet.Cells[14, 7].Value);
				Assert.AreEqual("Chicago", worksheet.Cells[15, 7].Value);
				Assert.AreEqual("Headlamp Total", worksheet.Cells[14, 8].Value);
				Assert.AreEqual("January Total", worksheet.Cells[13, 9].Value);
				Assert.AreEqual("February", worksheet.Cells[13, 10].Value);
				Assert.AreEqual("Sleeping Bag", worksheet.Cells[14, 10].Value);
				Assert.AreEqual("San Francisco", worksheet.Cells[15, 10].Value);
				Assert.AreEqual("Sleeping Bag Total", worksheet.Cells[14, 11].Value);
				Assert.AreEqual("Tent", worksheet.Cells[14, 12].Value);
				Assert.AreEqual("Nashville", worksheet.Cells[15, 12].Value);
				Assert.AreEqual("Tent Total", worksheet.Cells[14, 13].Value);
				Assert.AreEqual("February Total", worksheet.Cells[13, 14].Value);
				Assert.AreEqual("Grand Total", worksheet.Cells[13, 15].Value);
			}
		}
		
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableColumnFields.xlsx")]
		public void PivotTableRefreshColumnItemsWithAddedData()
		{
			var file = new FileInfo("PivotTableColumnFields.xlsx");
			Assert.IsTrue(file.Exists);
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
						this.CheckFieldItems(field);
				}
				Assert.AreEqual("20100076", worksheet.Cells[16, 2].Value);
				Assert.AreEqual("20100085", worksheet.Cells[17, 2].Value);
				Assert.AreEqual("20100083", worksheet.Cells[18, 2].Value);
				Assert.AreEqual("20100007", worksheet.Cells[19, 2].Value);
				Assert.AreEqual("20100070", worksheet.Cells[20, 2].Value);
				Assert.AreEqual("20100017", worksheet.Cells[21, 2].Value);
				Assert.AreEqual("20100090", worksheet.Cells[22, 2].Value);
				Assert.AreEqual("20100091", worksheet.Cells[23, 2].Value);
				Assert.AreEqual("January", worksheet.Cells[13, 3].Value);
				Assert.AreEqual("Car Rack", worksheet.Cells[14, 3].Value);
				Assert.AreEqual("San Francisco", worksheet.Cells[15, 3].Value);
				Assert.AreEqual("Chicago", worksheet.Cells[15, 4].Value);
				Assert.AreEqual("Nashville", worksheet.Cells[15, 5].Value);
				Assert.AreEqual("Car Rack Total", worksheet.Cells[14, 6].Value);
				Assert.AreEqual("January Total", worksheet.Cells[13, 7].Value);
				Assert.AreEqual("February", worksheet.Cells[13, 8].Value);
				Assert.AreEqual("Sleeping Bag", worksheet.Cells[14, 8].Value);
				Assert.AreEqual("San Francisco", worksheet.Cells[15, 8].Value);
				Assert.AreEqual("Sleeping Bag Total", worksheet.Cells[14, 9].Value);
				Assert.AreEqual("Tent", worksheet.Cells[14, 10].Value);
				Assert.AreEqual("Nashville", worksheet.Cells[15, 10].Value);
				Assert.AreEqual("Tent Total", worksheet.Cells[14, 11].Value);
				Assert.AreEqual("February Total", worksheet.Cells[13, 12].Value);
				Assert.AreEqual("March", worksheet.Cells[13, 13].Value);
				Assert.AreEqual("Car Rack", worksheet.Cells[14, 13].Value);
				Assert.AreEqual("Nashville", worksheet.Cells[15, 13].Value);
				Assert.AreEqual("Car Rack Total", worksheet.Cells[14, 14].Value);
				Assert.AreEqual("Headlamp", worksheet.Cells[14, 15].Value);
				Assert.AreEqual("Chicago", worksheet.Cells[15, 15].Value);
				Assert.AreEqual("Headlamp Total", worksheet.Cells[14, 16].Value);
				Assert.AreEqual("March Total", worksheet.Cells[13, 17].Value);
				Assert.AreEqual("December", worksheet.Cells[13, 18].Value);
				Assert.AreEqual("Bike", worksheet.Cells[14, 18].Value);
				Assert.AreEqual("Texas", worksheet.Cells[15, 18].Value);
				Assert.AreEqual("Bike Total", worksheet.Cells[14, 19].Value);
				Assert.AreEqual("December Total", worksheet.Cells[13, 20].Value);
				Assert.AreEqual("Grand Total", worksheet.Cells[13, 21].Value);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableColumnFields.xlsx")]
		public void PivotTableRefreshColumnItemsWithRemoveData()
		{
			var file = new FileInfo("PivotTableColumnFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var worksheet = package.Workbook.Worksheets.First();
				var pivotTable = worksheet.PivotTables["Sheet1PivotTable1"];
				var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
				cacheDefinition.SetSourceRangeAddress(worksheet, worksheet.Cells["A1:G5"]);
				cacheDefinition.UpdateData();
				Assert.AreEqual(7, pivotTable.Fields.Count);
				Assert.AreEqual(8, pivotTable.Fields[0].Items.Count);
				Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
				Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
				Assert.AreEqual(5, pivotTable.Fields[3].Items.Count);
				Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
				Assert.AreEqual(4, pivotTable.Fields[5].Items.Count);
				Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
				foreach (var field in pivotTable.Fields)
				{
					if (field.Items.Count > 0)
						this.CheckFieldItems(field);
				}
				Assert.AreEqual("January", worksheet.Cells[13, 3].Value);
				Assert.AreEqual("Car Rack", worksheet.Cells[14, 3].Value);
				Assert.AreEqual("San Francisco", worksheet.Cells[15, 3].Value);
				Assert.AreEqual("Chicago", worksheet.Cells[15, 4].Value);
				Assert.AreEqual("Car Rack Total", worksheet.Cells[14, 5].Value);
				Assert.AreEqual("January Total", worksheet.Cells[13, 6].Value);
				Assert.AreEqual("February", worksheet.Cells[13, 7].Value);
				Assert.AreEqual("Sleeping Bag", worksheet.Cells[14, 7].Value);
				Assert.AreEqual("San Francisco", worksheet.Cells[15, 7].Value);
				Assert.AreEqual("Sleeping Bag Total", worksheet.Cells[14, 8].Value);
				Assert.AreEqual("February Total", worksheet.Cells[13, 9].Value);
				Assert.AreEqual("March", worksheet.Cells[13, 10].Value);
				Assert.AreEqual("Headlamp", worksheet.Cells[14, 10].Value);
				Assert.AreEqual("Chicago", worksheet.Cells[15, 10].Value);
				Assert.AreEqual("Headlamp Total", worksheet.Cells[14, 11].Value);
				Assert.AreEqual("March Total", worksheet.Cells[13, 12].Value);
				Assert.AreEqual("Grand Total", worksheet.Cells[13, 13].Value);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableColumnFields.xlsx")]
		public void PivotTableRefreshDeletingSourceRow()
		{
			var file = new FileInfo("PivotTableColumnFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets.First();
					var pivotTable = worksheet.PivotTables["Sheet1PivotTable1"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					worksheet.DeleteRow(6);
					cacheDefinition.SetSourceRangeAddress(worksheet, worksheet.Cells["A1:G7"]);
					cacheDefinition.UpdateData();
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(8, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					foreach (var field in pivotTable.Fields)
					{
						if (field.Items.Count > 0)
							this.CheckFieldItems(field);
					}
					package.SaveAs(newFile.File);
				}
				string sheetName = "Sheet1";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 15, 2, "20100076"),
					new ExpectedCellValue(sheetName, 16, 2, "20100085"),
					new ExpectedCellValue(sheetName, 17, 2, "20100083"),
					new ExpectedCellValue(sheetName, 18, 2, "20100007"),
					new ExpectedCellValue(sheetName, 19, 2, "20100017"),
					new ExpectedCellValue(sheetName, 20, 2, "20100090"),
					new ExpectedCellValue(sheetName, 12, 3, "January"),
					new ExpectedCellValue(sheetName, 13, 3, "Car Rack"),
					new ExpectedCellValue(sheetName, 14, 3, "San Francisco"),
					new ExpectedCellValue(sheetName, 14, 4, "Chicago"),
					new ExpectedCellValue(sheetName, 14, 5, "Nashville"),
					new ExpectedCellValue(sheetName, 13, 6, "Car Rack Total"),
					new ExpectedCellValue(sheetName, 12, 7, "January Total"),
					new ExpectedCellValue(sheetName, 12, 8, "February"),
					new ExpectedCellValue(sheetName, 13, 8, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 14, 8, "San Francisco"),
					new ExpectedCellValue(sheetName, 13, 9, "Sleeping Bag Total"),
					new ExpectedCellValue(sheetName, 12, 10, "February Total"),
					new ExpectedCellValue(sheetName, 12, 11, "March"),
					new ExpectedCellValue(sheetName, 13, 11, "Car Rack"),
					new ExpectedCellValue(sheetName, 14, 11, "Nashville"),
					new ExpectedCellValue(sheetName, 13, 12, "Car Rack Total"),
					new ExpectedCellValue(sheetName, 13, 13, "Headlamp"),
					new ExpectedCellValue(sheetName, 14, 13, "Chicago"),
					new ExpectedCellValue(sheetName, 13, 14, "Headlamp Total"),
					new ExpectedCellValue(sheetName, 12, 15, "March Total"),
					new ExpectedCellValue(sheetName, 12, 16, "Grand Total")
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableColumnFields.xlsx")]
		public void PivotTableRefreshSingleColumnNoDataFields()
		{
			var file = new FileInfo("PivotTableColumnFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["Sheet1"];
					var pivotTable = worksheet.PivotTables["Sheet1PivotTable2"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					foreach (var field in pivotTable.Fields)
					{
						if (field.Items.Count > 0)
							this.CheckFieldItems(field);
					}
					package.SaveAs(newFile.File);
				}
				string sheetName = "Sheet1";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 28, 2, "January"),
					new ExpectedCellValue(sheetName, 29, 2, "February"),
					new ExpectedCellValue(sheetName, 30, 2, "March"),
					new ExpectedCellValue(sheetName, 31, 2, "Grand Total")
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableColumnFields.xlsx")]
		public void PivotTableRefreshSingleColumnTwoRowFieldsAndNoDataFields()
		{
			var file = new FileInfo("PivotTableColumnFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["Sheet1"];
					var pivotTable = worksheet.PivotTables["Sheet1PivotTable3"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					foreach (var field in pivotTable.Fields)
					{
						if (field.Items.Count > 0)
							this.CheckFieldItems(field);
					}
					package.SaveAs(newFile.File);
				}
				string sheetName = "Sheet1";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 28, 5, "January"),
					new ExpectedCellValue(sheetName, 29, 5, "San Francisco"),
					new ExpectedCellValue(sheetName, 30, 5, "Chicago"),
					new ExpectedCellValue(sheetName, 31, 5, "Nashville"),
					new ExpectedCellValue(sheetName, 32, 5, "February"),
					new ExpectedCellValue(sheetName, 33, 5, "San Francisco"),
					new ExpectedCellValue(sheetName, 34, 5, "Nashville"),
					new ExpectedCellValue(sheetName, 35, 5, "March"),
					new ExpectedCellValue(sheetName, 36, 5, "Chicago"),
					new ExpectedCellValue(sheetName, 37, 5, "Nashville"),
					new ExpectedCellValue(sheetName, 38, 5, "Grand Total")
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableWithMAttrbuteFieldItems.xlsx")]
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
		#endregion

		#region UpdateData Field Values Tests
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableColumnFields.xlsx")]
		public void PivotTableRefreshDataFieldOneRowFieldWithTrueSubtotalTop()
		{
			var file = new FileInfo("PivotTableColumnFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["RowItems"];
					var pivotTable = worksheet.PivotTables["RowItemsPivotTable1"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					foreach (var field in pivotTable.Fields)
					{
						if (field.Items.Count > 0)
							this.CheckFieldItems(field);
					}
					package.SaveAs(newFile.File);
				}
				string sheetName = "RowItems";
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
		[DeploymentItem(@"..\..\Workbooks\PivotTableColumnFields.xlsx")]
		public void PivotTableRefreshDataFieldTwoRowFieldsWithTrueSubtotalTop()
		{
			var file = new FileInfo("PivotTableColumnFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["RowItems"];
					var pivotTable = worksheet.PivotTables["RowItemsPivotTable2"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					foreach (var field in pivotTable.Fields)
					{
						if (field.Items.Count > 0)
							this.CheckFieldItems(field);
					}
					package.SaveAs(newFile.File);
				}
				string sheetName = "RowItems";
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
		[DeploymentItem(@"..\..\Workbooks\PivotTableColumnFields.xlsx")]
		public void PivotTableRefreshDataFieldTwoRowFieldsWithFalseSubtotalTop()
		{
			var file = new FileInfo("PivotTableColumnFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["RowItems"];
					var pivotTable = worksheet.PivotTables["RowItemsPivotTable2"];
					foreach (var field in pivotTable.Fields)
					{
						field.SubtotalTop = false;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					foreach (var field in pivotTable.Fields)
					{
						if (field.Items.Count > 0)
							this.CheckFieldItems(field);
					}
					package.SaveAs(newFile.File);
				}
				string sheetName = "RowItems";
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
		[DeploymentItem(@"..\..\Workbooks\PivotTableColumnFields.xlsx")]
		public void PivotTableRefreshDataFieldTwoRowFieldsWithNoSubtotal()
		{
			var file = new FileInfo("PivotTableColumnFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["RowItems"];
					var pivotTable = worksheet.PivotTables["RowItemsPivotTable4"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(3, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					foreach (var field in pivotTable.Fields)
					{
						if (field.Items.Count > 0)
							this.CheckFieldItems(field);
					}
					package.SaveAs(newFile.File);
				}
				string sheetName = "RowItems";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 27, 1, "January"),
					new ExpectedCellValue(sheetName, 28, 1, "Car Rack"),
					new ExpectedCellValue(sheetName, 29, 1, "February"),
					new ExpectedCellValue(sheetName, 30, 1, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 31, 1, "Tent"),
					new ExpectedCellValue(sheetName, 32, 1, "March"),
					new ExpectedCellValue(sheetName, 33, 1, "Car Rack"),
					new ExpectedCellValue(sheetName, 34, 1, "Headlamp"),
					new ExpectedCellValue(sheetName, 35, 1, "Grand Total"),
					new ExpectedCellValue(sheetName, 28, 2, 2078.75),
					new ExpectedCellValue(sheetName, 30, 2, 99d),
					new ExpectedCellValue(sheetName, 31, 2, 1194d),
					new ExpectedCellValue(sheetName, 33, 2, 831.5),
					new ExpectedCellValue(sheetName, 34, 2, 24.99),
					new ExpectedCellValue(sheetName, 35, 2, 4228.24)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableColumnFields.xlsx")]
		public void PivotTableRefreshDataFieldThreeRowFieldsWithTrueSubtotalTop()
		{
			var file = new FileInfo("PivotTableColumnFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["RowItems"];
					var pivotTable = worksheet.PivotTables["RowItemsPivotTable3"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					foreach (var field in pivotTable.Fields)
					{
						if (field.Items.Count > 0)
							this.CheckFieldItems(field);
					}
					package.SaveAs(newFile.File);
				}
				string sheetName = "RowItems";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 2, 12, "January"),
					new ExpectedCellValue(sheetName, 3, 12, "Car Rack"),
					new ExpectedCellValue(sheetName, 4, 12, "San Francisco"),
					new ExpectedCellValue(sheetName, 5, 12, "Chicago"),
					new ExpectedCellValue(sheetName, 6, 12, "Nashville"),
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
					new ExpectedCellValue(sheetName, 4, 13, 415.75),
					new ExpectedCellValue(sheetName, 5, 13, 831.5),
					new ExpectedCellValue(sheetName, 6, 13, 831.5),
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
		[DeploymentItem(@"..\..\Workbooks\PivotTableColumnFields.xlsx")]
		public void PivotTableRefreshDataFieldThreeRowFieldsWithFalseSubtotalTop()
		{
			var file = new FileInfo("PivotTableColumnFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["RowItems"];
					var pivotTable = worksheet.PivotTables["RowItemsPivotTable3"];
					foreach (var field in pivotTable.Fields)
					{
						field.SubtotalTop = false;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					foreach (var field in pivotTable.Fields)
					{
						if (field.Items.Count > 0)
							this.CheckFieldItems(field);
					}
					package.SaveAs(newFile.File);
				}
				string sheetName = "RowItems";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 2, 12, "January"),
					new ExpectedCellValue(sheetName, 3, 12, "Car Rack"),
					new ExpectedCellValue(sheetName, 4, 12, "San Francisco"),
					new ExpectedCellValue(sheetName, 5, 12, "Chicago"),
					new ExpectedCellValue(sheetName, 6, 12, "Nashville"),
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
					new ExpectedCellValue(sheetName, 4, 13, 415.75),
					new ExpectedCellValue(sheetName, 5, 13, 831.5),
					new ExpectedCellValue(sheetName, 6, 13, 831.5),
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
		[DeploymentItem(@"..\..\Workbooks\PivotTableColumnFields.xlsx")]
		public void PivotTableRefreshDataFieldsRowsAndColumnsWithNoSubtotal()
		{
			var file = new FileInfo("PivotTableColumnFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["NoSubtotals"];
					var pivotTable = worksheet.PivotTables["NoSubtotalsPivotTable1"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(3, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(3, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(3, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					foreach (var field in pivotTable.Fields)
					{
						if (field.Items.Count > 0)
							this.CheckFieldItems(field);
					}
					package.SaveAs(newFile.File);
				}
				string sheetName = "NoSubtotals";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 4, 1, "Car Rack"),
					new ExpectedCellValue(sheetName, 5, 1, "1"),
					new ExpectedCellValue(sheetName, 6, 1, "2"),
					new ExpectedCellValue(sheetName, 7, 1, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 8, 1, "1"),
					new ExpectedCellValue(sheetName, 9, 1, "Headlamp"),
					new ExpectedCellValue(sheetName, 10, 1, "1"),
					new ExpectedCellValue(sheetName, 11, 1, "Tent"),
					new ExpectedCellValue(sheetName, 12, 1, "6"),
					new ExpectedCellValue(sheetName, 13, 1, "Grand Total"),
					new ExpectedCellValue(sheetName, 5, 2, 415.75),
					new ExpectedCellValue(sheetName, 13, 2, 415.75),
					new ExpectedCellValue(sheetName, 8, 3, 99d),
					new ExpectedCellValue(sheetName, 13, 3, 99d),
					new ExpectedCellValue(sheetName, 6, 4, 415.75),
					new ExpectedCellValue(sheetName, 13, 4, 415.75),
					new ExpectedCellValue(sheetName, 10, 5, 24.99),
					new ExpectedCellValue(sheetName, 13, 5, 24.99),
					new ExpectedCellValue(sheetName, 6, 6, 415.75),
					new ExpectedCellValue(sheetName, 13, 6, 415.75),
					new ExpectedCellValue(sheetName, 12, 7, 199d),
					new ExpectedCellValue(sheetName, 13, 7, 199d),
					new ExpectedCellValue(sheetName, 6, 8, 415.75),
					new ExpectedCellValue(sheetName, 13, 8, 415.75)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableColumnFields.xlsx")]
		public void PivotTableRefreshDataFieldsRowsAndColumnsGrandTotalOff()
		{
			var file = new FileInfo("PivotTableColumnFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["GrandTotals"];
					var pivotTable = worksheet.PivotTables["GrandTotalsPivotTable1"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					pivotTable.RowGrandTotals = false;
					pivotTable.ColumnGrandTotals = false;
					cacheDefinition.UpdateData();
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					foreach (var field in pivotTable.Fields)
					{
						if (field.Items.Count > 0)
							this.CheckFieldItems(field);
					}
					package.SaveAs(newFile.File);
				}
				string sheetName = "GrandTotals";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 4, 1, "Car Rack"),
					new ExpectedCellValue(sheetName, 5, 1, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 6, 1, "Headlamp"),
					new ExpectedCellValue(sheetName, 7, 1, "Tent"),
					new ExpectedCellValue(sheetName, 2, 2, "January"),
					new ExpectedCellValue(sheetName, 3, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 4, 2, 415.75),
					new ExpectedCellValue(sheetName, 3, 3, "Chicago"),
					new ExpectedCellValue(sheetName, 4, 3, 831.5),
					new ExpectedCellValue(sheetName, 3, 4, "Nashville"),
					new ExpectedCellValue(sheetName, 4, 4, 831.5),
					new ExpectedCellValue(sheetName, 2, 5, "January Total"),
					new ExpectedCellValue(sheetName, 4, 5, 2078.75),
					new ExpectedCellValue(sheetName, 2, 6, "February"),
					new ExpectedCellValue(sheetName, 3, 6, "San Francisco"),
					new ExpectedCellValue(sheetName, 5, 6, 99d),
					new ExpectedCellValue(sheetName, 3, 7, "Nashville"),
					new ExpectedCellValue(sheetName, 7, 7, 1194d),
					new ExpectedCellValue(sheetName, 2, 8, "February Total"),
					new ExpectedCellValue(sheetName, 5, 8, 99d),
					new ExpectedCellValue(sheetName, 7, 8, 1194d),
					new ExpectedCellValue(sheetName, 2, 9, "March"),
					new ExpectedCellValue(sheetName, 3, 9, "Chicago"),
					new ExpectedCellValue(sheetName, 6, 9, 24.99),
					new ExpectedCellValue(sheetName, 3, 10, "Nashville"),
					new ExpectedCellValue(sheetName, 4, 10, 831.5),
					new ExpectedCellValue(sheetName, 2, 11, "March Total"),
					new ExpectedCellValue(sheetName, 4, 11, 831.5),
					new ExpectedCellValue(sheetName, 6, 11, 24.99)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableColumnFields.xlsx")]
		public void PivotTableRefreshDataFieldsColumnGrandTotalOff()
		{
			var file = new FileInfo("PivotTableColumnFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["GrandTotals"];
					var pivotTable = worksheet.PivotTables["GrandTotalsPivotTable1"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					pivotTable.ColumnGrandTotals = false;
					cacheDefinition.UpdateData();
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					foreach (var field in pivotTable.Fields)
					{
						if (field.Items.Count > 0)
							this.CheckFieldItems(field);
					}
					package.SaveAs(newFile.File);
				}
				string sheetName = "GrandTotals";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 4, 1, "Car Rack"),
					new ExpectedCellValue(sheetName, 5, 1, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 6, 1, "Headlamp"),
					new ExpectedCellValue(sheetName, 7, 1, "Tent"),
					new ExpectedCellValue(sheetName, 8, 1, "Grand Total"),
					new ExpectedCellValue(sheetName, 2, 2, "January"),
					new ExpectedCellValue(sheetName, 3, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 4, 2, 415.75),
					new ExpectedCellValue(sheetName, 8, 2, 415.75),
					new ExpectedCellValue(sheetName, 3, 3, "Chicago"),
					new ExpectedCellValue(sheetName, 4, 3, 831.5),
					new ExpectedCellValue(sheetName, 8, 3, 831.5),
					new ExpectedCellValue(sheetName, 3, 4, "Nashville"),
					new ExpectedCellValue(sheetName, 4, 4, 831.5),
					new ExpectedCellValue(sheetName, 8, 4, 831.5),
					new ExpectedCellValue(sheetName, 2, 5, "January Total"),
					new ExpectedCellValue(sheetName, 4, 5, 2078.75),
					new ExpectedCellValue(sheetName, 8, 5, 2078.75),
					new ExpectedCellValue(sheetName, 2, 6, "February"),
					new ExpectedCellValue(sheetName, 3, 6, "San Francisco"),
					new ExpectedCellValue(sheetName, 5, 6, 99d),
					new ExpectedCellValue(sheetName, 8, 6, 99d),
					new ExpectedCellValue(sheetName, 3, 7, "Nashville"),
					new ExpectedCellValue(sheetName, 7, 7, 1194d),
					new ExpectedCellValue(sheetName, 8, 7, 1194d),
					new ExpectedCellValue(sheetName, 2, 8, "February Total"),
					new ExpectedCellValue(sheetName, 5, 8, 99d),
					new ExpectedCellValue(sheetName, 7, 8, 1194d),
					new ExpectedCellValue(sheetName, 8, 8, 1293d),
					new ExpectedCellValue(sheetName, 2, 9, "March"),
					new ExpectedCellValue(sheetName, 3, 9, "Chicago"),
					new ExpectedCellValue(sheetName, 6, 9, 24.99),
					new ExpectedCellValue(sheetName, 8, 9, 24.99),
					new ExpectedCellValue(sheetName, 3, 10, "Nashville"),
					new ExpectedCellValue(sheetName, 4, 10, 831.5),
					new ExpectedCellValue(sheetName, 8, 10, 831.5),
					new ExpectedCellValue(sheetName, 2, 11, "March Total"),
					new ExpectedCellValue(sheetName, 4, 11, 831.5),
					new ExpectedCellValue(sheetName, 6, 11, 24.99),
					new ExpectedCellValue(sheetName, 8, 11, 856.49)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableColumnFields.xlsx")]
		public void PivotTableRefreshDataFieldsRowGrandTotalOff()
		{
			var file = new FileInfo("PivotTableColumnFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["GrandTotals"];
					var pivotTable = worksheet.PivotTables["GrandTotalsPivotTable1"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					pivotTable.RowGrandTotals = false;
					cacheDefinition.UpdateData();
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					foreach (var field in pivotTable.Fields)
					{
						if (field.Items.Count > 0)
							this.CheckFieldItems(field);
					}
					package.SaveAs(newFile.File);
				}
				string sheetName = "GrandTotals";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 4, 1, "Car Rack"),
					new ExpectedCellValue(sheetName, 5, 1, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 6, 1, "Headlamp"),
					new ExpectedCellValue(sheetName, 7, 1, "Tent"),
					new ExpectedCellValue(sheetName, 2, 2, "January"),
					new ExpectedCellValue(sheetName, 3, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 4, 2, 415.75),
					new ExpectedCellValue(sheetName, 3, 3, "Chicago"),
					new ExpectedCellValue(sheetName, 4, 3, 831.5),
					new ExpectedCellValue(sheetName, 3, 4, "Nashville"),
					new ExpectedCellValue(sheetName, 4, 4, 831.5),
					new ExpectedCellValue(sheetName, 2, 5, "January Total"),
					new ExpectedCellValue(sheetName, 4, 5, 2078.75),
					new ExpectedCellValue(sheetName, 2, 6, "February"),
					new ExpectedCellValue(sheetName, 3, 6, "San Francisco"),
					new ExpectedCellValue(sheetName, 5, 6, 99d),
					new ExpectedCellValue(sheetName, 3, 7, "Nashville"),
					new ExpectedCellValue(sheetName, 7, 7, 1194d),
					new ExpectedCellValue(sheetName, 2, 8, "February Total"),
					new ExpectedCellValue(sheetName, 5, 8, 99d),
					new ExpectedCellValue(sheetName, 7, 8, 1194d),
					new ExpectedCellValue(sheetName, 2, 9, "March"),
					new ExpectedCellValue(sheetName, 3, 9, "Chicago"),
					new ExpectedCellValue(sheetName, 6, 9, 24.99),
					new ExpectedCellValue(sheetName, 3, 10, "Nashville"),
					new ExpectedCellValue(sheetName, 4, 10, 831.5),
					new ExpectedCellValue(sheetName, 2, 11, "March Total"),
					new ExpectedCellValue(sheetName, 4, 11, 831.5),
					new ExpectedCellValue(sheetName, 6, 11, 24.99d),
					new ExpectedCellValue(sheetName, 2, 12, "Grand Total"),
					new ExpectedCellValue(sheetName, 4, 12, 2910.25),
					new ExpectedCellValue(sheetName, 5, 12, 99d),
					new ExpectedCellValue(sheetName, 6, 12, 24.99),
					new ExpectedCellValue(sheetName, 7, 12, 1194d)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableColumnFields.xlsx")]
		public void PivotTableRefreshMultipleDataFields()
		{
			var file = new FileInfo("PivotTableColumnFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["MultipleDataFields"];
					var pivotTable = worksheet.PivotTables["MultipleDataFieldsPivotTable1"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					Assert.AreEqual(1, pivotTable.Address.Start.Row);
					Assert.AreEqual(1, pivotTable.Address.Start.Column);
					Assert.AreEqual(5, pivotTable.Address.End.Row);
					Assert.AreEqual(3, pivotTable.Address.End.Column);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					foreach (var field in pivotTable.Fields)
					{
						if (field.Items.Count > 0)
							this.CheckFieldItems(field);
					}
					package.SaveAs(newFile.File);
				}
				string sheetName = "MultipleDataFields";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 2, 1, "San Francisco"),
					new ExpectedCellValue(sheetName, 3, 1, "Chicago"),
					new ExpectedCellValue(sheetName, 4, 1, "Nashville"),
					new ExpectedCellValue(sheetName, 5, 1, "Grand Total"),
					new ExpectedCellValue(sheetName, 1, 2, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 2, 2, 2),
					new ExpectedCellValue(sheetName, 3, 2, 3),
					new ExpectedCellValue(sheetName, 4, 2, 10),
					new ExpectedCellValue(sheetName, 5, 2, 15),
					new ExpectedCellValue(sheetName, 1, 3, "Sum of Total"),
					new ExpectedCellValue(sheetName, 2, 3, 514.75),
					new ExpectedCellValue(sheetName, 3, 3, 856.49),
					new ExpectedCellValue(sheetName, 4, 3, 2857d),
					new ExpectedCellValue(sheetName, 5, 3, 4228.24),
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableColumnFields.xlsx")]
		public void PivotTableRefreshMultipleDataFieldsNoGrandTotal()
		{
			var file = new FileInfo("PivotTableColumnFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["MultipleDataFields"];
					var pivotTable = worksheet.PivotTables["MultipleDataFieldsPivotTable1"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					pivotTable.RowGrandTotals = false;
					cacheDefinition.UpdateData();
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					foreach (var field in pivotTable.Fields)
					{
						if (field.Items.Count > 0)
							this.CheckFieldItems(field);
					}
					package.SaveAs(newFile.File);
				}
				string sheetName = "MultipleDataFields";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 2, 1, "San Francisco"),
					new ExpectedCellValue(sheetName, 3, 1, "Chicago"),
					new ExpectedCellValue(sheetName, 4, 1, "Nashville"),
					new ExpectedCellValue(sheetName, 1, 2, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 2, 2, 2),
					new ExpectedCellValue(sheetName, 3, 2, 3),
					new ExpectedCellValue(sheetName, 4, 2, 10),
					new ExpectedCellValue(sheetName, 1, 3, "Sum of Total"),
					new ExpectedCellValue(sheetName, 2, 3, 514.75),
					new ExpectedCellValue(sheetName, 3, 3, 856.49),
					new ExpectedCellValue(sheetName, 4, 3, 2857d),
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableColumnFields.xlsx")]
		public void PivotTableRefreshMultipleDataFieldsAndColumnHeaders()
		{
			var file = new FileInfo("PivotTableColumnFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["MultipleDataFields"];
					var pivotTable = worksheet.PivotTables["MultipleDataFieldsPivotTable2"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					foreach (var field in pivotTable.Fields)
					{
						if (field.Items.Count > 0)
							this.CheckFieldItems(field);
					}
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
					new ExpectedCellValue(sheetName, 12, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 13, 2, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 14, 2, 1d),
					new ExpectedCellValue(sheetName, 17, 2, 1d),
					new ExpectedCellValue(sheetName, 13, 3, "Sum of Total"),
					new ExpectedCellValue(sheetName, 14, 3, 415.75),
					new ExpectedCellValue(sheetName, 17, 3, 415.75),
					new ExpectedCellValue(sheetName, 12, 4, "Chicago"),
					new ExpectedCellValue(sheetName, 13, 4, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 14, 4, 2d),
					new ExpectedCellValue(sheetName, 17, 4, 2d),
					new ExpectedCellValue(sheetName, 13, 5, "Sum of Total"),
					new ExpectedCellValue(sheetName, 14, 5, 831.5),
					new ExpectedCellValue(sheetName, 17, 5, 831.5),
					new ExpectedCellValue(sheetName, 12, 6, "Nashville"),
					new ExpectedCellValue(sheetName, 13, 6, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 14, 6, 2d),
					new ExpectedCellValue(sheetName, 16, 6, 2d),
					new ExpectedCellValue(sheetName, 17, 6, 4d),
					new ExpectedCellValue(sheetName, 13, 7, "Sum of Total"),
					new ExpectedCellValue(sheetName, 14, 7, 831.5),
					new ExpectedCellValue(sheetName, 16, 7, 831.5),
					new ExpectedCellValue(sheetName, 17, 7, 1663d),
					new ExpectedCellValue(sheetName, 11, 8, "Car Rack Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 14, 8, 5d),
					new ExpectedCellValue(sheetName, 16, 8, 2d),
					new ExpectedCellValue(sheetName, 17, 8, 7d),
					new ExpectedCellValue(sheetName, 11, 9, "Car Rack Sum of Total"),
					new ExpectedCellValue(sheetName, 14, 9, 2078.75),
					new ExpectedCellValue(sheetName, 16, 9, 831.5),
					new ExpectedCellValue(sheetName, 17, 9, 2910.25),
					new ExpectedCellValue(sheetName, 11, 10, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 12, 10, "San Francisco"),
					new ExpectedCellValue(sheetName, 13, 10, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 15, 10, 1d),
					new ExpectedCellValue(sheetName, 17, 10, 1d),
					new ExpectedCellValue(sheetName, 13, 11, "Sum of Total"),
					new ExpectedCellValue(sheetName, 15, 11, 99d),
					new ExpectedCellValue(sheetName, 17, 11, 99d),
					new ExpectedCellValue(sheetName, 11, 12, "Sleeping Bag Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 15, 12, 1d),
					new ExpectedCellValue(sheetName, 17, 12, 1d),
					new ExpectedCellValue(sheetName, 11, 13, "Sleeping Bag Sum of Total"),
					new ExpectedCellValue(sheetName, 15, 13, 99d),
					new ExpectedCellValue(sheetName, 17, 13, 99d),
					new ExpectedCellValue(sheetName, 11, 14, "Headlamp"),
					new ExpectedCellValue(sheetName, 12, 14, "Chicago"),
					new ExpectedCellValue(sheetName, 13, 14, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 16, 14, 1d),
					new ExpectedCellValue(sheetName, 17, 14, 1d),
					new ExpectedCellValue(sheetName, 13, 15, "Sum of Total"),
					new ExpectedCellValue(sheetName, 16, 15, 24.99),
					new ExpectedCellValue(sheetName, 17, 15, 24.99),
					new ExpectedCellValue(sheetName, 11, 16, "Headlamp Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 16, 16, 1d),
					new ExpectedCellValue(sheetName, 17, 16, 1d),
					new ExpectedCellValue(sheetName, 11, 17, "Headlamp Sum of Total"),
					new ExpectedCellValue(sheetName, 16, 17, 24.99),
					new ExpectedCellValue(sheetName, 17, 17, 24.99),
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
		[DeploymentItem(@"..\..\Workbooks\PivotTableColumnFields.xlsx")]
		public void PivotTableRefreshMultipleRowDataFieldsOneRowAndOneColumn()
		{
			var file = new FileInfo("PivotTableColumnFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["RowDataFields"];
					var pivotTable = worksheet.PivotTables["RowDataFieldsPivotTable1"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					foreach (var field in pivotTable.Fields)
					{
						if (field.Items.Count > 0)
							this.CheckFieldItems(field);
					}
					package.SaveAs(newFile.File);
				}
				string sheetName = "RowDataFields";
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

					new ExpectedCellValue(sheetName, 2, 2, "January"),
					new ExpectedCellValue(sheetName, 4, 2, 1d),
					new ExpectedCellValue(sheetName, 5, 2, 415.75),
					new ExpectedCellValue(sheetName, 7, 2, 2d),
					new ExpectedCellValue(sheetName, 8, 2, 831.5),
					new ExpectedCellValue(sheetName, 10, 2, 2d),
					new ExpectedCellValue(sheetName, 11, 2, 831.5),
					new ExpectedCellValue(sheetName, 12, 2, 5d),
					new ExpectedCellValue(sheetName, 13, 2, 2078.75),

					new ExpectedCellValue(sheetName, 2, 3, "February"),
					new ExpectedCellValue(sheetName, 4, 3, 1d),
					new ExpectedCellValue(sheetName, 5, 3, 99d),
					new ExpectedCellValue(sheetName, 10, 3, 6d),
					new ExpectedCellValue(sheetName, 11, 3, 1194d),
					new ExpectedCellValue(sheetName, 12, 3, 7d),
					new ExpectedCellValue(sheetName, 13, 3, 1293d),

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
		[DeploymentItem(@"..\..\Workbooks\PivotTableColumnFields.xlsx")]
		public void PivotTableRefreshMultipleRowDataFieldsTwoRowsAndOneColumnSubtotalsOff()
		{
			var file = new FileInfo("PivotTableColumnFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["RowDataFields"];
					var pivotTable = worksheet.PivotTables["RowDataFieldsPivotTable2"];
					foreach (var field in pivotTable.Fields)
					{
						field.DisableDefaultSubtotal();
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(3, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(3, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					foreach (var field in pivotTable.Fields)
					{
						if (field.Items.Count > 0)
							this.CheckFieldItems(field);
					}
					package.SaveAs(newFile.File);
				}
				string sheetName = "RowDataFields";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 3, 8, "San Francisco"),
					new ExpectedCellValue(sheetName, 4, 8, "Car Rack"),
					new ExpectedCellValue(sheetName, 5, 8, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 6, 8, "Sum of Total"),
					new ExpectedCellValue(sheetName, 7, 8, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 8, 8, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 9, 8, "Sum of Total"),
					new ExpectedCellValue(sheetName, 10, 8, "Chicago"),
					new ExpectedCellValue(sheetName, 11, 8, "Car Rack"),
					new ExpectedCellValue(sheetName, 12, 8, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 13, 8, "Sum of Total"),
					new ExpectedCellValue(sheetName, 14, 8, "Headlamp"),
					new ExpectedCellValue(sheetName, 15, 8, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 16, 8, "Sum of Total"),
					new ExpectedCellValue(sheetName, 17, 8, "Nashville"),
					new ExpectedCellValue(sheetName, 18, 8, "Car Rack"),
					new ExpectedCellValue(sheetName, 19, 8, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 20, 8, "Sum of Total"),
					new ExpectedCellValue(sheetName, 21, 8, "Tent"),
					new ExpectedCellValue(sheetName, 22, 8, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 23, 8, "Sum of Total"),
					new ExpectedCellValue(sheetName, 24, 8, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 25, 8, "Total Sum of Total"),

					new ExpectedCellValue(sheetName, 2, 9, "January"),
					new ExpectedCellValue(sheetName, 5, 9, 1d),
					new ExpectedCellValue(sheetName, 6, 9, 415.75),
					new ExpectedCellValue(sheetName, 12, 9, 2d),
					new ExpectedCellValue(sheetName, 13, 9, 831.5),
					new ExpectedCellValue(sheetName, 19, 9, 2d),
					new ExpectedCellValue(sheetName, 20, 9, 831.5),
					new ExpectedCellValue(sheetName, 24, 9, 5d),
					new ExpectedCellValue(sheetName, 25, 9, 2078.75),

					new ExpectedCellValue(sheetName, 2, 10, "February"),
					new ExpectedCellValue(sheetName, 8, 10, 1d),
					new ExpectedCellValue(sheetName, 9, 10, 99d),
					new ExpectedCellValue(sheetName, 22, 10, 6d),
					new ExpectedCellValue(sheetName, 23, 10, 1194d),
					new ExpectedCellValue(sheetName, 24, 10, 7d),
					new ExpectedCellValue(sheetName, 25, 10, 1293d),

					new ExpectedCellValue(sheetName, 2, 11, "March"),
					new ExpectedCellValue(sheetName, 15, 11, 1d),
					new ExpectedCellValue(sheetName, 16, 11, 24.99),
					new ExpectedCellValue(sheetName, 19, 11, 2d),
					new ExpectedCellValue(sheetName, 20, 11, 831.5),
					new ExpectedCellValue(sheetName, 24, 11, 3d),
					new ExpectedCellValue(sheetName, 25, 11, 856.49),

					new ExpectedCellValue(sheetName, 2, 12, "Grand Total"),
					new ExpectedCellValue(sheetName, 5, 12, 1d),
					new ExpectedCellValue(sheetName, 6, 12, 415.75),
					new ExpectedCellValue(sheetName, 8, 12, 1d),
					new ExpectedCellValue(sheetName, 9, 12, 99d),
					new ExpectedCellValue(sheetName, 12, 12, 2d),
					new ExpectedCellValue(sheetName, 13, 12, 831.5),
					new ExpectedCellValue(sheetName, 15, 12, 1d),
					new ExpectedCellValue(sheetName, 16, 12, 24.99),
					new ExpectedCellValue(sheetName, 19, 12, 4d),
					new ExpectedCellValue(sheetName, 20, 12, 1663d),
					new ExpectedCellValue(sheetName, 22, 12, 6d),
					new ExpectedCellValue(sheetName, 23, 12, 1194d),
					new ExpectedCellValue(sheetName, 24, 12, 15d),
					new ExpectedCellValue(sheetName, 25, 12, 4228.24)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableColumnFields.xlsx")]
		public void PivotTableRefreshMultipleRowDataFieldsTwoRowsAndOneColumnSubtotalsOn()
		{
			var file = new FileInfo("PivotTableColumnFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["RowDataFields"];
					var pivotTable = worksheet.PivotTables["RowDataFieldsPivotTable2"];
					foreach (var field in pivotTable.Fields)
					{
						field.SubtotalTop = false;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					foreach (var field in pivotTable.Fields)
					{
						if (field.Items.Count > 0)
							this.CheckFieldItems(field);
					}
					package.SaveAs(newFile.File);
				}
				string sheetName = "RowDataFields";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 3, 8, "San Francisco"),
					new ExpectedCellValue(sheetName, 4, 8, "Car Rack"),
					new ExpectedCellValue(sheetName, 5, 8, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 6, 8, "Sum of Total"),
					new ExpectedCellValue(sheetName, 7, 8, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 8, 8, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 9, 8, "Sum of Total"),
					new ExpectedCellValue(sheetName, 10, 8, "San Francisco Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 11, 8, "San Francisco Sum of Total"),
					new ExpectedCellValue(sheetName, 12, 8, "Chicago"),
					new ExpectedCellValue(sheetName, 13, 8, "Car Rack"),
					new ExpectedCellValue(sheetName, 14, 8, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 15, 8, "Sum of Total"),
					new ExpectedCellValue(sheetName, 16, 8, "Headlamp"),
					new ExpectedCellValue(sheetName, 17, 8, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 18, 8, "Sum of Total"),
					new ExpectedCellValue(sheetName, 19, 8, "Chicago Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 20, 8, "Chicago Sum of Total"),
					new ExpectedCellValue(sheetName, 21, 8, "Nashville"),
					new ExpectedCellValue(sheetName, 22, 8, "Car Rack"),
					new ExpectedCellValue(sheetName, 23, 8, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 24, 8, "Sum of Total"),
					new ExpectedCellValue(sheetName, 25, 8, "Tent"),
					new ExpectedCellValue(sheetName, 26, 8, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 27, 8, "Sum of Total"),
					new ExpectedCellValue(sheetName, 28, 8, "Nashville Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 29, 8, "Nashville Sum of Total"),
					new ExpectedCellValue(sheetName, 30, 8, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 31, 8, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 2, 9, "January"),
					new ExpectedCellValue(sheetName, 5, 9, 1d),
					new ExpectedCellValue(sheetName, 6, 9, 415.75),
					new ExpectedCellValue(sheetName, 10, 9, 1d),
					new ExpectedCellValue(sheetName, 11, 9, 415.75),
					new ExpectedCellValue(sheetName, 14, 9, 2d),
					new ExpectedCellValue(sheetName, 15, 9, 831.5),
					new ExpectedCellValue(sheetName, 19, 9, 2d),
					new ExpectedCellValue(sheetName, 20, 9, 831.5),
					new ExpectedCellValue(sheetName, 23, 9, 2d),
					new ExpectedCellValue(sheetName, 24, 9, 831.5),
					new ExpectedCellValue(sheetName, 28, 9, 2d),
					new ExpectedCellValue(sheetName, 29, 9, 831.5),
					new ExpectedCellValue(sheetName, 30, 9, 5d),
					new ExpectedCellValue(sheetName, 31, 9, 2078.75),
					new ExpectedCellValue(sheetName, 2, 10, "February"),
					new ExpectedCellValue(sheetName, 8, 10, 1d),
					new ExpectedCellValue(sheetName, 9, 10, 99d),
					new ExpectedCellValue(sheetName, 10, 10, 1d),
					new ExpectedCellValue(sheetName, 11, 10, 99d),
					new ExpectedCellValue(sheetName, 26, 10, 6d),
					new ExpectedCellValue(sheetName, 27, 10, 1194d),
					new ExpectedCellValue(sheetName, 28, 10, 6d),
					new ExpectedCellValue(sheetName, 29, 10, 1194d),
					new ExpectedCellValue(sheetName, 30, 10, 7d),
					new ExpectedCellValue(sheetName, 31, 10, 1293d),
					new ExpectedCellValue(sheetName, 2, 11, "March"),
					new ExpectedCellValue(sheetName, 17, 11, 1d),
					new ExpectedCellValue(sheetName, 18, 11, 24.99),
					new ExpectedCellValue(sheetName, 19, 11, 1d),
					new ExpectedCellValue(sheetName, 20, 11, 24.99),
					new ExpectedCellValue(sheetName, 23, 11, 2d),
					new ExpectedCellValue(sheetName, 24, 11, 831.5),
					new ExpectedCellValue(sheetName, 28, 11, 2d),
					new ExpectedCellValue(sheetName, 29, 11, 831.5),
					new ExpectedCellValue(sheetName, 30, 11, 3d),
					new ExpectedCellValue(sheetName, 31, 11, 856.49),
					new ExpectedCellValue(sheetName, 2, 12, "Grand Total"),
					new ExpectedCellValue(sheetName, 5, 12, 1d),
					new ExpectedCellValue(sheetName, 6, 12, 415.75),
					new ExpectedCellValue(sheetName, 8, 12, 1d),
					new ExpectedCellValue(sheetName, 9, 12, 99d),
					new ExpectedCellValue(sheetName, 10, 12, 2d),
					new ExpectedCellValue(sheetName, 11, 12, 514.75),
					new ExpectedCellValue(sheetName, 14, 12, 2d),
					new ExpectedCellValue(sheetName, 15, 12, 831.5),
					new ExpectedCellValue(sheetName, 17, 12, 1d),
					new ExpectedCellValue(sheetName, 18, 12, 24.99),
					new ExpectedCellValue(sheetName, 19, 12, 3d),
					new ExpectedCellValue(sheetName, 20, 12, 856.49),
					new ExpectedCellValue(sheetName, 23, 12, 4d),
					new ExpectedCellValue(sheetName, 24, 12, 1663d),
					new ExpectedCellValue(sheetName, 26, 12, 6d),
					new ExpectedCellValue(sheetName, 27, 12, 1194d),
					new ExpectedCellValue(sheetName, 28, 12, 10d),
					new ExpectedCellValue(sheetName, 29, 12, 2857d),
					new ExpectedCellValue(sheetName, 30, 12, 15d),
					new ExpectedCellValue(sheetName, 31, 12, 4228.24)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableColumnFields.xlsx")]
		public void PivotTableRefreshMultipleRowDataFieldsThreeRowsAndOneColumnSubtotalsOff()
		{
			var file = new FileInfo("PivotTableColumnFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["RowDataFields"];
					var pivotTable = worksheet.PivotTables["RowDataFieldsPivotTable3"];
					foreach (var field in pivotTable.Fields)
					{
						field.DisableDefaultSubtotal();
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(7, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(3, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(3, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					foreach (var field in pivotTable.Fields)
					{
						if (field.Items.Count > 0)
							this.CheckFieldItems(field);
					}
					package.SaveAs(newFile.File);
				}
				string sheetName = "RowDataFields";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 3, 15, "San Francisco"),
					new ExpectedCellValue(sheetName, 4, 15, "Car Rack"),
					new ExpectedCellValue(sheetName, 5, 15, "20100076"),
					new ExpectedCellValue(sheetName, 6, 15, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 7, 15, "Sum of Total"),
					new ExpectedCellValue(sheetName, 8, 15, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 9, 15, "20100085"),
					new ExpectedCellValue(sheetName, 10, 15, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 11, 15, "Sum of Total"),
					new ExpectedCellValue(sheetName, 12, 15, "Chicago"),
					new ExpectedCellValue(sheetName, 13, 15, "Car Rack"),
					new ExpectedCellValue(sheetName, 14, 15, "20100007"),
					new ExpectedCellValue(sheetName, 15, 15, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 16, 15, "Sum of Total"),
					new ExpectedCellValue(sheetName, 17, 15, "Headlamp"),
					new ExpectedCellValue(sheetName, 18, 15, "20100083"),
					new ExpectedCellValue(sheetName, 19, 15, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 20, 15, "Sum of Total"),
					new ExpectedCellValue(sheetName, 21, 15, "Nashville"),
					new ExpectedCellValue(sheetName, 22, 15, "Car Rack"),
					new ExpectedCellValue(sheetName, 23, 15, "20100017"),
					new ExpectedCellValue(sheetName, 24, 15, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 25, 15, "Sum of Total"),
					new ExpectedCellValue(sheetName, 26, 15, "20100090"),
					new ExpectedCellValue(sheetName, 27, 15, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 28, 15, "Sum of Total"),
					new ExpectedCellValue(sheetName, 29, 15, "Tent"),
					new ExpectedCellValue(sheetName, 30, 15, "20100070"),
					new ExpectedCellValue(sheetName, 31, 15, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 32, 15, "Sum of Total"),
					new ExpectedCellValue(sheetName, 33, 15, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 34, 15, "Total Sum of Total"),

					new ExpectedCellValue(sheetName, 2, 16, "January"),
					new ExpectedCellValue(sheetName, 6, 16, 1d),
					new ExpectedCellValue(sheetName, 7, 16, 415.75),
					new ExpectedCellValue(sheetName, 15, 16, 2d),
					new ExpectedCellValue(sheetName, 16, 16, 831.5),
					new ExpectedCellValue(sheetName, 27, 16, 2d),
					new ExpectedCellValue(sheetName, 28, 16, 831.5),
					new ExpectedCellValue(sheetName, 33, 16, 5d),
					new ExpectedCellValue(sheetName, 34, 16, 2078.75),

					new ExpectedCellValue(sheetName, 2, 17, "February"),
					new ExpectedCellValue(sheetName, 10, 17, 1d),
					new ExpectedCellValue(sheetName, 11, 17, 99d),
					new ExpectedCellValue(sheetName, 31, 17, 6d),
					new ExpectedCellValue(sheetName, 32, 17, 1194d),
					new ExpectedCellValue(sheetName, 33, 17, 7d),
					new ExpectedCellValue(sheetName, 34, 17, 1293d),

					new ExpectedCellValue(sheetName, 2, 18, "March"),
					new ExpectedCellValue(sheetName, 19, 18, 1d),
					new ExpectedCellValue(sheetName, 20, 18, 24.99),
					new ExpectedCellValue(sheetName, 24, 18, 2d),
					new ExpectedCellValue(sheetName, 25, 18, 831.5),
					new ExpectedCellValue(sheetName, 33, 18, 3d),
					new ExpectedCellValue(sheetName, 34, 18, 856.49),

					new ExpectedCellValue(sheetName, 2, 19, "Grand Total"),
					new ExpectedCellValue(sheetName, 6, 19, 1d),
					new ExpectedCellValue(sheetName, 7, 19, 415.75),
					new ExpectedCellValue(sheetName, 10, 19, 1d),
					new ExpectedCellValue(sheetName, 11, 19, 99d),
					new ExpectedCellValue(sheetName, 15, 19, 2d),
					new ExpectedCellValue(sheetName, 16, 19, 831.5),
					new ExpectedCellValue(sheetName, 19, 19, 1d),
					new ExpectedCellValue(sheetName, 20, 19, 24.99),
					new ExpectedCellValue(sheetName, 24, 19, 2d),
					new ExpectedCellValue(sheetName, 25, 19, 831.5),
					new ExpectedCellValue(sheetName, 27, 19, 2d),
					new ExpectedCellValue(sheetName, 28, 19, 831.5),
					new ExpectedCellValue(sheetName, 31, 19, 6d),
					new ExpectedCellValue(sheetName, 32, 19, 1194d),
					new ExpectedCellValue(sheetName, 33, 19, 15d),
					new ExpectedCellValue(sheetName, 34, 19, 4228.24)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableColumnFields.xlsx")]
		public void PivotTableRefreshMultipleRowDataFieldsThreeRowsAndOneColumnSubtotalsOn()
		{
			var file = new FileInfo("PivotTableColumnFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["RowDataFields"];
					var pivotTable = worksheet.PivotTables["RowDataFieldsPivotTable3"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(8, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					foreach (var field in pivotTable.Fields)
					{
						if (field.Items.Count > 0)
							this.CheckFieldItems(field);
					}
					package.SaveAs(newFile.File);
				}
				string sheetName = "RowDataFields";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 3, 15, "San Francisco"),
					new ExpectedCellValue(sheetName, 4, 15, "Car Rack"),
					new ExpectedCellValue(sheetName, 5, 15, "20100076"),
					new ExpectedCellValue(sheetName, 6, 15, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 7, 15, "Sum of Total"),
					new ExpectedCellValue(sheetName, 8, 15, "Car Rack Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 9, 15, "Car Rack Sum of Total"),
					new ExpectedCellValue(sheetName, 10, 15, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 11, 15, "20100085"),
					new ExpectedCellValue(sheetName, 12, 15, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 13, 15, "Sum of Total"),
					new ExpectedCellValue(sheetName, 14, 15, "Sleeping Bag Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 15, 15, "Sleeping Bag Sum of Total"),
					new ExpectedCellValue(sheetName, 16, 15, "San Francisco Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 17, 15, "San Francisco Sum of Total"),
					new ExpectedCellValue(sheetName, 18, 15, "Chicago"),
					new ExpectedCellValue(sheetName, 19, 15, "Car Rack"),
					new ExpectedCellValue(sheetName, 20, 15, "20100007"),
					new ExpectedCellValue(sheetName, 21, 15, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 22, 15, "Sum of Total"),
					new ExpectedCellValue(sheetName, 23, 15, "Car Rack Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 24, 15, "Car Rack Sum of Total"),
					new ExpectedCellValue(sheetName, 25, 15, "Headlamp"),
					new ExpectedCellValue(sheetName, 26, 15, "20100083"),
					new ExpectedCellValue(sheetName, 27, 15, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 28, 15, "Sum of Total"),
					new ExpectedCellValue(sheetName, 29, 15, "Headlamp Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 30, 15, "Headlamp Sum of Total"),
					new ExpectedCellValue(sheetName, 31, 15, "Chicago Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 32, 15, "Chicago Sum of Total"),
					new ExpectedCellValue(sheetName, 33, 15, "Nashville"),
					new ExpectedCellValue(sheetName, 34, 15, "Car Rack"),
					new ExpectedCellValue(sheetName, 35, 15, "20100017"),
					new ExpectedCellValue(sheetName, 36, 15, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 37, 15, "Sum of Total"),
					new ExpectedCellValue(sheetName, 38, 15, "20100090"),
					new ExpectedCellValue(sheetName, 39, 15, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 40, 15, "Sum of Total"),
					new ExpectedCellValue(sheetName, 41, 15, "Car Rack Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 42, 15, "Car Rack Sum of Total"),
					new ExpectedCellValue(sheetName, 43, 15, "Tent"),
					new ExpectedCellValue(sheetName, 44, 15, "20100070"),
					new ExpectedCellValue(sheetName, 45, 15, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 46, 15, "Sum of Total"),
					new ExpectedCellValue(sheetName, 47, 15, "Tent Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 48, 15, "Tent Sum of Total"),
					new ExpectedCellValue(sheetName, 49, 15, "Nashville Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 50, 15, "Nashville Sum of Total"),
					new ExpectedCellValue(sheetName, 51, 15, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 52, 15, "Total Sum of Total"),

					new ExpectedCellValue(sheetName, 2, 16, "January"),
					new ExpectedCellValue(sheetName, 6, 16, 1d),
					new ExpectedCellValue(sheetName, 7, 16, 415.75),
					new ExpectedCellValue(sheetName, 8, 16, 1d),
					new ExpectedCellValue(sheetName, 9, 16, 415.75),
					new ExpectedCellValue(sheetName, 16, 16, 1d),
					new ExpectedCellValue(sheetName, 17, 16, 415.75),
					new ExpectedCellValue(sheetName, 21, 16, 2d),
					new ExpectedCellValue(sheetName, 22, 16, 831.5),
					new ExpectedCellValue(sheetName, 23, 16, 2d),
					new ExpectedCellValue(sheetName, 24, 16, 831.5),
					new ExpectedCellValue(sheetName, 31, 16, 2d),
					new ExpectedCellValue(sheetName, 32, 16, 831.5),
					new ExpectedCellValue(sheetName, 39, 16, 2d),
					new ExpectedCellValue(sheetName, 40, 16, 831.5),
					new ExpectedCellValue(sheetName, 41, 16, 2d),
					new ExpectedCellValue(sheetName, 42, 16, 831.5),
					new ExpectedCellValue(sheetName, 49, 16, 2d),
					new ExpectedCellValue(sheetName, 50, 16, 831.5),
					new ExpectedCellValue(sheetName, 51, 16, 5d),
					new ExpectedCellValue(sheetName, 52, 16, 2078.75),

					new ExpectedCellValue(sheetName, 2, 17, "February"),
					new ExpectedCellValue(sheetName, 12, 17, 1d),
					new ExpectedCellValue(sheetName, 13, 17, 99d),
					new ExpectedCellValue(sheetName, 14, 17, 1d),
					new ExpectedCellValue(sheetName, 15, 17, 99d),
					new ExpectedCellValue(sheetName, 16, 17, 1d),
					new ExpectedCellValue(sheetName, 17, 17, 99d),
					new ExpectedCellValue(sheetName, 45, 17, 6d),
					new ExpectedCellValue(sheetName, 46, 17, 1194d),
					new ExpectedCellValue(sheetName, 47, 17, 6d),
					new ExpectedCellValue(sheetName, 48, 17, 1194d),
					new ExpectedCellValue(sheetName, 49, 17, 6d),
					new ExpectedCellValue(sheetName, 50, 17, 1194d),
					new ExpectedCellValue(sheetName, 51, 17, 7d),
					new ExpectedCellValue(sheetName, 52, 17, 1293d),

					new ExpectedCellValue(sheetName, 2, 18, "March"),
					new ExpectedCellValue(sheetName, 27, 18, 1d),
					new ExpectedCellValue(sheetName, 28, 18, 24.99),
					new ExpectedCellValue(sheetName, 29, 18, 1d),
					new ExpectedCellValue(sheetName, 30, 18, 24.99),
					new ExpectedCellValue(sheetName, 31, 18, 1d),
					new ExpectedCellValue(sheetName, 32, 18, 24.99),
					new ExpectedCellValue(sheetName, 36, 18, 2d),
					new ExpectedCellValue(sheetName, 37, 18, 831.5),
					new ExpectedCellValue(sheetName, 41, 18, 2d),
					new ExpectedCellValue(sheetName, 42, 18, 831.5),
					new ExpectedCellValue(sheetName, 49, 18, 2d),
					new ExpectedCellValue(sheetName, 50, 18, 831.5),
					new ExpectedCellValue(sheetName, 51, 18, 3d),
					new ExpectedCellValue(sheetName, 52, 18, 856.49),

					new ExpectedCellValue(sheetName, 2, 19, "Grand Total"),
					new ExpectedCellValue(sheetName, 6, 19, 1d),
					new ExpectedCellValue(sheetName, 7, 19, 415.75),
					new ExpectedCellValue(sheetName, 8, 19, 1d),
					new ExpectedCellValue(sheetName, 9, 19, 415.75),
					new ExpectedCellValue(sheetName, 12, 19, 1d),
					new ExpectedCellValue(sheetName, 13, 19, 99d),
					new ExpectedCellValue(sheetName, 14, 19, 1d),
					new ExpectedCellValue(sheetName, 15, 19, 99d),
					new ExpectedCellValue(sheetName, 16, 19, 2d),
					new ExpectedCellValue(sheetName, 17, 19, 514.75),
					new ExpectedCellValue(sheetName, 21, 19, 2d),
					new ExpectedCellValue(sheetName, 22, 19, 831.5),
					new ExpectedCellValue(sheetName, 23, 19, 2d),
					new ExpectedCellValue(sheetName, 24, 19, 831.5),
					new ExpectedCellValue(sheetName, 27, 19, 1d),
					new ExpectedCellValue(sheetName, 28, 19, 24.99),
					new ExpectedCellValue(sheetName, 29, 19, 1d),
					new ExpectedCellValue(sheetName, 30, 19, 24.99),
					new ExpectedCellValue(sheetName, 31, 19, 3d),
					new ExpectedCellValue(sheetName, 32, 19, 856.49),
					new ExpectedCellValue(sheetName, 36, 19, 2d),
					new ExpectedCellValue(sheetName, 37, 19, 831.5),
					new ExpectedCellValue(sheetName, 39, 19, 2d),
					new ExpectedCellValue(sheetName, 40, 19, 831.5),
					new ExpectedCellValue(sheetName, 41, 19, 4d),
					new ExpectedCellValue(sheetName, 42, 19, 1663),
					new ExpectedCellValue(sheetName, 45, 19, 6d),
					new ExpectedCellValue(sheetName, 46, 19, 1194d),
					new ExpectedCellValue(sheetName, 47, 19, 6d),
					new ExpectedCellValue(sheetName, 48, 19, 1194d),
					new ExpectedCellValue(sheetName, 49, 19, 10d),
					new ExpectedCellValue(sheetName, 50, 19, 2857d),
					new ExpectedCellValue(sheetName, 51, 19, 15d),
					new ExpectedCellValue(sheetName, 52, 19, 4228.24)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableColumnFields.xlsx")]
		public void PivotTableRefreshMultipleRowDataFieldsOneRowAndNoColumns()
		{
			var file = new FileInfo("PivotTableColumnFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["RowDataFields"];
					var pivotTable = worksheet.PivotTables["RowDataFieldsPivotTable4"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					foreach (var field in pivotTable.Fields)
					{
						if (field.Items.Count > 0)
							this.CheckFieldItems(field);
					}
					package.SaveAs(newFile.File);
				}
				string sheetName = "RowDataFields";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 38, 1, "January"),
					new ExpectedCellValue(sheetName, 39, 1, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 40, 1, "Sum of Total"),
					new ExpectedCellValue(sheetName, 41, 1, "February"),
					new ExpectedCellValue(sheetName, 42, 1, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 43, 1, "Sum of Total"),
					new ExpectedCellValue(sheetName, 44, 1, "March"),
					new ExpectedCellValue(sheetName, 45, 1, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 46, 1, "Sum of Total"),
					new ExpectedCellValue(sheetName, 47, 1, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 48, 1, "Total Sum of Total"),

					new ExpectedCellValue(sheetName, 39, 2, 5d),
					new ExpectedCellValue(sheetName, 40, 2, 2078.75),
					new ExpectedCellValue(sheetName, 42, 2, 7d),
					new ExpectedCellValue(sheetName, 43, 2, 1293d),
					new ExpectedCellValue(sheetName, 45, 2, 3d),
					new ExpectedCellValue(sheetName, 46, 2, 856.49),
					new ExpectedCellValue(sheetName, 47, 2, 15d),
					new ExpectedCellValue(sheetName, 48, 2, 4228.24)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableColumnFields.xlsx")]
		public void PivotTableRefreshMultipleRowDataFieldsTwoRowsAndNoColumnsSubtotalsOff()
		{
			var file = new FileInfo("PivotTableColumnFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["RowDataFields"];
					var pivotTable = worksheet.PivotTables["RowDataFieldsPivotTable5"];
					foreach (var field in pivotTable.Fields)
					{
						field.DisableDefaultSubtotal();
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(3, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(3, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					foreach (var field in pivotTable.Fields)
					{
						if (field.Items.Count > 0)
							this.CheckFieldItems(field);
					}
					package.SaveAs(newFile.File);
				}
				string sheetName = "RowDataFields";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 38, 5, "January"),
					new ExpectedCellValue(sheetName, 39, 5, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 40, 5, "San Francisco"),
					new ExpectedCellValue(sheetName, 41, 5, "Chicago"),
					new ExpectedCellValue(sheetName, 42, 5, "Nashville"),
					new ExpectedCellValue(sheetName, 43, 5, "Sum of Total"),
					new ExpectedCellValue(sheetName, 44, 5, "San Francisco"),
					new ExpectedCellValue(sheetName, 45, 5, "Chicago"),
					new ExpectedCellValue(sheetName, 46, 5, "Nashville"),
					new ExpectedCellValue(sheetName, 47, 5, "February"),
					new ExpectedCellValue(sheetName, 48, 5, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 49, 5, "San Francisco"),
					new ExpectedCellValue(sheetName, 50, 5, "Nashville"),
					new ExpectedCellValue(sheetName, 51, 5, "Sum of Total"),
					new ExpectedCellValue(sheetName, 52, 5, "San Francisco"),
					new ExpectedCellValue(sheetName, 53, 5, "Nashville"),
					new ExpectedCellValue(sheetName, 54, 5, "March"),
					new ExpectedCellValue(sheetName, 55, 5, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 56, 5, "Chicago"),
					new ExpectedCellValue(sheetName, 57, 5, "Nashville"),
					new ExpectedCellValue(sheetName, 58, 5, "Sum of Total"),
					new ExpectedCellValue(sheetName, 59, 5, "Chicago"),
					new ExpectedCellValue(sheetName, 60, 5, "Nashville"),
					new ExpectedCellValue(sheetName, 61, 5, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 62, 5, "Total Sum of Total"),

					new ExpectedCellValue(sheetName, 40, 6, 1d),
					new ExpectedCellValue(sheetName, 41, 6, 2d),
					new ExpectedCellValue(sheetName, 42, 6, 2d),
					new ExpectedCellValue(sheetName, 44, 6, 415.75),
					new ExpectedCellValue(sheetName, 45, 6, 831.5),
					new ExpectedCellValue(sheetName, 46, 6, 831.5),
					new ExpectedCellValue(sheetName, 49, 6, 1d),
					new ExpectedCellValue(sheetName, 50, 6, 6d),
					new ExpectedCellValue(sheetName, 52, 6, 99d),
					new ExpectedCellValue(sheetName, 53, 6, 1194d),
					new ExpectedCellValue(sheetName, 56, 6, 1d),
					new ExpectedCellValue(sheetName, 57, 6, 2d),
					new ExpectedCellValue(sheetName, 59, 6, 24.99),
					new ExpectedCellValue(sheetName, 60, 6, 831.5),
					new ExpectedCellValue(sheetName, 61, 6, 15d),
					new ExpectedCellValue(sheetName, 62, 6, 4228.24)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableColumnFields.xlsx")]
		public void PivotTableRefreshMultipleRowDataFieldsTwoRowsAndNoColumnsLastColumnDataField()
		{
			var file = new FileInfo("PivotTableColumnFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["RowDataFields"];
					var pivotTable = worksheet.PivotTables["RowDataFieldsPivotTable6"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					foreach (var field in pivotTable.Fields)
					{
						if (field.Items.Count > 0)
							this.CheckFieldItems(field);
					}
					package.SaveAs(newFile.File);
				}
				string sheetName = "RowDataFields";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 38, 8, "January"),
					new ExpectedCellValue(sheetName, 39, 8, "San Francisco"),
					new ExpectedCellValue(sheetName, 40, 8, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 41, 8, "Sum of Total"),
					new ExpectedCellValue(sheetName, 42, 8, "Chicago"),
					new ExpectedCellValue(sheetName, 43, 8, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 44, 8, "Sum of Total"),
					new ExpectedCellValue(sheetName, 45, 8, "Nashville"),
					new ExpectedCellValue(sheetName, 46, 8, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 47, 8, "Sum of Total"),
					new ExpectedCellValue(sheetName, 48, 8, "January Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 49, 8, "January Sum of Total"),
					new ExpectedCellValue(sheetName, 50, 8, "February"),
					new ExpectedCellValue(sheetName, 51, 8, "San Francisco"),
					new ExpectedCellValue(sheetName, 52, 8, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 53, 8, "Sum of Total"),
					new ExpectedCellValue(sheetName, 54, 8, "Nashville"),
					new ExpectedCellValue(sheetName, 55, 8, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 56, 8, "Sum of Total"),
					new ExpectedCellValue(sheetName, 57, 8, "February Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 58, 8, "February Sum of Total"),
					new ExpectedCellValue(sheetName, 59, 8, "March"),
					new ExpectedCellValue(sheetName, 60, 8, "Chicago"),
					new ExpectedCellValue(sheetName, 61, 8, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 62, 8, "Sum of Total"),
					new ExpectedCellValue(sheetName, 63, 8, "Nashville"),
					new ExpectedCellValue(sheetName, 64, 8, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 65, 8, "Sum of Total"),
					new ExpectedCellValue(sheetName, 66, 8, "March Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 67, 8, "March Sum of Total"),
					new ExpectedCellValue(sheetName, 68, 8, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 69, 8, "Total Sum of Total"),

					new ExpectedCellValue(sheetName, 40, 9, 1d),
					new ExpectedCellValue(sheetName, 41, 9, 415.75d),
					new ExpectedCellValue(sheetName, 43, 9, 2d),
					new ExpectedCellValue(sheetName, 44, 9, 831.5d),
					new ExpectedCellValue(sheetName, 46, 9, 2d),
					new ExpectedCellValue(sheetName, 47, 9, 831.5d),
					new ExpectedCellValue(sheetName, 48, 9, 5d),
					new ExpectedCellValue(sheetName, 49, 9, 2078.75),
					new ExpectedCellValue(sheetName, 52, 9, 1d),
					new ExpectedCellValue(sheetName, 53, 9, 99d),
					new ExpectedCellValue(sheetName, 55, 9, 6d),
					new ExpectedCellValue(sheetName, 56, 9, 1194d),
					new ExpectedCellValue(sheetName, 57, 9, 7d),
					new ExpectedCellValue(sheetName, 58, 9, 1293d),
					new ExpectedCellValue(sheetName, 61, 9, 1d),
					new ExpectedCellValue(sheetName, 62, 9, 24.99),
					new ExpectedCellValue(sheetName, 64, 9, 2d),
					new ExpectedCellValue(sheetName, 65, 9, 831.5),
					new ExpectedCellValue(sheetName, 66, 9, 3d),
					new ExpectedCellValue(sheetName, 67, 9, 856.49),
					new ExpectedCellValue(sheetName, 68, 9, 15),
					new ExpectedCellValue(sheetName, 69, 9, 4228.24)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableColumnFields.xlsx")]
		public void PivotTableRefreshMultipleRowDataFieldsThreeRowsAndNoColumnsLastColumnDataField()
		{
			var file = new FileInfo("PivotTableColumnFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["RowDataFields"];
					var pivotTable = worksheet.PivotTables["RowDataFieldsPivotTable7"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(8, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					foreach (var field in pivotTable.Fields)
					{
						if (field.Items.Count > 0)
							this.CheckFieldItems(field);
					}
					package.SaveAs(newFile.File);
				}
				string sheetName = "RowDataFields";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 38, 11, "January"),
					new ExpectedCellValue(sheetName, 39, 11, "San Francisco"),
					new ExpectedCellValue(sheetName, 40, 11, "20100076"),
					new ExpectedCellValue(sheetName, 41, 11, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 42, 11, "Sum of Total"),
					new ExpectedCellValue(sheetName, 43, 11, "San Francisco Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 44, 11, "San Francisco Sum of Total"),
					new ExpectedCellValue(sheetName, 45, 11, "Chicago"),
					new ExpectedCellValue(sheetName, 46, 11, "20100007"),
					new ExpectedCellValue(sheetName, 47, 11, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 48, 11, "Sum of Total"),
					new ExpectedCellValue(sheetName, 49, 11, "Chicago Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 50, 11, "Chicago Sum of Total"),
					new ExpectedCellValue(sheetName, 51, 11, "Nashville"),
					new ExpectedCellValue(sheetName, 52, 11, "20100090"),
					new ExpectedCellValue(sheetName, 53, 11, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 54, 11, "Sum of Total"),
					new ExpectedCellValue(sheetName, 55, 11, "Nashville Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 56, 11, "Nashville Sum of Total"),
					new ExpectedCellValue(sheetName, 57, 11, "January Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 58, 11, "January Sum of Total"),
					new ExpectedCellValue(sheetName, 59, 11, "February"),
					new ExpectedCellValue(sheetName, 60, 11, "San Francisco"),
					new ExpectedCellValue(sheetName, 61, 11, "20100085"),
					new ExpectedCellValue(sheetName, 62, 11, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 63, 11, "Sum of Total"),
					new ExpectedCellValue(sheetName, 64, 11, "San Francisco Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 65, 11, "San Francisco Sum of Total"),
					new ExpectedCellValue(sheetName, 66, 11, "Nashville"),
					new ExpectedCellValue(sheetName, 67, 11, "20100070"),
					new ExpectedCellValue(sheetName, 68, 11, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 69, 11, "Sum of Total"),
					new ExpectedCellValue(sheetName, 70, 11, "Nashville Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 71, 11, "Nashville Sum of Total"),
					new ExpectedCellValue(sheetName, 72, 11, "February Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 73, 11, "February Sum of Total"),
					new ExpectedCellValue(sheetName, 74, 11, "March"),
					new ExpectedCellValue(sheetName, 75, 11, "Chicago"),
					new ExpectedCellValue(sheetName, 76, 11, "20100083"),
					new ExpectedCellValue(sheetName, 77, 11, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 78, 11, "Sum of Total"),
					new ExpectedCellValue(sheetName, 79, 11, "Chicago Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 80, 11, "Chicago Sum of Total"),
					new ExpectedCellValue(sheetName, 81, 11, "Nashville"),
					new ExpectedCellValue(sheetName, 82, 11, "20100017"),
					new ExpectedCellValue(sheetName, 83, 11, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 84, 11, "Sum of Total"),
					new ExpectedCellValue(sheetName, 85, 11, "Nashville Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 86, 11, "Nashville Sum of Total"),
					new ExpectedCellValue(sheetName, 87, 11, "March Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 88, 11, "March Sum of Total"),
					new ExpectedCellValue(sheetName, 89, 11, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 90, 11, "Total Sum of Total"),

					new ExpectedCellValue(sheetName, 41, 12, 1d),
					new ExpectedCellValue(sheetName, 42, 12, 415.75d),
					new ExpectedCellValue(sheetName, 43, 12, 1d),
					new ExpectedCellValue(sheetName, 44, 12, 415.75d),
					new ExpectedCellValue(sheetName, 47, 12, 2d),
					new ExpectedCellValue(sheetName, 48, 12, 831.5d),
					new ExpectedCellValue(sheetName, 49, 12, 2d),
					new ExpectedCellValue(sheetName, 50, 12, 831.5d),
					new ExpectedCellValue(sheetName, 53, 12, 2d),
					new ExpectedCellValue(sheetName, 54, 12, 831.5d),
					new ExpectedCellValue(sheetName, 55, 12, 2d),
					new ExpectedCellValue(sheetName, 56, 12, 831.5d),
					new ExpectedCellValue(sheetName, 57, 12, 5d),
					new ExpectedCellValue(sheetName, 58, 12, 2078.75),
					new ExpectedCellValue(sheetName, 62, 12, 1d),
					new ExpectedCellValue(sheetName, 63, 12, 99d),
					new ExpectedCellValue(sheetName, 64, 12, 1d),
					new ExpectedCellValue(sheetName, 65, 12, 99d),
					new ExpectedCellValue(sheetName, 68, 12, 6d),
					new ExpectedCellValue(sheetName, 69, 12, 1194d),
					new ExpectedCellValue(sheetName, 70, 12, 6d),
					new ExpectedCellValue(sheetName, 71, 12, 1194d),
					new ExpectedCellValue(sheetName, 72, 12, 7d),
					new ExpectedCellValue(sheetName, 73, 12, 1293d),
					new ExpectedCellValue(sheetName, 77, 12, 1d),
					new ExpectedCellValue(sheetName, 78, 12, 24.99),
					new ExpectedCellValue(sheetName, 79, 12, 1d),
					new ExpectedCellValue(sheetName, 80, 12, 24.99),
					new ExpectedCellValue(sheetName, 83, 12, 2d),
					new ExpectedCellValue(sheetName, 84, 12, 831.5),
					new ExpectedCellValue(sheetName, 85, 12, 2d),
					new ExpectedCellValue(sheetName, 86, 12, 831.5),
					new ExpectedCellValue(sheetName, 87, 12, 3d),
					new ExpectedCellValue(sheetName, 88, 12, 856.49),
					new ExpectedCellValue(sheetName, 89, 12, 15),
					new ExpectedCellValue(sheetName, 90, 12, 4228.24)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableColumnFields.xlsx")]
		public void PivotTableRefreshMultipleRowDataFieldsTwoRowsAndNoColumnsSubtotalsOn()
		{
			var file = new FileInfo("PivotTableColumnFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["RowDataFields"];
					var pivotTable = worksheet.PivotTables["RowDataFieldsPivotTable5"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					foreach (var field in pivotTable.Fields)
					{
						if (field.Items.Count > 0)
							this.CheckFieldItems(field);
					}
					package.SaveAs(newFile.File);
				}
				string sheetName = "RowDataFields";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 38, 5, "January"),
					new ExpectedCellValue(sheetName, 39, 5, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 40, 5, "San Francisco"),
					new ExpectedCellValue(sheetName, 41, 5, "Chicago"),
					new ExpectedCellValue(sheetName, 42, 5, "Nashville"),
					new ExpectedCellValue(sheetName, 43, 5, "Sum of Total"),
					new ExpectedCellValue(sheetName, 44, 5, "San Francisco"),
					new ExpectedCellValue(sheetName, 45, 5, "Chicago"),
					new ExpectedCellValue(sheetName, 46, 5, "Nashville"),
					new ExpectedCellValue(sheetName, 47, 5, "January Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 48, 5, "January Sum of Total"),
					new ExpectedCellValue(sheetName, 49, 5, "February"),
					new ExpectedCellValue(sheetName, 50, 5, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 51, 5, "San Francisco"),
					new ExpectedCellValue(sheetName, 52, 5, "Nashville"),
					new ExpectedCellValue(sheetName, 53, 5, "Sum of Total"),
					new ExpectedCellValue(sheetName, 54, 5, "San Francisco"),
					new ExpectedCellValue(sheetName, 55, 5, "Nashville"),
					new ExpectedCellValue(sheetName, 56, 5, "February Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 57, 5, "February Sum of Total"),
					new ExpectedCellValue(sheetName, 58, 5, "March"),
					new ExpectedCellValue(sheetName, 59, 5, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 60, 5, "Chicago"),
					new ExpectedCellValue(sheetName, 61, 5, "Nashville"),
					new ExpectedCellValue(sheetName, 62, 5, "Sum of Total"),
					new ExpectedCellValue(sheetName, 63, 5, "Chicago"),
					new ExpectedCellValue(sheetName, 64, 5, "Nashville"),
					new ExpectedCellValue(sheetName, 65, 5, "March Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 66, 5, "March Sum of Total"),
					new ExpectedCellValue(sheetName, 67, 5, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 68, 5, "Total Sum of Total"),

					new ExpectedCellValue(sheetName, 40, 6, 1d),
					new ExpectedCellValue(sheetName, 41, 6, 2d),
					new ExpectedCellValue(sheetName, 42, 6, 2d),
					new ExpectedCellValue(sheetName, 44, 6, 415.75),
					new ExpectedCellValue(sheetName, 45, 6, 831.5),
					new ExpectedCellValue(sheetName, 46, 6, 831.5),
					new ExpectedCellValue(sheetName, 47, 6, 5d),
					new ExpectedCellValue(sheetName, 48, 6, 2078.75),
					new ExpectedCellValue(sheetName, 51, 6, 1d),
					new ExpectedCellValue(sheetName, 52, 6, 6d),
					new ExpectedCellValue(sheetName, 54, 6, 99d),
					new ExpectedCellValue(sheetName, 55, 6, 1194d),
					new ExpectedCellValue(sheetName, 56, 6, 7d),
					new ExpectedCellValue(sheetName, 57, 6, 1293d),
					new ExpectedCellValue(sheetName, 60, 6, 1d),
					new ExpectedCellValue(sheetName, 61, 6, 2d),
					new ExpectedCellValue(sheetName, 63, 6, 24.99),
					new ExpectedCellValue(sheetName, 64, 6, 831.5),
					new ExpectedCellValue(sheetName, 65, 6, 3d),
					new ExpectedCellValue(sheetName, 66, 6, 856.49),
					new ExpectedCellValue(sheetName, 67, 6, 15d),
					new ExpectedCellValue(sheetName, 68, 6, 4228.24)
				});
			}
		}
		#endregion

		#region Multiple Column Data Fields
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableColumnFields.xlsx")]
		public void PivotTableRefreshMultipleColumnDataFieldsAtLeafNode()
		{
			var file = new FileInfo("PivotTableColumnFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["ColumnDataFields"];
					var pivotTable = worksheet.PivotTables["ColumnDataFieldsPivotTable1"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					foreach (var field in pivotTable.Fields)
					{
						if (field.Items.Count > 0)
							this.CheckFieldItems(field);
					}
					package.SaveAs(newFile.File);
				}
				string sheetName = "ColumnDataFields";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 4, 1, "San Francisco"),
					new ExpectedCellValue(sheetName, 5, 1, "Chicago"),
					new ExpectedCellValue(sheetName, 6, 1, "Nashville"),
					new ExpectedCellValue(sheetName, 7, 1, "Grand Total"),

					new ExpectedCellValue(sheetName, 2, 2, "January"),
					new ExpectedCellValue(sheetName, 3, 2, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 4, 2, 1d),
					new ExpectedCellValue(sheetName, 5, 2, 2d),
					new ExpectedCellValue(sheetName, 6, 2, 2d),
					new ExpectedCellValue(sheetName, 7, 2, 5d),

					new ExpectedCellValue(sheetName, 3, 3, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 4, 3, 415.75),
					new ExpectedCellValue(sheetName, 5, 3, 415.75),
					new ExpectedCellValue(sheetName, 6, 3, 415.75),
					new ExpectedCellValue(sheetName, 7, 3, 1247.25),

					new ExpectedCellValue(sheetName, 2, 4, "February"),
					new ExpectedCellValue(sheetName, 3, 4, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 4, 4, 1d),
					new ExpectedCellValue(sheetName, 6, 4, 6d),
					new ExpectedCellValue(sheetName, 7, 4, 7d),

					new ExpectedCellValue(sheetName, 3, 5, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 4, 5, 99d),
					new ExpectedCellValue(sheetName, 6, 5, 199d),
					new ExpectedCellValue(sheetName, 7, 5, 298d),

					new ExpectedCellValue(sheetName, 2, 6, "March"),
					new ExpectedCellValue(sheetName, 3, 6, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 5, 6, 1d),
					new ExpectedCellValue(sheetName, 6, 6, 2d),
					new ExpectedCellValue(sheetName, 7, 6, 3d),

					new ExpectedCellValue(sheetName, 3, 7, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 5, 7, 24.99),
					new ExpectedCellValue(sheetName, 6, 7, 415.75),
					new ExpectedCellValue(sheetName, 7, 7, 440.74),

					new ExpectedCellValue(sheetName, 2, 8, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 4, 8, 2d),
					new ExpectedCellValue(sheetName, 5, 8, 3d),
					new ExpectedCellValue(sheetName, 6, 8, 10d),
					new ExpectedCellValue(sheetName, 7, 8, 15d),

					new ExpectedCellValue(sheetName, 2, 9, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 4, 9, 514.75),
					new ExpectedCellValue(sheetName, 5, 9, 440.74),
					new ExpectedCellValue(sheetName, 6, 9, 1030.5),
					new ExpectedCellValue(sheetName, 7, 9, 1985.99),
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableColumnFields.xlsx")]
		public void PivotTableRefreshMultipleColumnDataFieldsAsParentRowDepthTwo()
		{
			var file = new FileInfo("PivotTableColumnFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["ColumnDataFields"];
					var pivotTable = worksheet.PivotTables["ColumnDataFieldsPivotTable2"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					foreach (var field in pivotTable.Fields)
					{
						if (field.Items.Count > 0)
							this.CheckFieldItems(field);
					}
					package.SaveAs(newFile.File);
				}
				string sheetName = "ColumnDataFields";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 14, 1, "San Francisco"),
					new ExpectedCellValue(sheetName, 15, 1, "Chicago"),
					new ExpectedCellValue(sheetName, 16, 1, "Nashville"),
					new ExpectedCellValue(sheetName, 17, 1, "Grand Total"),

					new ExpectedCellValue(sheetName, 12, 2, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 13, 2, "January"),
					new ExpectedCellValue(sheetName, 14, 2, 1d),
					new ExpectedCellValue(sheetName, 15, 2, 2d),
					new ExpectedCellValue(sheetName, 16, 2, 2d),
					new ExpectedCellValue(sheetName, 17, 2, 5d),

					new ExpectedCellValue(sheetName, 13, 3, "February"),
					new ExpectedCellValue(sheetName, 14, 3, 1d),
					new ExpectedCellValue(sheetName, 16, 3, 6d),
					new ExpectedCellValue(sheetName, 17, 3, 7d),

					new ExpectedCellValue(sheetName, 13, 4, "March"),
					new ExpectedCellValue(sheetName, 15, 4, 1d),
					new ExpectedCellValue(sheetName, 16, 4, 2d),
					new ExpectedCellValue(sheetName, 17, 4, 3d),

					new ExpectedCellValue(sheetName, 12, 5, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 13, 5, "January"),
					new ExpectedCellValue(sheetName, 14, 5, 415.75),
					new ExpectedCellValue(sheetName, 15, 5, 415.75),
					new ExpectedCellValue(sheetName, 16, 5, 415.75),
					new ExpectedCellValue(sheetName, 17, 5, 1247.25),

					new ExpectedCellValue(sheetName, 13, 6, "February"),
					new ExpectedCellValue(sheetName, 14, 6, 99d),
					new ExpectedCellValue(sheetName, 16, 6, 199d),
					new ExpectedCellValue(sheetName, 17, 6, 298d),

					new ExpectedCellValue(sheetName, 13, 7, "March"),
					new ExpectedCellValue(sheetName, 15, 7, 24.99),
					new ExpectedCellValue(sheetName, 16, 7, 415.75),
					new ExpectedCellValue(sheetName, 17, 7, 440.74),

					new ExpectedCellValue(sheetName, 12, 8, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 14, 8, 2d),
					new ExpectedCellValue(sheetName, 15, 8, 3d),
					new ExpectedCellValue(sheetName, 16, 8, 10d),
					new ExpectedCellValue(sheetName, 17, 8, 15d),

					new ExpectedCellValue(sheetName, 12, 9, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 14, 9, 514.75),
					new ExpectedCellValue(sheetName, 15, 9, 440.74),
					new ExpectedCellValue(sheetName, 16, 9, 1030.5),
					new ExpectedCellValue(sheetName, 17, 9, 1985.99),
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableColumnFields.xlsx")]
		public void PivotTableRefreshMultipleColumnDataFieldsAsParentNodeColumnDepthThreeSubtotalsOn()
		{
			var file = new FileInfo("PivotTableColumnFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["ColumnDataFields"];
					var pivotTable = worksheet.PivotTables["ColumnDataFieldsPivotTable9"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					foreach (var field in pivotTable.Fields)
					{
						if (field.Items.Count > 0)
							this.CheckFieldItems(field);
					}
					package.SaveAs(newFile.File);
				}
				string sheetName = "ColumnDataFields";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 111, 1, "Car Rack"),
					new ExpectedCellValue(sheetName, 112, 1, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 113, 1, "Headlamp"),
					new ExpectedCellValue(sheetName, 114, 1, "Tent"),
					new ExpectedCellValue(sheetName, 115, 1, "Grand Total"),

					new ExpectedCellValue(sheetName, 108, 2, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 109, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 110, 2, "January"),
					new ExpectedCellValue(sheetName, 111, 2, 1d),
					new ExpectedCellValue(sheetName, 115, 2, 1d),

					new ExpectedCellValue(sheetName, 110, 3, "February"),
					new ExpectedCellValue(sheetName, 112, 3, 1d),
					new ExpectedCellValue(sheetName, 115, 3, 1d),

					new ExpectedCellValue(sheetName, 109, 4, "San Francisco Total"),
					new ExpectedCellValue(sheetName, 111, 4, 1d),
					new ExpectedCellValue(sheetName, 112, 4, 1d),
					new ExpectedCellValue(sheetName, 115, 4, 2d),

					new ExpectedCellValue(sheetName, 109, 5, "Chicago"),
					new ExpectedCellValue(sheetName, 110, 5, "January"),
					new ExpectedCellValue(sheetName, 111, 5, 2d),
					new ExpectedCellValue(sheetName, 115, 5, 2d),

					new ExpectedCellValue(sheetName, 110, 6, "March"),
					new ExpectedCellValue(sheetName, 113, 6, 1d),
					new ExpectedCellValue(sheetName, 115, 6, 1d),

					new ExpectedCellValue(sheetName, 109, 7, "Chicago Total"),
					new ExpectedCellValue(sheetName, 111, 7, 2d),
					new ExpectedCellValue(sheetName, 113, 7, 1d),
					new ExpectedCellValue(sheetName, 115, 7, 3d),

					new ExpectedCellValue(sheetName, 109, 8, "Nashville"),
					new ExpectedCellValue(sheetName, 110, 8, "January"),
					new ExpectedCellValue(sheetName, 111, 8, 2d),
					new ExpectedCellValue(sheetName, 115, 8, 2d),

					new ExpectedCellValue(sheetName, 110, 9, "February"),
					new ExpectedCellValue(sheetName, 114, 9, 6d),
					new ExpectedCellValue(sheetName, 115, 9, 6d),

					new ExpectedCellValue(sheetName, 110, 10, "March"),
					new ExpectedCellValue(sheetName, 111, 10, 2d),
					new ExpectedCellValue(sheetName, 115, 10, 2d),

					new ExpectedCellValue(sheetName, 109, 11, "Nashville Total"),
					new ExpectedCellValue(sheetName, 111, 11, 4d),
					new ExpectedCellValue(sheetName, 114, 11, 6d),
					new ExpectedCellValue(sheetName, 115, 11, 10d),

					new ExpectedCellValue(sheetName, 108, 12, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 109, 12, "San Francisco"),
					new ExpectedCellValue(sheetName, 110, 12, "January"),
					new ExpectedCellValue(sheetName, 111, 12, 415.75),
					new ExpectedCellValue(sheetName, 115, 12, 415.75),

					new ExpectedCellValue(sheetName, 110, 13, "February"),
					new ExpectedCellValue(sheetName, 112, 13, 99d),
					new ExpectedCellValue(sheetName, 115, 13, 99d),

					new ExpectedCellValue(sheetName, 109, 14, "San Francisco Total"),
					new ExpectedCellValue(sheetName, 111, 14, 1d),
					new ExpectedCellValue(sheetName, 112, 14, 1d),
					new ExpectedCellValue(sheetName, 115, 14, 2d),

					new ExpectedCellValue(sheetName, 109, 15, "Chicago"),
					new ExpectedCellValue(sheetName, 110, 15, "January"),
					new ExpectedCellValue(sheetName, 111, 15, 415.75),
					new ExpectedCellValue(sheetName, 115, 15, 415.75),

					new ExpectedCellValue(sheetName, 110, 16, "March"),
					new ExpectedCellValue(sheetName, 113, 16, 24.99),
					new ExpectedCellValue(sheetName, 115, 16, 24.99),

					new ExpectedCellValue(sheetName, 109, 17, "Chicago Total"),
					new ExpectedCellValue(sheetName, 111, 17, 2d),
					new ExpectedCellValue(sheetName, 113, 17, 1d),
					new ExpectedCellValue(sheetName, 115, 17, 3d),

					new ExpectedCellValue(sheetName, 109, 18, "Nashville"),
					new ExpectedCellValue(sheetName, 110, 18, "January"),
					new ExpectedCellValue(sheetName, 111, 18, 415.75),
					new ExpectedCellValue(sheetName, 115, 18, 415.75),

					new ExpectedCellValue(sheetName, 110, 19, "February"),
					new ExpectedCellValue(sheetName, 114, 19, 199d),
					new ExpectedCellValue(sheetName, 115, 19, 199d),

					new ExpectedCellValue(sheetName, 110, 20, "March"),
					new ExpectedCellValue(sheetName, 111, 20, 415.75),
					new ExpectedCellValue(sheetName, 115, 20, 415.75),

					new ExpectedCellValue(sheetName, 109, 21, "Nashville Total"),
					new ExpectedCellValue(sheetName, 111, 21, 4d),
					new ExpectedCellValue(sheetName, 114, 21, 6d),
					new ExpectedCellValue(sheetName, 115, 21, 10d),

					new ExpectedCellValue(sheetName, 108, 22, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 111, 22, 7d),
					new ExpectedCellValue(sheetName, 112, 22, 1d),
					new ExpectedCellValue(sheetName, 113, 22, 1d),
					new ExpectedCellValue(sheetName, 114, 22, 6d),
					new ExpectedCellValue(sheetName, 115, 22, 15d),

					new ExpectedCellValue(sheetName, 108, 23, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 111, 23, 1663d),
					new ExpectedCellValue(sheetName, 112, 23, 99d),
					new ExpectedCellValue(sheetName, 113, 23, 24.99),
					new ExpectedCellValue(sheetName, 114, 23, 199d),
					new ExpectedCellValue(sheetName, 115, 23, 1985.99),
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableColumnFields.xlsx")]
		public void PivotTableRefreshMultipleColumnDataFieldsAsParentNodeColumnDepthThreeSubtotalsOff()
		{
			var file = new FileInfo("PivotTableColumnFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["ColumnDataFields"];
					var pivotTable = worksheet.PivotTables["ColumnDataFieldsPivotTable10"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(3, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(3, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					foreach (var field in pivotTable.Fields)
					{
						if (field.Items.Count > 0)
							this.CheckFieldItems(field);
					}
					package.SaveAs(newFile.File);
				}
				string sheetName = "ColumnDataFields";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 123, 1, "Car Rack"),
					new ExpectedCellValue(sheetName, 124, 1, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 125, 1, "Headlamp"),
					new ExpectedCellValue(sheetName, 126, 1, "Tent"),
					new ExpectedCellValue(sheetName, 127, 1, "Grand Total"),

					new ExpectedCellValue(sheetName, 120, 2, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 121, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 122, 2, "January"),
					new ExpectedCellValue(sheetName, 123, 2, 1d),
					new ExpectedCellValue(sheetName, 127, 2, 1d),

					new ExpectedCellValue(sheetName, 122, 3, "February"),
					new ExpectedCellValue(sheetName, 124, 3, 1d),
					new ExpectedCellValue(sheetName, 127, 3, 1d),

					new ExpectedCellValue(sheetName, 121, 4, "Chicago"),
					new ExpectedCellValue(sheetName, 122, 4, "January"),
					new ExpectedCellValue(sheetName, 123, 4, 2d),
					new ExpectedCellValue(sheetName, 127, 4, 2d),

					new ExpectedCellValue(sheetName, 122, 5, "March"),
					new ExpectedCellValue(sheetName, 125, 5, 1d),
					new ExpectedCellValue(sheetName, 127, 5, 1d),

					new ExpectedCellValue(sheetName, 121, 6, "Nashville"),
					new ExpectedCellValue(sheetName, 122, 6, "January"),
					new ExpectedCellValue(sheetName, 123, 6, 2d),
					new ExpectedCellValue(sheetName, 127, 6, 2d),

					new ExpectedCellValue(sheetName, 122, 7, "February"),
					new ExpectedCellValue(sheetName, 126, 7, 6d),
					new ExpectedCellValue(sheetName, 127, 7, 6d),

					new ExpectedCellValue(sheetName, 122, 8, "March"),
					new ExpectedCellValue(sheetName, 123, 8, 2d),
					new ExpectedCellValue(sheetName, 127, 8, 2d),

					new ExpectedCellValue(sheetName, 120, 9, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 121, 9, "San Francisco"),
					new ExpectedCellValue(sheetName, 122, 9, "January"),
					new ExpectedCellValue(sheetName, 123, 9, 415.75),
					new ExpectedCellValue(sheetName, 127, 9, 415.75),

					new ExpectedCellValue(sheetName, 122, 10, "February"),
					new ExpectedCellValue(sheetName, 124, 10, 99d),
					new ExpectedCellValue(sheetName, 127, 10, 99d),

					new ExpectedCellValue(sheetName, 121, 11, "Chicago"),
					new ExpectedCellValue(sheetName, 122, 11, "January"),
					new ExpectedCellValue(sheetName, 123, 11, 415.75),
					new ExpectedCellValue(sheetName, 127, 11, 415.75),

					new ExpectedCellValue(sheetName, 122, 12, "March"),
					new ExpectedCellValue(sheetName, 125, 12, 24.99),
					new ExpectedCellValue(sheetName, 127, 12, 24.99),
				
					new ExpectedCellValue(sheetName, 121, 13, "Nashville"),
					new ExpectedCellValue(sheetName, 122, 13, "January"),
					new ExpectedCellValue(sheetName, 123, 13, 415.75),
					new ExpectedCellValue(sheetName, 127, 13, 415.75),

					new ExpectedCellValue(sheetName, 122, 14, "February"),
					new ExpectedCellValue(sheetName, 126, 14, 199d),
					new ExpectedCellValue(sheetName, 127, 14, 199d),

					new ExpectedCellValue(sheetName, 122, 15, "March"),
					new ExpectedCellValue(sheetName, 123, 15, 415.75),
					new ExpectedCellValue(sheetName, 127, 15, 415.75),

					new ExpectedCellValue(sheetName, 120, 16, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 123, 16, 7d),
					new ExpectedCellValue(sheetName, 124, 16, 1d),
					new ExpectedCellValue(sheetName, 125, 16, 1d),
					new ExpectedCellValue(sheetName, 126, 16, 6d),
					new ExpectedCellValue(sheetName, 127, 16, 15d),

					new ExpectedCellValue(sheetName, 120, 17, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 123, 17, 1663d),
					new ExpectedCellValue(sheetName, 124, 17, 99d),
					new ExpectedCellValue(sheetName, 125, 17, 24.99),
					new ExpectedCellValue(sheetName, 126, 17, 199d),
					new ExpectedCellValue(sheetName, 127, 17, 1985.99),
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableColumnFields.xlsx")]
		public void PivotTableRefreshMultipleColumnDataFieldsAsInnerChildSubtotalsOn()
		{
			var file = new FileInfo("PivotTableColumnFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["ColumnDataFields"];
					var pivotTable = worksheet.PivotTables["ColumnDataFieldsPivotTable3"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					foreach (var field in pivotTable.Fields)
					{
						if (field.Items.Count > 0)
							this.CheckFieldItems(field);
					}
					package.SaveAs(newFile.File);
				}
				string sheetName = "ColumnDataFields";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 25, 1, "San Francisco"),
					new ExpectedCellValue(sheetName, 26, 1, "Chicago"),
					new ExpectedCellValue(sheetName, 27, 1, "Nashville"),
					new ExpectedCellValue(sheetName, 28, 1, "Grand Total"),

					new ExpectedCellValue(sheetName, 22, 2, "January"),
					new ExpectedCellValue(sheetName, 23, 2, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 24, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 25, 2, 1d),
					new ExpectedCellValue(sheetName, 26, 2, 2d),
					new ExpectedCellValue(sheetName, 27, 2, 2d),
					new ExpectedCellValue(sheetName, 28, 2, 5d),

					new ExpectedCellValue(sheetName, 23, 3, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 24, 3, "Car Rack"),
					new ExpectedCellValue(sheetName, 25, 3, 415.75),
					new ExpectedCellValue(sheetName, 26, 3, 415.75),
					new ExpectedCellValue(sheetName, 27, 3, 415.75),
					new ExpectedCellValue(sheetName, 28, 3, 1247.25),

					new ExpectedCellValue(sheetName, 22, 4, "January Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 25, 4, 1d),
					new ExpectedCellValue(sheetName, 26, 4, 2d),
					new ExpectedCellValue(sheetName, 27, 4, 2d),
					new ExpectedCellValue(sheetName, 28, 4, 5d),

					new ExpectedCellValue(sheetName, 22, 5, "January Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 25, 5, 415.75),
					new ExpectedCellValue(sheetName, 26, 5, 415.75),
					new ExpectedCellValue(sheetName, 27, 5, 415.75),
					new ExpectedCellValue(sheetName, 28, 5, 1247.25),

					new ExpectedCellValue(sheetName, 22, 6, "February"),
					new ExpectedCellValue(sheetName, 23, 6, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 24, 6, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 25, 6, 1d),
					new ExpectedCellValue(sheetName, 28, 6, 1d),

					new ExpectedCellValue(sheetName, 24, 7, "Tent"),
					new ExpectedCellValue(sheetName, 27, 7, 6d),
					new ExpectedCellValue(sheetName, 28, 7, 6d),

					new ExpectedCellValue(sheetName, 23, 8, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 24, 8, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 25, 8, 99d),
					new ExpectedCellValue(sheetName, 28, 8, 99d),

					new ExpectedCellValue(sheetName, 24, 9, "Tent"),
					new ExpectedCellValue(sheetName, 27, 9, 199d),
					new ExpectedCellValue(sheetName, 28, 9, 199d),

					new ExpectedCellValue(sheetName, 22, 10, "February Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 25, 10, 1d),
					new ExpectedCellValue(sheetName, 27, 10, 6d),
					new ExpectedCellValue(sheetName, 28, 10, 7d),

					new ExpectedCellValue(sheetName, 22, 11, "February Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 25, 11, 99d),
					new ExpectedCellValue(sheetName, 27, 11, 199d),
					new ExpectedCellValue(sheetName, 28, 11, 298d),

					new ExpectedCellValue(sheetName, 22, 12, "March"),
					new ExpectedCellValue(sheetName, 23, 12, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 24, 12, "Car Rack"),
					new ExpectedCellValue(sheetName, 27, 12, 2d),
					new ExpectedCellValue(sheetName, 28, 12, 2d),

					new ExpectedCellValue(sheetName, 24, 13, "Headlamp"),
					new ExpectedCellValue(sheetName, 26, 13, 1d),
					new ExpectedCellValue(sheetName, 28, 13, 1d),

					new ExpectedCellValue(sheetName, 23, 14, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 24, 14, "Car Rack"),
					new ExpectedCellValue(sheetName, 27, 14, 415.75),
					new ExpectedCellValue(sheetName, 28, 14, 415.75),

					new ExpectedCellValue(sheetName, 24, 15, "Headlamp"),
					new ExpectedCellValue(sheetName, 26, 15, 24.99),
					new ExpectedCellValue(sheetName, 28, 15, 24.99),

					new ExpectedCellValue(sheetName, 22, 16, "March Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 26, 16, 1d),
					new ExpectedCellValue(sheetName, 27, 16, 2d),
					new ExpectedCellValue(sheetName, 28, 16, 3d),

					new ExpectedCellValue(sheetName, 22, 17, "March Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 26, 17, 24.99),
					new ExpectedCellValue(sheetName, 27, 17, 415.75),
					new ExpectedCellValue(sheetName, 28, 17, 440.74),

					new ExpectedCellValue(sheetName, 22, 18, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 25, 18, 2d),
					new ExpectedCellValue(sheetName, 26, 18, 3d),
					new ExpectedCellValue(sheetName, 27, 18, 10d),
					new ExpectedCellValue(sheetName, 28, 18, 15d),

					new ExpectedCellValue(sheetName, 22, 19, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 25, 19, 514.75),
					new ExpectedCellValue(sheetName, 26, 19, 440.74),
					new ExpectedCellValue(sheetName, 27, 19, 1030.5),
					new ExpectedCellValue(sheetName, 28, 19, 1985.99)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableColumnFields.xlsx")]
		public void PivotTableRefreshMultipleColumnDataFieldsAsInnerChildSubtotalsOff()
		{
			var file = new FileInfo("PivotTableColumnFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["ColumnDataFields"];
					var pivotTable = worksheet.PivotTables["ColumnDataFieldsPivotTable4"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(0, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(3, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(3, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					foreach (var field in pivotTable.Fields)
					{
						if (field.Items.Count > 0)
							this.CheckFieldItems(field);
					}
					package.SaveAs(newFile.File);
				}
				string sheetName = "ColumnDataFields";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 36, 1, "San Francisco"),
					new ExpectedCellValue(sheetName, 37, 1, "Chicago"),
					new ExpectedCellValue(sheetName, 38, 1, "Nashville"),
					new ExpectedCellValue(sheetName, 39, 1, "Grand Total"),

					new ExpectedCellValue(sheetName, 33, 2, "January"),
					new ExpectedCellValue(sheetName, 34, 2, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 35, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 36, 2, 1d),
					new ExpectedCellValue(sheetName, 37, 2, 2d),
					new ExpectedCellValue(sheetName, 38, 2, 2d),
					new ExpectedCellValue(sheetName, 39, 2, 5d),

					new ExpectedCellValue(sheetName, 34, 3, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 35, 3, "Car Rack"),
					new ExpectedCellValue(sheetName, 36, 3, 415.75),
					new ExpectedCellValue(sheetName, 37, 3, 415.75),
					new ExpectedCellValue(sheetName, 38, 3, 415.75),
					new ExpectedCellValue(sheetName, 39, 3, 1247.25),

					new ExpectedCellValue(sheetName, 33, 4, "February"),
					new ExpectedCellValue(sheetName, 34, 4, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 35, 4, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 36, 4, 1d),
					new ExpectedCellValue(sheetName, 39, 4, 1d),

					new ExpectedCellValue(sheetName, 35, 5, "Tent"),
					new ExpectedCellValue(sheetName, 38, 5, 6d),
					new ExpectedCellValue(sheetName, 39, 5, 6d),

					new ExpectedCellValue(sheetName, 34, 6, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 35, 6, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 36, 6, 99d),
					new ExpectedCellValue(sheetName, 39, 6, 99d),

					new ExpectedCellValue(sheetName, 35, 7, "Tent"),
					new ExpectedCellValue(sheetName, 38, 7, 199d),
					new ExpectedCellValue(sheetName, 39, 7, 199d),

					new ExpectedCellValue(sheetName, 33, 8, "March"),
					new ExpectedCellValue(sheetName, 34, 8, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 35, 8, "Car Rack"),
					new ExpectedCellValue(sheetName, 38, 8, 2d),
					new ExpectedCellValue(sheetName, 39, 8, 2d),

					new ExpectedCellValue(sheetName, 35, 9, "Headlamp"),
					new ExpectedCellValue(sheetName, 37, 9, 1d),
					new ExpectedCellValue(sheetName, 39, 9, 1d),

					new ExpectedCellValue(sheetName, 34, 10, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 35, 10, "Car Rack"),
					new ExpectedCellValue(sheetName, 38, 10, 415.75),
					new ExpectedCellValue(sheetName, 39, 10, 415.75),

					new ExpectedCellValue(sheetName, 35, 11, "Headlamp"),
					new ExpectedCellValue(sheetName, 37, 11, 24.99),
					new ExpectedCellValue(sheetName, 39, 11, 24.99),

					new ExpectedCellValue(sheetName, 33, 12, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 36, 12, 2d),
					new ExpectedCellValue(sheetName, 37, 12, 3d),
					new ExpectedCellValue(sheetName, 38, 12, 10d),
					new ExpectedCellValue(sheetName, 39, 12, 15d),

					new ExpectedCellValue(sheetName, 33, 13, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 36, 13, 514.75),
					new ExpectedCellValue(sheetName, 37, 13, 440.74),
					new ExpectedCellValue(sheetName, 38, 13, 1030.5),
					new ExpectedCellValue(sheetName, 39, 13, 1985.99)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableColumnFields.xlsx")]
		public void PivotTableRefreshMultipleColumnDataFieldsAsFirstInnerChildSubtotalsOn()
		{
			var file = new FileInfo("PivotTableColumnFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["ColumnDataFields"];
					var pivotTable = worksheet.PivotTables["ColumnDataFieldsPivotTable5"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(8, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					foreach (var field in pivotTable.Fields)
					{
						if (field.Items.Count > 0)
							this.CheckFieldItems(field);
					}
					package.SaveAs(newFile.File);
				}
				string sheetName = "ColumnDataFields";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 48, 1, 20100076),
					new ExpectedCellValue(sheetName, 49, 1, 20100085),
					new ExpectedCellValue(sheetName, 50, 1, 20100083),
					new ExpectedCellValue(sheetName, 51, 1, 20100007),
					new ExpectedCellValue(sheetName, 52, 1, 20100070),
					new ExpectedCellValue(sheetName, 53, 1, 20100017),
					new ExpectedCellValue(sheetName, 54, 1, 20100090),
					new ExpectedCellValue(sheetName, 55, 1, "Grand Total"),

					new ExpectedCellValue(sheetName, 44, 2, "January"),
					new ExpectedCellValue(sheetName, 45, 2, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 46, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 47, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 48, 2, 1d),
					new ExpectedCellValue(sheetName, 55, 2, 1d),

					new ExpectedCellValue(sheetName, 46, 3, "San Francisco Total"),
					new ExpectedCellValue(sheetName, 48, 3, 1d),
					new ExpectedCellValue(sheetName, 55, 3, 1d),

					new ExpectedCellValue(sheetName, 46, 4, "Chicago"),
					new ExpectedCellValue(sheetName, 47, 4, "Car Rack"),
					new ExpectedCellValue(sheetName, 51, 4, 2d),
					new ExpectedCellValue(sheetName, 55, 4, 2d),

					new ExpectedCellValue(sheetName, 46, 5, "Chicago Total"),
					new ExpectedCellValue(sheetName, 51, 5, 2d),
					new ExpectedCellValue(sheetName, 55, 5, 2d),

					new ExpectedCellValue(sheetName, 46, 6, "Nashville"),
					new ExpectedCellValue(sheetName, 47, 6, "Car Rack"),
					new ExpectedCellValue(sheetName, 54, 6, 2d),
					new ExpectedCellValue(sheetName, 55, 6, 2d),

					new ExpectedCellValue(sheetName, 46, 7, "Nashville Total"),
					new ExpectedCellValue(sheetName, 54, 7, 2d),
					new ExpectedCellValue(sheetName, 55, 7, 2d),

					new ExpectedCellValue(sheetName, 46, 8, "San Francisco"),
					new ExpectedCellValue(sheetName, 47, 8, "Car Rack"),
					new ExpectedCellValue(sheetName, 48, 8, 415.75),
					new ExpectedCellValue(sheetName, 55, 8, 415.75),

					new ExpectedCellValue(sheetName, 46, 9, "San Francisco Total"),
					new ExpectedCellValue(sheetName, 48, 9, 1d),
					new ExpectedCellValue(sheetName, 55, 9, 1d),

					new ExpectedCellValue(sheetName, 46, 10, "Chicago"),
					new ExpectedCellValue(sheetName, 47, 10, "Car Rack"),
					new ExpectedCellValue(sheetName, 51, 10, 415.75),
					new ExpectedCellValue(sheetName, 55, 10, 415.75),

					new ExpectedCellValue(sheetName, 46, 11, "Chicago Total"),
					new ExpectedCellValue(sheetName, 51, 11, 2d),
					new ExpectedCellValue(sheetName, 55, 11, 2d),

					new ExpectedCellValue(sheetName, 46, 12, "Nashville"),
					new ExpectedCellValue(sheetName, 47, 12, "Car Rack"),
					new ExpectedCellValue(sheetName, 54, 12, 415.75),
					new ExpectedCellValue(sheetName, 55, 12, 415.75),

					new ExpectedCellValue(sheetName, 46, 13, "Nashville Total"),
					new ExpectedCellValue(sheetName, 54, 13, 2d),
					new ExpectedCellValue(sheetName, 55, 13, 2d),

					new ExpectedCellValue(sheetName, 44, 14, "January Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 48, 14, 1d),
					new ExpectedCellValue(sheetName, 51, 14, 2d),
					new ExpectedCellValue(sheetName, 54, 14, 2d),
					new ExpectedCellValue(sheetName, 55, 14, 5d),

					new ExpectedCellValue(sheetName, 44, 15, "January Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 48, 15, 415.75),
					new ExpectedCellValue(sheetName, 51, 15, 415.75),
					new ExpectedCellValue(sheetName, 54, 15, 415.75),
					new ExpectedCellValue(sheetName, 55, 15, 1247.25),

					new ExpectedCellValue(sheetName, 44, 16, "February"),
					new ExpectedCellValue(sheetName, 45, 16, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 46, 16, "San Francisco"),
					new ExpectedCellValue(sheetName, 47, 16, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 49, 16, 1d),
					new ExpectedCellValue(sheetName, 55, 16, 1d),

					new ExpectedCellValue(sheetName, 46, 17, "San Francisco Total"),
					new ExpectedCellValue(sheetName, 49, 17, 1d),
					new ExpectedCellValue(sheetName, 55, 17, 1d),

					new ExpectedCellValue(sheetName, 46, 18, "Nashville"),
					new ExpectedCellValue(sheetName, 47, 18, "Tent"),
					new ExpectedCellValue(sheetName, 52, 18, 6d),
					new ExpectedCellValue(sheetName, 55, 18, 6d),

					new ExpectedCellValue(sheetName, 46, 19, "Nashville Total"),
					new ExpectedCellValue(sheetName, 52, 19, 6d),
					new ExpectedCellValue(sheetName, 55, 19, 6d),

					new ExpectedCellValue(sheetName, 46, 20, "San Francisco"),
					new ExpectedCellValue(sheetName, 47, 20, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 49, 20, 99d),
					new ExpectedCellValue(sheetName, 55, 20, 99d),

					new ExpectedCellValue(sheetName, 46, 21, "San Francisco Total"),
					new ExpectedCellValue(sheetName, 49, 21, 1d),
					new ExpectedCellValue(sheetName, 55, 21, 1d),

					new ExpectedCellValue(sheetName, 46, 22, "Nashville"),
					new ExpectedCellValue(sheetName, 47, 22, "Tent"),
					new ExpectedCellValue(sheetName, 52, 22, 199d),
					new ExpectedCellValue(sheetName, 55, 22, 199d),

					new ExpectedCellValue(sheetName, 46, 23, "Nashville Total"),
					new ExpectedCellValue(sheetName, 52, 23, 6d),
					new ExpectedCellValue(sheetName, 55, 23, 6d),

					new ExpectedCellValue(sheetName, 44, 24, "February Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 49, 24, 1d),
					new ExpectedCellValue(sheetName, 52, 24, 6d),
					new ExpectedCellValue(sheetName, 55, 24, 7d),

					new ExpectedCellValue(sheetName, 44, 25, "February Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 49, 25, 99d),
					new ExpectedCellValue(sheetName, 52, 25, 199d),
					new ExpectedCellValue(sheetName, 55, 25, 298d),

					new ExpectedCellValue(sheetName, 44, 26, "March"),
					new ExpectedCellValue(sheetName, 45, 26, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 46, 26, "Chicago"),
					new ExpectedCellValue(sheetName, 47, 26, "Headlamp"),
					new ExpectedCellValue(sheetName, 50, 26, 1d),
					new ExpectedCellValue(sheetName, 55, 26, 1d),

					new ExpectedCellValue(sheetName, 46, 27, "Chicago Total"),
					new ExpectedCellValue(sheetName, 50, 27, 1d),
					new ExpectedCellValue(sheetName, 55, 27, 1d),

					new ExpectedCellValue(sheetName, 46, 28, "Nashville"),
					new ExpectedCellValue(sheetName, 47, 28, "Car Rack"),
					new ExpectedCellValue(sheetName, 53, 28, 2d),
					new ExpectedCellValue(sheetName, 55, 28, 2d),

					new ExpectedCellValue(sheetName, 46, 29, "Nashville Total"),
					new ExpectedCellValue(sheetName, 53, 29, 2d),
					new ExpectedCellValue(sheetName, 55, 29, 2d),

					new ExpectedCellValue(sheetName, 46, 30, "Chicago"),
					new ExpectedCellValue(sheetName, 47, 30, "Headlamp"),
					new ExpectedCellValue(sheetName, 50, 30, 24.99),
					new ExpectedCellValue(sheetName, 55, 30, 24.99),

					new ExpectedCellValue(sheetName, 46, 31, "Chicago Total"),
					new ExpectedCellValue(sheetName, 50, 31, 1d),
					new ExpectedCellValue(sheetName, 55, 31, 1d),

					new ExpectedCellValue(sheetName, 46, 32, "Nashville"),
					new ExpectedCellValue(sheetName, 47, 32, "Car Rack"),
					new ExpectedCellValue(sheetName, 53, 32, 415.75),
					new ExpectedCellValue(sheetName, 55, 32, 415.75),

					new ExpectedCellValue(sheetName, 46, 33, "Nashville Total"),
					new ExpectedCellValue(sheetName, 53, 33, 2d),
					new ExpectedCellValue(sheetName, 55, 33, 2d),

					new ExpectedCellValue(sheetName, 44, 34, "March Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 50, 34, 1d),
					new ExpectedCellValue(sheetName, 53, 34, 2d),
					new ExpectedCellValue(sheetName, 55, 34, 3d),

					new ExpectedCellValue(sheetName, 44, 35, "March Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 50, 35, 24.99),
					new ExpectedCellValue(sheetName, 53, 35, 415.75),
					new ExpectedCellValue(sheetName, 55, 35, 440.74),

					new ExpectedCellValue(sheetName, 44, 36, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 48, 36, 1d),
					new ExpectedCellValue(sheetName, 49, 36, 1d),
					new ExpectedCellValue(sheetName, 50, 36, 1d),
					new ExpectedCellValue(sheetName, 51, 36, 2d),
					new ExpectedCellValue(sheetName, 52, 36, 6d),
					new ExpectedCellValue(sheetName, 53, 36, 2d),
					new ExpectedCellValue(sheetName, 54, 36, 2d),
					new ExpectedCellValue(sheetName, 55, 36, 15d),

					new ExpectedCellValue(sheetName, 44, 37, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 48, 37, 415.75),
					new ExpectedCellValue(sheetName, 49, 37, 99d),
					new ExpectedCellValue(sheetName, 50, 37, 24.99),
					new ExpectedCellValue(sheetName, 51, 37, 415.75),
					new ExpectedCellValue(sheetName, 52, 37, 199d),
					new ExpectedCellValue(sheetName, 53, 37, 415.75),
					new ExpectedCellValue(sheetName, 54, 37, 415.75),
					new ExpectedCellValue(sheetName, 55, 37, 1985.99)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableColumnFields.xlsx")]
		public void PivotTableRefreshMultipleColumnDataFieldsAsFirstInnerChildSubtotalsOff()
		{
			var file = new FileInfo("PivotTableColumnFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["ColumnDataFields"];
					var pivotTable = worksheet.PivotTables["ColumnDataFieldsPivotTable6"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(7, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(3, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(3, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					foreach (var field in pivotTable.Fields)
					{
						if (field.Items.Count > 0)
							this.CheckFieldItems(field);
					}
					package.SaveAs(newFile.File);
				}
				string sheetName = "ColumnDataFields";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 64, 1, 20100076),
					new ExpectedCellValue(sheetName, 65, 1, 20100085),
					new ExpectedCellValue(sheetName, 66, 1, 20100083),
					new ExpectedCellValue(sheetName, 67, 1, 20100007),
					new ExpectedCellValue(sheetName, 68, 1, 20100070),
					new ExpectedCellValue(sheetName, 69, 1, 20100017),
					new ExpectedCellValue(sheetName, 70, 1, 20100090),
					new ExpectedCellValue(sheetName, 71, 1, "Grand Total"),

					new ExpectedCellValue(sheetName, 60, 2, "January"),
					new ExpectedCellValue(sheetName, 61, 2, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 62, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 63, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 64, 2, 1d),
					new ExpectedCellValue(sheetName, 71, 2, 1d),

					new ExpectedCellValue(sheetName, 62, 3, "Chicago"),
					new ExpectedCellValue(sheetName, 63, 3, "Car Rack"),
					new ExpectedCellValue(sheetName, 67, 3, 2d),
					new ExpectedCellValue(sheetName, 71, 3, 2d),

					new ExpectedCellValue(sheetName, 62, 4, "Nashville"),
					new ExpectedCellValue(sheetName, 63, 4, "Car Rack"),
					new ExpectedCellValue(sheetName, 70, 4, 2d),
					new ExpectedCellValue(sheetName, 71, 4, 2d),

					new ExpectedCellValue(sheetName, 61, 5, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 62, 5, "San Francisco"),
					new ExpectedCellValue(sheetName, 63, 5, "Car Rack"),
					new ExpectedCellValue(sheetName, 64, 5, 415.75),
					new ExpectedCellValue(sheetName, 71, 5, 415.75),

					new ExpectedCellValue(sheetName, 62, 6, "Chicago"),
					new ExpectedCellValue(sheetName, 63, 6, "Car Rack"),
					new ExpectedCellValue(sheetName, 67, 6, 415.75),
					new ExpectedCellValue(sheetName, 71, 6, 415.75),

					new ExpectedCellValue(sheetName, 62, 7, "Nashville"),
					new ExpectedCellValue(sheetName, 63, 7, "Car Rack"),
					new ExpectedCellValue(sheetName, 70, 7, 415.75),
					new ExpectedCellValue(sheetName, 71, 7, 415.75),

					new ExpectedCellValue(sheetName, 60, 8, "February"),
					new ExpectedCellValue(sheetName, 61, 8, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 62, 8, "San Francisco"),
					new ExpectedCellValue(sheetName, 63, 8, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 65, 8, 1d),
					new ExpectedCellValue(sheetName, 71, 8, 1d),

					new ExpectedCellValue(sheetName, 62, 9, "Nashville"),
					new ExpectedCellValue(sheetName, 63, 9, "Tent"),
					new ExpectedCellValue(sheetName, 68, 9, 6d),
					new ExpectedCellValue(sheetName, 71, 9, 6d),

					new ExpectedCellValue(sheetName, 61, 10, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 62, 10, "San Francisco"),
					new ExpectedCellValue(sheetName, 63, 10, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 65, 10, 99d),
					new ExpectedCellValue(sheetName, 71, 10, 99d),

					new ExpectedCellValue(sheetName, 62, 11, "Nashville"),
					new ExpectedCellValue(sheetName, 63, 11, "Tent"),
					new ExpectedCellValue(sheetName, 68, 11, 199d),
					new ExpectedCellValue(sheetName, 71, 11, 199d),

					new ExpectedCellValue(sheetName, 60, 12, "March"),
					new ExpectedCellValue(sheetName, 61, 12, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 62, 12, "Chicago"),
					new ExpectedCellValue(sheetName, 63, 12, "Headlamp"),
					new ExpectedCellValue(sheetName, 66, 12, 1d),
					new ExpectedCellValue(sheetName, 71, 12, 1d),

					new ExpectedCellValue(sheetName, 62, 13, "Nashville"),
					new ExpectedCellValue(sheetName, 63, 13, "Car Rack"),
					new ExpectedCellValue(sheetName, 69, 13, 2d),
					new ExpectedCellValue(sheetName, 71, 13, 2d),

					new ExpectedCellValue(sheetName, 61, 14, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 62, 14, "Chicago"),
					new ExpectedCellValue(sheetName, 63, 14, "Headlamp"),
					new ExpectedCellValue(sheetName, 66, 14, 24.99),
					new ExpectedCellValue(sheetName, 71, 14, 24.99),

					new ExpectedCellValue(sheetName, 62, 15, "Nashville"),
					new ExpectedCellValue(sheetName, 63, 15, "Car Rack"),
					new ExpectedCellValue(sheetName, 69, 15, 415.75),
					new ExpectedCellValue(sheetName, 71, 15, 415.75),

					new ExpectedCellValue(sheetName, 60, 16, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 64, 16, 1d),
					new ExpectedCellValue(sheetName, 65, 16, 1d),
					new ExpectedCellValue(sheetName, 66, 16, 1d),
					new ExpectedCellValue(sheetName, 67, 16, 2d),
					new ExpectedCellValue(sheetName, 68, 16, 6d),
					new ExpectedCellValue(sheetName, 69, 16, 2d),
					new ExpectedCellValue(sheetName, 70, 16, 2d),
					new ExpectedCellValue(sheetName, 71, 16, 15d),

					new ExpectedCellValue(sheetName, 60, 17, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 64, 17, 415.75),
					new ExpectedCellValue(sheetName, 65, 17, 99d),
					new ExpectedCellValue(sheetName, 66, 17, 24.99),
					new ExpectedCellValue(sheetName, 67, 17, 415.75),
					new ExpectedCellValue(sheetName, 68, 17, 199d),
					new ExpectedCellValue(sheetName, 69, 17, 415.75),
					new ExpectedCellValue(sheetName, 70, 17, 415.75),
					new ExpectedCellValue(sheetName, 71, 17, 1985.99)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableColumnFields.xlsx")]
		public void PivotTableRefreshMultipleColumnDataFieldsAsLastInnerChildSubtotalsOn()
		{
			var file = new FileInfo("PivotTableColumnFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["ColumnDataFields"];
					var pivotTable = worksheet.PivotTables["ColumnDataFieldsPivotTable7"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(8, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(5, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					foreach (var field in pivotTable.Fields)
					{
						if (field.Items.Count > 0)
							this.CheckFieldItems(field);
					}
					package.SaveAs(newFile.File);
				}
				string sheetName = "ColumnDataFields";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 80, 1, 20100076),
					new ExpectedCellValue(sheetName, 81, 1, 20100085),
					new ExpectedCellValue(sheetName, 82, 1, 20100083),
					new ExpectedCellValue(sheetName, 83, 1, 20100007),
					new ExpectedCellValue(sheetName, 84, 1, 20100070),
					new ExpectedCellValue(sheetName, 85, 1, 20100017),
					new ExpectedCellValue(sheetName, 86, 1, 20100090),
					new ExpectedCellValue(sheetName, 87, 1, "Grand Total"),

					new ExpectedCellValue(sheetName, 76, 2, "January"),
					new ExpectedCellValue(sheetName, 77, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 78, 2, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 79, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 80, 2, 1d),
					new ExpectedCellValue(sheetName, 87, 2, 1d),

					new ExpectedCellValue(sheetName, 78, 3, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 79, 3, "Car Rack"),
					new ExpectedCellValue(sheetName, 80, 3, 415.75),
					new ExpectedCellValue(sheetName, 87, 3, 415.75),

					new ExpectedCellValue(sheetName, 77, 4, "San Francisco Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 80, 4, 1d),
					new ExpectedCellValue(sheetName, 87, 4, 1d),

					new ExpectedCellValue(sheetName, 77, 5, "San Francisco Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 80, 5, 415.75),
					new ExpectedCellValue(sheetName, 87, 5, 415.75),

					new ExpectedCellValue(sheetName, 77, 6, "Chicago"),
					new ExpectedCellValue(sheetName, 78, 6, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 79, 6, "Car Rack"),
					new ExpectedCellValue(sheetName, 83, 6, 2d),
					new ExpectedCellValue(sheetName, 87, 6, 2d),

					new ExpectedCellValue(sheetName, 78, 7, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 79, 7, "Car Rack"),
					new ExpectedCellValue(sheetName, 83, 7, 415.75),
					new ExpectedCellValue(sheetName, 87, 7, 415.75),

					new ExpectedCellValue(sheetName, 77, 8, "Chicago Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 83, 8, 2d),
					new ExpectedCellValue(sheetName, 87, 8, 2d),

					new ExpectedCellValue(sheetName, 77, 9, "Chicago Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 83, 9, 415.75),
					new ExpectedCellValue(sheetName, 87, 9, 415.75),

					new ExpectedCellValue(sheetName, 77, 10, "Nashville"),
					new ExpectedCellValue(sheetName, 78, 10, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 79, 10, "Car Rack"),
					new ExpectedCellValue(sheetName, 86, 10, 2d),
					new ExpectedCellValue(sheetName, 87, 10, 2d),

					new ExpectedCellValue(sheetName, 78, 11, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 79, 11, "Car Rack"),
					new ExpectedCellValue(sheetName, 86, 11, 415.75),
					new ExpectedCellValue(sheetName, 87, 11, 415.75),

					new ExpectedCellValue(sheetName, 77, 12, "Nashville Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 86, 12, 2d),
					new ExpectedCellValue(sheetName, 87, 12, 2d),

					new ExpectedCellValue(sheetName, 77, 13, "Nashville Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 86, 13, 415.75),
					new ExpectedCellValue(sheetName, 87, 13, 415.75),

					new ExpectedCellValue(sheetName, 76, 14, "January Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 80, 14, 1d),
					new ExpectedCellValue(sheetName, 83, 14, 2d),
					new ExpectedCellValue(sheetName, 86, 14, 2d),
					new ExpectedCellValue(sheetName, 87, 14, 5d),

					new ExpectedCellValue(sheetName, 76, 15, "January Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 80, 15, 415.75),
					new ExpectedCellValue(sheetName, 83, 15, 415.75),
					new ExpectedCellValue(sheetName, 86, 15, 415.75),
					new ExpectedCellValue(sheetName, 87, 15, 1247.25),

					new ExpectedCellValue(sheetName, 76, 16, "February"),
					new ExpectedCellValue(sheetName, 77, 16, "San Francisco"),
					new ExpectedCellValue(sheetName, 78, 16, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 79, 16, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 81, 16, 1d),
					new ExpectedCellValue(sheetName, 87, 16, 1d),

					new ExpectedCellValue(sheetName, 78, 17, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 79, 17, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 81, 17, 99d),
					new ExpectedCellValue(sheetName, 87, 17, 99d),

					new ExpectedCellValue(sheetName, 77, 18, "San Francisco Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 81, 18, 1d),
					new ExpectedCellValue(sheetName, 87, 18, 1d),

					new ExpectedCellValue(sheetName, 77, 19, "San Francisco Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 81, 19, 99d),
					new ExpectedCellValue(sheetName, 87, 19, 99d),

					new ExpectedCellValue(sheetName, 77, 20, "Nashville"),
					new ExpectedCellValue(sheetName, 78, 20, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 79, 20, "Tent"),
					new ExpectedCellValue(sheetName, 84, 20, 6d),
					new ExpectedCellValue(sheetName, 87, 20, 6d),

					new ExpectedCellValue(sheetName, 78, 21, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 79, 21, "Tent"),
					new ExpectedCellValue(sheetName, 84, 21, 199d),
					new ExpectedCellValue(sheetName, 87, 21, 199d),

					new ExpectedCellValue(sheetName, 77, 22, "Nashville Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 84, 22, 6d),
					new ExpectedCellValue(sheetName, 87, 22, 6d),

					new ExpectedCellValue(sheetName, 77, 23, "Nashville Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 84, 23, 199d),
					new ExpectedCellValue(sheetName, 87, 23, 199d),

					new ExpectedCellValue(sheetName, 76, 24, "February Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 81, 24, 1d),
					new ExpectedCellValue(sheetName, 84, 24, 6d),
					new ExpectedCellValue(sheetName, 87, 24, 7d),

					new ExpectedCellValue(sheetName, 76, 25, "February Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 81, 25, 99d),
					new ExpectedCellValue(sheetName, 84, 25, 199d),
					new ExpectedCellValue(sheetName, 87, 25, 298d),

					new ExpectedCellValue(sheetName, 76, 26, "March"),
					new ExpectedCellValue(sheetName, 77, 26, "Chicago"),
					new ExpectedCellValue(sheetName, 78, 26, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 79, 26, "Headlamp"),
					new ExpectedCellValue(sheetName, 82, 26, 1d),
					new ExpectedCellValue(sheetName, 87, 26, 1d),

					new ExpectedCellValue(sheetName, 78, 27, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 79, 27, "Headlamp"),
					new ExpectedCellValue(sheetName, 82, 27, 24.99),
					new ExpectedCellValue(sheetName, 87, 27, 24.99),

					new ExpectedCellValue(sheetName, 77, 28, "Chicago Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 82, 28, 1d),
					new ExpectedCellValue(sheetName, 87, 28, 1d),

					new ExpectedCellValue(sheetName, 77, 29, "Chicago Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 82, 29, 24.99),
					new ExpectedCellValue(sheetName, 87, 29, 24.99),

					new ExpectedCellValue(sheetName, 77, 30, "Nashville"),
					new ExpectedCellValue(sheetName, 78, 30, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 79, 30, "Car Rack"),
					new ExpectedCellValue(sheetName, 85, 30, 2d),
					new ExpectedCellValue(sheetName, 87, 30, 2d),

					new ExpectedCellValue(sheetName, 78, 31, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 79, 31, "Car Rack"),
					new ExpectedCellValue(sheetName, 85, 31, 415.75),
					new ExpectedCellValue(sheetName, 87, 31, 415.75),

					new ExpectedCellValue(sheetName, 77, 32, "Nashville Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 85, 32, 2d),
					new ExpectedCellValue(sheetName, 87, 32, 2d),

					new ExpectedCellValue(sheetName, 77, 33, "Nashville Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 85, 33, 415.75),
					new ExpectedCellValue(sheetName, 87, 33, 415.75),

					new ExpectedCellValue(sheetName, 76, 34, "March Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 82, 34, 1d),
					new ExpectedCellValue(sheetName, 85, 34, 2d),
					new ExpectedCellValue(sheetName, 87, 34, 3d),

					new ExpectedCellValue(sheetName, 76, 35, "March Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 82, 35, 24.99),
					new ExpectedCellValue(sheetName, 85, 35, 415.75),
					new ExpectedCellValue(sheetName, 87, 35, 440.74),

					new ExpectedCellValue(sheetName, 76, 36, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 80, 36, 1d),
					new ExpectedCellValue(sheetName, 81, 36, 1d),
					new ExpectedCellValue(sheetName, 82, 36, 1d),
					new ExpectedCellValue(sheetName, 83, 36, 2d),
					new ExpectedCellValue(sheetName, 84, 36, 6d),
					new ExpectedCellValue(sheetName, 85, 36, 2d),
					new ExpectedCellValue(sheetName, 86, 36, 2d),
					new ExpectedCellValue(sheetName, 87, 36, 15d),

					new ExpectedCellValue(sheetName, 76, 37, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 80, 37, 415.75),
					new ExpectedCellValue(sheetName, 81, 37, 99d),
					new ExpectedCellValue(sheetName, 82, 37, 24.99),
					new ExpectedCellValue(sheetName, 83, 37, 415.75),
					new ExpectedCellValue(sheetName, 84, 37, 199d),
					new ExpectedCellValue(sheetName, 85, 37, 415.75),
					new ExpectedCellValue(sheetName, 86, 37, 415.75),
					new ExpectedCellValue(sheetName, 87, 37, 1985.99)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableColumnFields.xlsx")]
		public void PivotTableRefreshMultipleColumnDataFieldsAsLastInnerChildSubtotalsOff()
		{
			var file = new FileInfo("PivotTableColumnFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["ColumnDataFields"];
					var pivotTable = worksheet.PivotTables["ColumnDataFieldsPivotTable8"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					Assert.AreEqual(7, pivotTable.Fields.Count);
					Assert.AreEqual(7, pivotTable.Fields[0].Items.Count);
					Assert.AreEqual(3, pivotTable.Fields[1].Items.Count);
					Assert.AreEqual(3, pivotTable.Fields[2].Items.Count);
					Assert.AreEqual(4, pivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, pivotTable.Fields[6].Items.Count);
					foreach (var field in pivotTable.Fields)
					{
						if (field.Items.Count > 0)
							this.CheckFieldItems(field);
					}
					package.SaveAs(newFile.File);
				}
				string sheetName = "ColumnDataFields";
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 96, 1, 20100076),
					new ExpectedCellValue(sheetName, 97, 1, 20100085),
					new ExpectedCellValue(sheetName, 98, 1, 20100083),
					new ExpectedCellValue(sheetName, 99, 1, 20100007),
					new ExpectedCellValue(sheetName, 100, 1, 20100070),
					new ExpectedCellValue(sheetName, 101, 1, 20100017),
					new ExpectedCellValue(sheetName, 102, 1, 20100090),
					new ExpectedCellValue(sheetName, 103, 1, "Grand Total"),

					new ExpectedCellValue(sheetName, 92, 2, "January"),
					new ExpectedCellValue(sheetName, 93, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 94, 2, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 95, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 96, 2, 1d),
					new ExpectedCellValue(sheetName, 103, 2, 1d),

					new ExpectedCellValue(sheetName, 94, 3, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 95, 3, "Car Rack"),
					new ExpectedCellValue(sheetName, 96, 3, 415.75),
					new ExpectedCellValue(sheetName, 103, 3, 415.75),

					new ExpectedCellValue(sheetName, 93, 4, "Chicago"),
					new ExpectedCellValue(sheetName, 94, 4, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 95, 4, "Car Rack"),
					new ExpectedCellValue(sheetName, 99, 4, 2d),
					new ExpectedCellValue(sheetName, 103, 4, 2d),

					new ExpectedCellValue(sheetName, 94, 5, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 95, 5, "Car Rack"),
					new ExpectedCellValue(sheetName, 99, 5, 415.75),
					new ExpectedCellValue(sheetName, 103, 5, 415.75),

					new ExpectedCellValue(sheetName, 93, 6, "Nashville"),
					new ExpectedCellValue(sheetName, 94, 6, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 95, 6, "Car Rack"),
					new ExpectedCellValue(sheetName, 102, 6, 2d),
					new ExpectedCellValue(sheetName, 103, 6, 2d),

					new ExpectedCellValue(sheetName, 94, 7, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 95, 7, "Car Rack"),
					new ExpectedCellValue(sheetName, 102, 7, 415.75),
					new ExpectedCellValue(sheetName, 103, 7, 415.75),

					new ExpectedCellValue(sheetName, 92, 8, "February"),
					new ExpectedCellValue(sheetName, 93, 8, "San Francisco"),
					new ExpectedCellValue(sheetName, 94, 8, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 95, 8, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 97, 8, 1d),
					new ExpectedCellValue(sheetName, 103, 8, 1d),

					new ExpectedCellValue(sheetName, 94, 9, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 95, 9, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 97, 9, 99d),
					new ExpectedCellValue(sheetName, 103, 9, 99d),

					new ExpectedCellValue(sheetName, 93, 10, "Nashville"),
					new ExpectedCellValue(sheetName, 94, 10, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 95, 10, "Tent"),
					new ExpectedCellValue(sheetName, 100, 10, 6d),
					new ExpectedCellValue(sheetName, 103, 10, 6d),

					new ExpectedCellValue(sheetName, 94, 11, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 95, 11, "Tent"),
					new ExpectedCellValue(sheetName, 100, 11, 199d),
					new ExpectedCellValue(sheetName, 103, 11, 199d),

					new ExpectedCellValue(sheetName, 92, 12, "March"),
					new ExpectedCellValue(sheetName, 93, 12, "Chicago"),
					new ExpectedCellValue(sheetName, 94, 12, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 95, 12, "Headlamp"),
					new ExpectedCellValue(sheetName, 98, 12, 1d),
					new ExpectedCellValue(sheetName, 103, 12, 1d),

					new ExpectedCellValue(sheetName, 94, 13, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 95, 13, "Headlamp"),
					new ExpectedCellValue(sheetName, 98, 13, 24.99),
					new ExpectedCellValue(sheetName, 103, 13, 24.99),

					new ExpectedCellValue(sheetName, 93, 14, "Nashville"),
					new ExpectedCellValue(sheetName, 94, 14, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 95, 14, "Car Rack"),
					new ExpectedCellValue(sheetName, 101, 14, 2d),
					new ExpectedCellValue(sheetName, 103, 14, 2d),

					new ExpectedCellValue(sheetName, 94, 15, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 95, 15, "Car Rack"),
					new ExpectedCellValue(sheetName, 101, 15, 415.75),
					new ExpectedCellValue(sheetName, 103, 15, 415.75),

					new ExpectedCellValue(sheetName, 92, 16, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 96, 16, 1d),
					new ExpectedCellValue(sheetName, 97, 16, 1d),
					new ExpectedCellValue(sheetName, 98, 16, 1d),
					new ExpectedCellValue(sheetName, 99, 16, 2d),
					new ExpectedCellValue(sheetName, 100, 16, 6d),
					new ExpectedCellValue(sheetName, 101, 16, 2d),
					new ExpectedCellValue(sheetName, 102, 16, 2d),
					new ExpectedCellValue(sheetName, 103, 16, 15d),

					new ExpectedCellValue(sheetName, 92, 17, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 96, 17, 415.75),
					new ExpectedCellValue(sheetName, 97, 17, 99d),
					new ExpectedCellValue(sheetName, 98, 17, 24.99),
					new ExpectedCellValue(sheetName, 99, 17, 415.75),
					new ExpectedCellValue(sheetName, 100, 17, 199d),
					new ExpectedCellValue(sheetName, 101, 17, 415.75),
					new ExpectedCellValue(sheetName, 102, 17, 415.75),
					new ExpectedCellValue(sheetName, 103, 17, 1985.99)
				});
			}
		}
		#endregion
		#endregion

		#region Helper Methods
		private void CheckFieldItems(ExcelPivotTableField field)
		{
			int i = 0;
			for (; i < field.Items.Count - 1; i++)
			{
				Assert.AreEqual(i, field.Items[i].X);
			}
			var lastItem = field.Items[field.Items.Count - 1];
			if (string.IsNullOrEmpty(lastItem.T))
				Assert.AreEqual(i, lastItem.X);
		}
		#endregion
	}
}