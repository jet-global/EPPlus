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
				cacheDefinition.SourceRange = worksheet.Cells["C3:F7"];
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
				cacheDefinition.SourceRange = worksheet.Cells["C3:F5"];
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
				cacheDefinition.SourceRange = worksheet.Cells["A1:G9"];
				cacheDefinition.UpdateData();
				Assert.AreEqual(7, pivotTable.Fields.Count);
				Assert.AreEqual(9, pivotTable.Fields[0].Items.Count);
				Assert.AreEqual(5, pivotTable.Fields[1].Items.Count);
				Assert.AreEqual(5, pivotTable.Fields[2].Items.Count);
				Assert.AreEqual(6, pivotTable.Fields[3].Items.Count);
				Assert.AreEqual(0, pivotTable.Fields[4].Items.Count);
				Assert.AreEqual(4, pivotTable.Fields[5].Items.Count);
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
				cacheDefinition.SourceRange = worksheet.Cells["A1:G5"];
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
					cacheDefinition.SourceRange = worksheet.Cells["A1:G7"];
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
					package.SaveAs(new FileInfo(@"C:\Users\mcl\Downloads\PivotTables\ThreeRowFields_FalseSubtotalTop.xlsx"));
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
					package.SaveAs(new FileInfo(@"C:\Users\mcl\Downloads\PivotTables\NoSubtotalTop.xlsx"));
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
					package.SaveAs(new FileInfo(@"C:\Users\mcl\Downloads\PivotTables\Book3_RowGrandTotalsOff.xlsx"));
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
					package.SaveAs(new FileInfo(@"C:\Users\mcl\Downloads\PivotTables\Book3_RowGrandTotalsOff.xlsx"));
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