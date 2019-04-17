using System;
using System.IO;
using System.Linq;
using EPPlusTest.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Table.PivotTable;

namespace EPPlusTest.Table.PivotTable.PivotTableRefresh
{
	[TestClass]
	public class PivotTableRefreshShowDataAsTest
	{
		#region ShowDataAs Tests
		#region PercentOfTotal Tests
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableShowDataAs.xlsx")]
		public void PivotTableRefreshShowDataAsRowFieldsColumnDataFields()
		{
			var file = new FileInfo("PivotTableShowDataAs.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				void validateSheet() => TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 2, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 2, 3, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 2, 4, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 3, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 4, 2, "January"),
					new ExpectedCellValue(sheetName, 4, 3, 0.209341437),
					new ExpectedCellValue(sheetName, 4, 4, 2),
					new ExpectedCellValue(sheetName, 5, 2, "March"),
					new ExpectedCellValue(sheetName, 5, 3, 0.012583145),
					new ExpectedCellValue(sheetName, 5, 4, 1),
					new ExpectedCellValue(sheetName, 6, 2, "Nashville"),
					new ExpectedCellValue(sheetName, 7, 2, "January"),
					new ExpectedCellValue(sheetName, 7, 3, 0.209341437),
					new ExpectedCellValue(sheetName, 7, 4, 2),
					new ExpectedCellValue(sheetName, 8, 2, "February"),
					new ExpectedCellValue(sheetName, 8, 3, 0.100201914),
					new ExpectedCellValue(sheetName, 8, 4, 6),
					new ExpectedCellValue(sheetName, 9, 2, "March"),
					new ExpectedCellValue(sheetName, 9, 3, 0.209341437),
					new ExpectedCellValue(sheetName, 9, 4, 2),
					new ExpectedCellValue(sheetName, 10, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 11, 2, "January"),
					new ExpectedCellValue(sheetName, 11, 3, 0.209341437),
					new ExpectedCellValue(sheetName, 11, 4, 1),
					new ExpectedCellValue(sheetName, 12, 2, "February"),
					new ExpectedCellValue(sheetName, 12, 3, 0.049849194),
					new ExpectedCellValue(sheetName, 12, 4, 1),
					new ExpectedCellValue(sheetName, 13, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 13, 3, 1),
					new ExpectedCellValue(sheetName, 13, 4, 15)
				});

				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable1"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					wholesalePriceDataField.ShowDataAs = ShowDataAs.PercentOfTotal;
					unitsSoldDataField.ShowDataAs = ShowDataAs.NoCalculation;
					foreach (var field in pivotTable.Fields)
					{
						field.SubtotalLocation = SubtotalLocation.Top;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B2:D13"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				validateSheet();
				// Validate that top subtotals were written in.
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 3, 3, .2219),
					new ExpectedCellValue(sheetName, 3, 4, 3),
					new ExpectedCellValue(sheetName, 6, 3, .5189),
					new ExpectedCellValue(sheetName, 6, 4, 10),
					new ExpectedCellValue(sheetName, 10, 3, .2592),
					new ExpectedCellValue(sheetName, 10, 4, 2),
				});

				// Test again with subtotals turned off.
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable1"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					wholesalePriceDataField.ShowDataAs = ShowDataAs.PercentOfTotal;
					unitsSoldDataField.ShowDataAs = ShowDataAs.NoCalculation;
					foreach (var field in pivotTable.Fields)
					{
						field.SubTotalFunctions = eSubTotalFunctions.None;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B2:D13"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				validateSheet();
				// Validate that no subtotals were written in.
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 3, 3, null),
					new ExpectedCellValue(sheetName, 3, 4, null),
					new ExpectedCellValue(sheetName, 6, 3, null),
					new ExpectedCellValue(sheetName, 6, 4, null),
					new ExpectedCellValue(sheetName, 10, 3, null),
					new ExpectedCellValue(sheetName, 10, 4, null),
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableShowDataAs.xlsx")]
		public void PivotTableRefreshShowDataAsPercentOfGrandTotalRowFieldNoColumnFields()
		{
			var file = new FileInfo("PivotTableShowDataAs.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable7"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Count of Wholesale Price");
					wholesalePriceDataField.ShowDataAs = ShowDataAs.PercentOfTotal;
					foreach (var field in pivotTable.Fields)
					{
						field.SubtotalLocation = SubtotalLocation.Top;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B93:C97"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 93, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 93, 3, "Count of Wholesale Price"),
					new ExpectedCellValue(sheetName, 94, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 94, 3, .2857),
					new ExpectedCellValue(sheetName, 95, 2, "Nashville"),
					new ExpectedCellValue(sheetName, 95, 3, .4286),
					new ExpectedCellValue(sheetName, 96, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 96, 3, 0.2857),
					new ExpectedCellValue(sheetName, 97, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 97, 3, 1),
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableShowDataAs.xlsx")]
		public void PivotTableRefreshShowDataAsColumnFieldsRowDataFields()
		{
			var file = new FileInfo("PivotTableShowDataAs.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable2"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					wholesalePriceDataField.ShowDataAs = ShowDataAs.PercentOfTotal;
					unitsSoldDataField.ShowDataAs = ShowDataAs.NoCalculation;
					foreach (var field in pivotTable.Fields)
					{
						field.SubTotalFunctions = eSubTotalFunctions.Default;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B18:K22"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 18, 2, null),
					new ExpectedCellValue(sheetName, 18, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 18, 4, null),
					new ExpectedCellValue(sheetName, 19, 2, null),
					new ExpectedCellValue(sheetName, 19, 3, "January"),
					new ExpectedCellValue(sheetName, 19, 4, "January Total"),
					new ExpectedCellValue(sheetName, 19, 5, "February"),
					new ExpectedCellValue(sheetName, 19, 6, null),
					new ExpectedCellValue(sheetName, 19, 7, "February Total"),
					new ExpectedCellValue(sheetName, 19, 8, "March"),
					new ExpectedCellValue(sheetName, 19, 9, null),
					new ExpectedCellValue(sheetName, 19, 10, "March Total"),
					new ExpectedCellValue(sheetName, 19, 11, "Grand Total"),
					new ExpectedCellValue(sheetName, 20, 2, "Values"),
					new ExpectedCellValue(sheetName, 20, 3, "Car Rack"),
					new ExpectedCellValue(sheetName, 20, 4, null),
					new ExpectedCellValue(sheetName, 20, 5, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 20, 6, "Tent"),
					new ExpectedCellValue(sheetName, 20, 7, null),
					new ExpectedCellValue(sheetName, 20, 8, "Car Rack"),
					new ExpectedCellValue(sheetName, 20, 9, "Headlamp"),
					new ExpectedCellValue(sheetName, 20, 10, null),
					new ExpectedCellValue(sheetName, 20, 11, null),
					new ExpectedCellValue(sheetName, 21, 2, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 21, 3, .6280),
					new ExpectedCellValue(sheetName, 21, 4, .6280),
					new ExpectedCellValue(sheetName, 21, 5, .0498),
					new ExpectedCellValue(sheetName, 21, 6, .1002),
					new ExpectedCellValue(sheetName, 21, 7, .1501),
					new ExpectedCellValue(sheetName, 21, 8, .2093),
					new ExpectedCellValue(sheetName, 21, 9, .0126),
					new ExpectedCellValue(sheetName, 21, 10, .2219),
					new ExpectedCellValue(sheetName, 21, 11, 1),
					new ExpectedCellValue(sheetName, 22, 2, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 22, 3, 5),
					new ExpectedCellValue(sheetName, 22, 4, 5),
					new ExpectedCellValue(sheetName, 22, 5, 1),
					new ExpectedCellValue(sheetName, 22, 6, 6),
					new ExpectedCellValue(sheetName, 22, 7, 7),
					new ExpectedCellValue(sheetName, 22, 8, 2),
					new ExpectedCellValue(sheetName, 22, 9, 1),
					new ExpectedCellValue(sheetName, 22, 10, 3),
					new ExpectedCellValue(sheetName, 22, 11, 15)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableShowDataAs.xlsx")]
		public void PivotTableRefreshShowDataAsColumnFieldsRowDataFieldsSubtotalsOff()
		{
			var file = new FileInfo("PivotTableShowDataAs.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable2"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					wholesalePriceDataField.ShowDataAs = ShowDataAs.PercentOfTotal;
					unitsSoldDataField.ShowDataAs = ShowDataAs.NoCalculation;
					foreach (var field in pivotTable.Fields)
					{
						field.SubTotalFunctions = eSubTotalFunctions.None;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B18:H22"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 18, 2, null),
					new ExpectedCellValue(sheetName, 18, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 18, 4, null),
					new ExpectedCellValue(sheetName, 19, 2, null),
					new ExpectedCellValue(sheetName, 19, 3, "January"),
					new ExpectedCellValue(sheetName, 19, 4, "February"),
					new ExpectedCellValue(sheetName, 19, 5, null),
					new ExpectedCellValue(sheetName, 19, 6, "March"),
					new ExpectedCellValue(sheetName, 19, 7, null),
					new ExpectedCellValue(sheetName, 19, 8, "Grand Total"),
					new ExpectedCellValue(sheetName, 20, 2, "Values"),
					new ExpectedCellValue(sheetName, 20, 3, "Car Rack"),
					new ExpectedCellValue(sheetName, 20, 4, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 20, 5, "Tent"),
					new ExpectedCellValue(sheetName, 20, 6, "Car Rack"),
					new ExpectedCellValue(sheetName, 20, 7, "Headlamp"),
					new ExpectedCellValue(sheetName, 20, 8, null),
					new ExpectedCellValue(sheetName, 21, 2, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 21, 3, .6280),
					new ExpectedCellValue(sheetName, 21, 4, .0498),
					new ExpectedCellValue(sheetName, 21, 5, .1002),
					new ExpectedCellValue(sheetName, 21, 6, .2093),
					new ExpectedCellValue(sheetName, 21, 7, .0126),
					new ExpectedCellValue(sheetName, 21, 8, 1),
					new ExpectedCellValue(sheetName, 22, 2, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 22, 3, 5),
					new ExpectedCellValue(sheetName, 22, 4, 1),
					new ExpectedCellValue(sheetName, 22, 5, 6),
					new ExpectedCellValue(sheetName, 22, 6, 2),
					new ExpectedCellValue(sheetName, 22, 7, 1),
					new ExpectedCellValue(sheetName, 22, 8, 15)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableShowDataAs.xlsx")]
		public void PivotTableRefreshShowDataAsRowFieldsColumnFieldsAndColumnDataFieldsSubtotalsTop()
		{
			var file = new FileInfo("PivotTableShowDataAs.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable3"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					wholesalePriceDataField.ShowDataAs = ShowDataAs.PercentOfTotal;
					unitsSoldDataField.ShowDataAs = ShowDataAs.NoCalculation;
					foreach (var field in pivotTable.Fields)
					{
						field.SubtotalLocation = SubtotalLocation.Top;
						field.SubTotalFunctions = eSubTotalFunctions.Default;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B30:J41"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 31, 2, null),
					new ExpectedCellValue(sheetName, 31, 3, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 31, 4, null),
					new ExpectedCellValue(sheetName, 31, 5, null),
					new ExpectedCellValue(sheetName, 31, 6, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 31, 7, null),
					new ExpectedCellValue(sheetName, 31, 8, null),
					new ExpectedCellValue(sheetName, 31, 9, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 31, 10, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 32, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 32, 3, "Chicago"),
					new ExpectedCellValue(sheetName, 32, 4, "Nashville"),
					new ExpectedCellValue(sheetName, 32, 5, "San Francisco"),
					new ExpectedCellValue(sheetName, 32, 6, "Chicago"),
					new ExpectedCellValue(sheetName, 32, 7, "Nashville"),
					new ExpectedCellValue(sheetName, 32, 8, "San Francisco"),
					new ExpectedCellValue(sheetName, 32, 9, null),
					new ExpectedCellValue(sheetName, 32, 10, null),
					new ExpectedCellValue(sheetName, 33, 2, "January"),
					new ExpectedCellValue(sheetName, 33, 3, .2093),
					new ExpectedCellValue(sheetName, 33, 4, .2093),
					new ExpectedCellValue(sheetName, 33, 5, .2093),
					new ExpectedCellValue(sheetName, 33, 6, 2),
					new ExpectedCellValue(sheetName, 33, 7, 2),
					new ExpectedCellValue(sheetName, 33, 8, 1),
					new ExpectedCellValue(sheetName, 33, 9, .628),
					new ExpectedCellValue(sheetName, 33, 10, 5),
					new ExpectedCellValue(sheetName, 34, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 34, 3, .2093),
					new ExpectedCellValue(sheetName, 34, 4, .2093),
					new ExpectedCellValue(sheetName, 34, 5, .2093),
					new ExpectedCellValue(sheetName, 34, 6, 2),
					new ExpectedCellValue(sheetName, 34, 7, 2),
					new ExpectedCellValue(sheetName, 34, 8, 1),
					new ExpectedCellValue(sheetName, 34, 9, .628),
					new ExpectedCellValue(sheetName, 34, 10, 5),
					new ExpectedCellValue(sheetName, 35, 2, "February"),
					new ExpectedCellValue(sheetName, 35, 3, 0),
					new ExpectedCellValue(sheetName, 35, 4, .1002),
					new ExpectedCellValue(sheetName, 35, 5, .0498),
					new ExpectedCellValue(sheetName, 35, 6, null),
					new ExpectedCellValue(sheetName, 35, 7, 6),
					new ExpectedCellValue(sheetName, 35, 8, 1),
					new ExpectedCellValue(sheetName, 35, 9, .1501),
					new ExpectedCellValue(sheetName, 35, 10, 7),
					new ExpectedCellValue(sheetName, 36, 2, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 36, 3, 0),
					new ExpectedCellValue(sheetName, 36, 4, 0),
					new ExpectedCellValue(sheetName, 36, 5, .0498),
					new ExpectedCellValue(sheetName, 36, 6, null),
					new ExpectedCellValue(sheetName, 36, 7, null),
					new ExpectedCellValue(sheetName, 36, 8, 1),
					new ExpectedCellValue(sheetName, 36, 9, .0498),
					new ExpectedCellValue(sheetName, 36, 10, 1),
					new ExpectedCellValue(sheetName, 37, 2, "Tent"),
					new ExpectedCellValue(sheetName, 37, 3, 0),
					new ExpectedCellValue(sheetName, 37, 4, .1002),
					new ExpectedCellValue(sheetName, 37, 5, 0),
					new ExpectedCellValue(sheetName, 37, 6, null),
					new ExpectedCellValue(sheetName, 37, 7, 6),
					new ExpectedCellValue(sheetName, 37, 8, null),
					new ExpectedCellValue(sheetName, 37, 9, .1002),
					new ExpectedCellValue(sheetName, 37, 10, 6),
					new ExpectedCellValue(sheetName, 38, 2, "March"),
					new ExpectedCellValue(sheetName, 38, 3, .0126),
					new ExpectedCellValue(sheetName, 38, 4, .2093),
					new ExpectedCellValue(sheetName, 38, 5, 0),
					new ExpectedCellValue(sheetName, 38, 6, 1),
					new ExpectedCellValue(sheetName, 38, 7, 2),
					new ExpectedCellValue(sheetName, 38, 8, null),
					new ExpectedCellValue(sheetName, 38, 9, .2219),
					new ExpectedCellValue(sheetName, 38, 10, 3),
					new ExpectedCellValue(sheetName, 39, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 39, 3, 0),
					new ExpectedCellValue(sheetName, 39, 4, .2093),
					new ExpectedCellValue(sheetName, 39, 5, 0),
					new ExpectedCellValue(sheetName, 39, 6, null),
					new ExpectedCellValue(sheetName, 39, 7, 2),
					new ExpectedCellValue(sheetName, 39, 8, null),
					new ExpectedCellValue(sheetName, 39, 9, .2093),
					new ExpectedCellValue(sheetName, 39, 10, 2),
					new ExpectedCellValue(sheetName, 40, 2, "Headlamp"),
					new ExpectedCellValue(sheetName, 40, 3, .0126),
					new ExpectedCellValue(sheetName, 40, 4, 0),
					new ExpectedCellValue(sheetName, 40, 5, 0),
					new ExpectedCellValue(sheetName, 40, 6, 1),
					new ExpectedCellValue(sheetName, 40, 7, null),
					new ExpectedCellValue(sheetName, 40, 8, null),
					new ExpectedCellValue(sheetName, 40, 9, .0126),
					new ExpectedCellValue(sheetName, 40, 10, 1),
					new ExpectedCellValue(sheetName, 41, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 41, 3, .2219),
					new ExpectedCellValue(sheetName, 41, 4, .5189),
					new ExpectedCellValue(sheetName, 41, 5, .2592),
					new ExpectedCellValue(sheetName, 41, 6, 3),
					new ExpectedCellValue(sheetName, 41, 7, 10),
					new ExpectedCellValue(sheetName, 41, 8, 2),
					new ExpectedCellValue(sheetName, 41, 9, 1),
					new ExpectedCellValue(sheetName, 41, 10, 15)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableShowDataAs.xlsx")]
		public void PivotTableRefreshShowDataAsRowFieldsColumnFieldsAndRowDataFields()
		{
			var file = new FileInfo("PivotTableShowDataAs.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable4"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					wholesalePriceDataField.ShowDataAs = ShowDataAs.NoCalculation;
					unitsSoldDataField.ShowDataAs = ShowDataAs.PercentOfTotal;
					foreach (var field in pivotTable.Fields)
					{
						field.SubTotalFunctions = eSubTotalFunctions.None;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B46:F67"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 47, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 47, 3, "Chicago"),
					new ExpectedCellValue(sheetName, 47, 4, "Nashville"),
					new ExpectedCellValue(sheetName, 47, 5, "San Francisco"),
					new ExpectedCellValue(sheetName, 47, 6, "Grand Total"),
					new ExpectedCellValue(sheetName, 48, 2, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 48, 3, null),
					new ExpectedCellValue(sheetName, 48, 6, null),
					new ExpectedCellValue(sheetName, 49, 2, "January"),
					new ExpectedCellValue(sheetName, 49, 3, null),
					new ExpectedCellValue(sheetName, 49, 6, null),
					new ExpectedCellValue(sheetName, 50, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 50, 3, .1333),
					new ExpectedCellValue(sheetName, 50, 4, .1333),
					new ExpectedCellValue(sheetName, 50, 5, .0667),
					new ExpectedCellValue(sheetName, 50, 6, .3333),
					new ExpectedCellValue(sheetName, 51, 2, "February"),
					new ExpectedCellValue(sheetName, 51, 4, null),
					new ExpectedCellValue(sheetName, 51, 5, null),
					new ExpectedCellValue(sheetName, 52, 2, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 52, 3, 0),
					new ExpectedCellValue(sheetName, 52, 4, 0),
					new ExpectedCellValue(sheetName, 52, 5, .0667),
					new ExpectedCellValue(sheetName, 52, 6, .0667),
					new ExpectedCellValue(sheetName, 53, 2, "Tent"),
					new ExpectedCellValue(sheetName, 53, 3, 0),
					new ExpectedCellValue(sheetName, 53, 4, .4),
					new ExpectedCellValue(sheetName, 53, 5, 0),
					new ExpectedCellValue(sheetName, 53, 6, .4),
					new ExpectedCellValue(sheetName, 54, 2, "March"),
					new ExpectedCellValue(sheetName, 54, 3, null),
					new ExpectedCellValue(sheetName, 54, 6, null),
					new ExpectedCellValue(sheetName, 55, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 55, 3, 0),
					new ExpectedCellValue(sheetName, 55, 4, .1333),
					new ExpectedCellValue(sheetName, 55, 5, 0),
					new ExpectedCellValue(sheetName, 55, 6, .1333),
					new ExpectedCellValue(sheetName, 56, 2, "Headlamp"),
					new ExpectedCellValue(sheetName, 56, 3, .0667),
					new ExpectedCellValue(sheetName, 56, 4, 0),
					new ExpectedCellValue(sheetName, 56, 5, 0),
					new ExpectedCellValue(sheetName, 56, 6, .0667),
					new ExpectedCellValue(sheetName, 57, 2, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 57, 3, null),
					new ExpectedCellValue(sheetName, 57, 6, null),
					new ExpectedCellValue(sheetName, 58, 2, "January"),
					new ExpectedCellValue(sheetName, 58, 3, null),
					new ExpectedCellValue(sheetName, 58, 4, null),
					new ExpectedCellValue(sheetName, 59, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 59, 3, 415.75),
					new ExpectedCellValue(sheetName, 59, 4, 415.75),
					new ExpectedCellValue(sheetName, 59, 5, 415.75),
					new ExpectedCellValue(sheetName, 59, 6, 1247.25),
					new ExpectedCellValue(sheetName, 60, 2, "February"),
					new ExpectedCellValue(sheetName, 60, 3, null),
					new ExpectedCellValue(sheetName, 60, 5, null),
					new ExpectedCellValue(sheetName, 61, 2, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 61, 3, null),
					new ExpectedCellValue(sheetName, 61, 4, null),
					new ExpectedCellValue(sheetName, 61, 5, 99),
					new ExpectedCellValue(sheetName, 61, 6, 99),
					new ExpectedCellValue(sheetName, 62, 2, "Tent"),
					new ExpectedCellValue(sheetName, 62, 3, null),
					new ExpectedCellValue(sheetName, 62, 4, 199),
					new ExpectedCellValue(sheetName, 62, 5, null),
					new ExpectedCellValue(sheetName, 62, 6, 199),
					new ExpectedCellValue(sheetName, 63, 2, "March"),
					new ExpectedCellValue(sheetName, 63, 5, null),
					new ExpectedCellValue(sheetName, 63, 6, null),
					new ExpectedCellValue(sheetName, 64, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 64, 3, null),
					new ExpectedCellValue(sheetName, 64, 4, 415.75),
					new ExpectedCellValue(sheetName, 64, 5, null),
					new ExpectedCellValue(sheetName, 64, 6, 415.75),
					new ExpectedCellValue(sheetName, 65, 2, "Headlamp"),
					new ExpectedCellValue(sheetName, 65, 3, 24.99),
					new ExpectedCellValue(sheetName, 65, 4, null),
					new ExpectedCellValue(sheetName, 65, 5, null),
					new ExpectedCellValue(sheetName, 65, 6, 24.99),
					new ExpectedCellValue(sheetName, 66, 2, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 66, 3, .2),
					new ExpectedCellValue(sheetName, 66, 4, .6667),
					new ExpectedCellValue(sheetName, 66, 5, .1333),
					new ExpectedCellValue(sheetName, 66, 6, 1),
					new ExpectedCellValue(sheetName, 67, 2, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 67, 3, 440.74),
					new ExpectedCellValue(sheetName, 67, 4, 1030.5),
					new ExpectedCellValue(sheetName, 67, 5, 514.75),
					new ExpectedCellValue(sheetName, 67, 6, 1985.99)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableShowDataAs.xlsx")]
		public void PivotTableRefreshShowDataAsRowFieldsColumnFieldsAndRowDataFieldsGrandTotalsOff()
		{
			var file = new FileInfo("PivotTableShowDataAs.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable4"];
					pivotTable.RowGrandTotals = false;
					pivotTable.ColumnGrandTotals = false;
					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					wholesalePriceDataField.ShowDataAs = ShowDataAs.NoCalculation;
					unitsSoldDataField.ShowDataAs = ShowDataAs.PercentOfTotal;
					foreach (var field in pivotTable.Fields)
					{
						field.SubTotalFunctions = eSubTotalFunctions.None;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B46:E65"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 47, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 47, 3, "Chicago"),
					new ExpectedCellValue(sheetName, 47, 4, "Nashville"),
					new ExpectedCellValue(sheetName, 47, 5, "San Francisco"),
					new ExpectedCellValue(sheetName, 48, 2, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 48, 3, null),
					new ExpectedCellValue(sheetName, 49, 2, "January"),
					new ExpectedCellValue(sheetName, 49, 3, null),
					new ExpectedCellValue(sheetName, 50, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 50, 3, .1333),
					new ExpectedCellValue(sheetName, 50, 4, .1333),
					new ExpectedCellValue(sheetName, 50, 5, .0667),
					new ExpectedCellValue(sheetName, 51, 2, "February"),
					new ExpectedCellValue(sheetName, 51, 4, null),
					new ExpectedCellValue(sheetName, 51, 5, null),
					new ExpectedCellValue(sheetName, 52, 2, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 52, 3, 0),
					new ExpectedCellValue(sheetName, 52, 4, 0),
					new ExpectedCellValue(sheetName, 52, 5, .0667),
					new ExpectedCellValue(sheetName, 53, 2, "Tent"),
					new ExpectedCellValue(sheetName, 53, 3, 0),
					new ExpectedCellValue(sheetName, 53, 4, .4),
					new ExpectedCellValue(sheetName, 53, 5, 0),
					new ExpectedCellValue(sheetName, 54, 2, "March"),
					new ExpectedCellValue(sheetName, 54, 3, null),
					new ExpectedCellValue(sheetName, 55, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 55, 3, 0),
					new ExpectedCellValue(sheetName, 55, 4, .1333),
					new ExpectedCellValue(sheetName, 55, 5, 0),
					new ExpectedCellValue(sheetName, 56, 2, "Headlamp"),
					new ExpectedCellValue(sheetName, 56, 3, .0667),
					new ExpectedCellValue(sheetName, 56, 4, 0),
					new ExpectedCellValue(sheetName, 56, 5, 0),
					new ExpectedCellValue(sheetName, 57, 2, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 57, 3, null),
					new ExpectedCellValue(sheetName, 57, 6, null),
					new ExpectedCellValue(sheetName, 58, 2, "January"),
					new ExpectedCellValue(sheetName, 58, 3, null),
					new ExpectedCellValue(sheetName, 58, 4, null),
					new ExpectedCellValue(sheetName, 59, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 59, 3, 415.75),
					new ExpectedCellValue(sheetName, 59, 4, 415.75),
					new ExpectedCellValue(sheetName, 59, 5, 415.75),
					new ExpectedCellValue(sheetName, 60, 2, "February"),
					new ExpectedCellValue(sheetName, 60, 3, null),
					new ExpectedCellValue(sheetName, 60, 5, null),
					new ExpectedCellValue(sheetName, 61, 2, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 61, 3, null),
					new ExpectedCellValue(sheetName, 61, 4, null),
					new ExpectedCellValue(sheetName, 61, 5, 99),
					new ExpectedCellValue(sheetName, 62, 2, "Tent"),
					new ExpectedCellValue(sheetName, 62, 3, null),
					new ExpectedCellValue(sheetName, 62, 4, 199),
					new ExpectedCellValue(sheetName, 62, 5, null),
					new ExpectedCellValue(sheetName, 63, 2, "March"),
					new ExpectedCellValue(sheetName, 63, 5, null),
					new ExpectedCellValue(sheetName, 64, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 64, 3, null),
					new ExpectedCellValue(sheetName, 64, 4, 415.75),
					new ExpectedCellValue(sheetName, 64, 5, null),
					new ExpectedCellValue(sheetName, 65, 2, "Headlamp"),
					new ExpectedCellValue(sheetName, 65, 3, 24.99),
					new ExpectedCellValue(sheetName, 65, 4, null),
					new ExpectedCellValue(sheetName, 65, 5, null)
				});
			}
		}
		#endregion

		#region PercentOfCol Tests
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableShowDataAs.xlsx")]
		public void PivotTableRefreshShowDataAsPercentOfColRowFieldsColumnFieldsAndRowDataFields()
		{
			var file = new FileInfo("PivotTableShowDataAs.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable4"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					wholesalePriceDataField.ShowDataAs = ShowDataAs.NoCalculation;
					unitsSoldDataField.ShowDataAs = ShowDataAs.PercentOfCol;
					foreach (var field in pivotTable.Fields)
					{
						field.SubTotalFunctions = eSubTotalFunctions.None;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B46:F67"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 47, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 47, 3, "Chicago"),
					new ExpectedCellValue(sheetName, 47, 4, "Nashville"),
					new ExpectedCellValue(sheetName, 47, 5, "San Francisco"),
					new ExpectedCellValue(sheetName, 47, 6, "Grand Total"),
					new ExpectedCellValue(sheetName, 48, 2, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 48, 3, null),
					new ExpectedCellValue(sheetName, 48, 6, null),
					new ExpectedCellValue(sheetName, 49, 2, "January"),
					new ExpectedCellValue(sheetName, 49, 3, null),
					new ExpectedCellValue(sheetName, 49, 6, null),
					new ExpectedCellValue(sheetName, 50, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 50, 3, .6667),
					new ExpectedCellValue(sheetName, 50, 4, .2),
					new ExpectedCellValue(sheetName, 50, 5, .5),
					new ExpectedCellValue(sheetName, 50, 6, .3333),
					new ExpectedCellValue(sheetName, 51, 2, "February"),
					new ExpectedCellValue(sheetName, 51, 4, null),
					new ExpectedCellValue(sheetName, 51, 5, null),
					new ExpectedCellValue(sheetName, 52, 2, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 52, 3, 0),
					new ExpectedCellValue(sheetName, 52, 4, 0),
					new ExpectedCellValue(sheetName, 52, 5, .5),
					new ExpectedCellValue(sheetName, 52, 6, .0667),
					new ExpectedCellValue(sheetName, 53, 2, "Tent"),
					new ExpectedCellValue(sheetName, 53, 3, 0),
					new ExpectedCellValue(sheetName, 53, 4, .6),
					new ExpectedCellValue(sheetName, 53, 5, 0),
					new ExpectedCellValue(sheetName, 53, 6, .4),
					new ExpectedCellValue(sheetName, 54, 2, "March"),
					new ExpectedCellValue(sheetName, 54, 3, null),
					new ExpectedCellValue(sheetName, 54, 6, null),
					new ExpectedCellValue(sheetName, 55, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 55, 3, 0),
					new ExpectedCellValue(sheetName, 55, 4, .2),
					new ExpectedCellValue(sheetName, 55, 5, 0),
					new ExpectedCellValue(sheetName, 55, 6, .1333),
					new ExpectedCellValue(sheetName, 56, 2, "Headlamp"),
					new ExpectedCellValue(sheetName, 56, 3, .3333),
					new ExpectedCellValue(sheetName, 56, 4, 0),
					new ExpectedCellValue(sheetName, 56, 5, 0),
					new ExpectedCellValue(sheetName, 56, 6, .0667),
					new ExpectedCellValue(sheetName, 57, 2, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 57, 3, null),
					new ExpectedCellValue(sheetName, 57, 6, null),
					new ExpectedCellValue(sheetName, 58, 2, "January"),
					new ExpectedCellValue(sheetName, 58, 3, null),
					new ExpectedCellValue(sheetName, 58, 4, null),
					new ExpectedCellValue(sheetName, 59, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 59, 3, 415.75),
					new ExpectedCellValue(sheetName, 59, 4, 415.75),
					new ExpectedCellValue(sheetName, 59, 5, 415.75),
					new ExpectedCellValue(sheetName, 59, 6, 1247.25),
					new ExpectedCellValue(sheetName, 60, 2, "February"),
					new ExpectedCellValue(sheetName, 60, 3, null),
					new ExpectedCellValue(sheetName, 60, 5, null),
					new ExpectedCellValue(sheetName, 61, 2, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 61, 3, null),
					new ExpectedCellValue(sheetName, 61, 4, null),
					new ExpectedCellValue(sheetName, 61, 5, 99),
					new ExpectedCellValue(sheetName, 61, 6, 99),
					new ExpectedCellValue(sheetName, 62, 2, "Tent"),
					new ExpectedCellValue(sheetName, 62, 3, null),
					new ExpectedCellValue(sheetName, 62, 4, 199),
					new ExpectedCellValue(sheetName, 62, 5, null),
					new ExpectedCellValue(sheetName, 62, 6, 199),
					new ExpectedCellValue(sheetName, 63, 2, "March"),
					new ExpectedCellValue(sheetName, 63, 5, null),
					new ExpectedCellValue(sheetName, 63, 6, null),
					new ExpectedCellValue(sheetName, 64, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 64, 3, null),
					new ExpectedCellValue(sheetName, 64, 4, 415.75),
					new ExpectedCellValue(sheetName, 64, 5, null),
					new ExpectedCellValue(sheetName, 64, 6, 415.75),
					new ExpectedCellValue(sheetName, 65, 2, "Headlamp"),
					new ExpectedCellValue(sheetName, 65, 3, 24.99),
					new ExpectedCellValue(sheetName, 65, 4, null),
					new ExpectedCellValue(sheetName, 65, 5, null),
					new ExpectedCellValue(sheetName, 65, 6, 24.99),
					new ExpectedCellValue(sheetName, 66, 2, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 66, 3, 1),
					new ExpectedCellValue(sheetName, 66, 4, 1),
					new ExpectedCellValue(sheetName, 66, 5, 1),
					new ExpectedCellValue(sheetName, 66, 6, 1),
					new ExpectedCellValue(sheetName, 67, 2, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 67, 3, 440.74),
					new ExpectedCellValue(sheetName, 67, 4, 1030.5),
					new ExpectedCellValue(sheetName, 67, 5, 514.75),
					new ExpectedCellValue(sheetName, 67, 6, 1985.99)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableShowDataAs.xlsx")]
		public void PivotTableRefreshShowDataAsPercentOfColRowFieldsColumnFieldsAndColumnDataFields()
		{
			var file = new FileInfo("PivotTableShowDataAs.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable3"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					wholesalePriceDataField.ShowDataAs = ShowDataAs.PercentOfCol;
					unitsSoldDataField.ShowDataAs = ShowDataAs.NoCalculation;
					foreach (var field in pivotTable.Fields)
					{
						field.SubTotalFunctions = eSubTotalFunctions.None;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B30:J41"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 31, 2, null),
					new ExpectedCellValue(sheetName, 31, 3, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 31, 4, null),
					new ExpectedCellValue(sheetName, 31, 5, null),
					new ExpectedCellValue(sheetName, 31, 6, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 31, 7, null),
					new ExpectedCellValue(sheetName, 31, 8, null),
					new ExpectedCellValue(sheetName, 31, 9, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 31, 10, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 32, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 32, 3, "Chicago"),
					new ExpectedCellValue(sheetName, 32, 4, "Nashville"),
					new ExpectedCellValue(sheetName, 32, 5, "San Francisco"),
					new ExpectedCellValue(sheetName, 32, 6, "Chicago"),
					new ExpectedCellValue(sheetName, 32, 7, "Nashville"),
					new ExpectedCellValue(sheetName, 32, 8, "San Francisco"),
					new ExpectedCellValue(sheetName, 32, 9, null),
					new ExpectedCellValue(sheetName, 32, 10, null),
					new ExpectedCellValue(sheetName, 33, 2, "January"),
					new ExpectedCellValue(sheetName, 33, 3, null),
					new ExpectedCellValue(sheetName, 33, 4, null),
					new ExpectedCellValue(sheetName, 33, 5, null),
					new ExpectedCellValue(sheetName, 33, 10, null),
					new ExpectedCellValue(sheetName, 34, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 34, 3, .9433),
					new ExpectedCellValue(sheetName, 34, 4, .4034),
					new ExpectedCellValue(sheetName, 34, 5, .8077),
					new ExpectedCellValue(sheetName, 34, 6, 2),
					new ExpectedCellValue(sheetName, 34, 7, 2),
					new ExpectedCellValue(sheetName, 34, 8, 1),
					new ExpectedCellValue(sheetName, 34, 9, .628),
					new ExpectedCellValue(sheetName, 34, 10, 5),
					new ExpectedCellValue(sheetName, 35, 2, "February"),
					new ExpectedCellValue(sheetName, 35, 3, null),
					new ExpectedCellValue(sheetName, 35, 6, null),
					new ExpectedCellValue(sheetName, 35, 10, null),
					new ExpectedCellValue(sheetName, 36, 2, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 36, 3, 0),
					new ExpectedCellValue(sheetName, 36, 4, 0),
					new ExpectedCellValue(sheetName, 36, 5, .1923),
					new ExpectedCellValue(sheetName, 36, 6, null),
					new ExpectedCellValue(sheetName, 36, 7, null),
					new ExpectedCellValue(sheetName, 36, 8, 1),
					new ExpectedCellValue(sheetName, 36, 9, .0498),
					new ExpectedCellValue(sheetName, 36, 10, 1),
					new ExpectedCellValue(sheetName, 37, 2, "Tent"),
					new ExpectedCellValue(sheetName, 37, 3, 0),
					new ExpectedCellValue(sheetName, 37, 4, .1931),
					new ExpectedCellValue(sheetName, 37, 5, 0),
					new ExpectedCellValue(sheetName, 37, 6, null),
					new ExpectedCellValue(sheetName, 37, 7, 6),
					new ExpectedCellValue(sheetName, 37, 8, null),
					new ExpectedCellValue(sheetName, 37, 9, .1002),
					new ExpectedCellValue(sheetName, 37, 10, 6),
					new ExpectedCellValue(sheetName, 38, 2, "March"),
					new ExpectedCellValue(sheetName, 38, 3, null),
					new ExpectedCellValue(sheetName, 38, 5, null),
					new ExpectedCellValue(sheetName, 38, 7, null),
					new ExpectedCellValue(sheetName, 38, 9, null),
					new ExpectedCellValue(sheetName, 38, 10, null),
					new ExpectedCellValue(sheetName, 39, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 39, 3, 0),
					new ExpectedCellValue(sheetName, 39, 4, .4034),
					new ExpectedCellValue(sheetName, 39, 5, 0),
					new ExpectedCellValue(sheetName, 39, 6, null),
					new ExpectedCellValue(sheetName, 39, 7, 2),
					new ExpectedCellValue(sheetName, 39, 8, null),
					new ExpectedCellValue(sheetName, 39, 9, .2093),
					new ExpectedCellValue(sheetName, 39, 10, 2),
					new ExpectedCellValue(sheetName, 40, 2, "Headlamp"),
					new ExpectedCellValue(sheetName, 40, 3, .0567),
					new ExpectedCellValue(sheetName, 40, 4, 0),
					new ExpectedCellValue(sheetName, 40, 5, 0),
					new ExpectedCellValue(sheetName, 40, 6, 1),
					new ExpectedCellValue(sheetName, 40, 7, null),
					new ExpectedCellValue(sheetName, 40, 8, null),
					new ExpectedCellValue(sheetName, 40, 9, .0126),
					new ExpectedCellValue(sheetName, 40, 10, 1),
					new ExpectedCellValue(sheetName, 41, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 41, 3, 1),
					new ExpectedCellValue(sheetName, 41, 4, 1),
					new ExpectedCellValue(sheetName, 41, 5, 1),
					new ExpectedCellValue(sheetName, 41, 6, 3),
					new ExpectedCellValue(sheetName, 41, 7, 10),
					new ExpectedCellValue(sheetName, 41, 8, 2),
					new ExpectedCellValue(sheetName, 41, 9, 1),
					new ExpectedCellValue(sheetName, 41, 10, 15)
				});
			}
		}
		#endregion

		#region PercentOfRow Tests
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableShowDataAs.xlsx")]
		public void PivotTableRefreshShowDataAsPercentOfRowRowFieldsColumnFieldsAndRowDataFields()
		{
			var file = new FileInfo("PivotTableShowDataAs.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable4"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					wholesalePriceDataField.ShowDataAs = ShowDataAs.NoCalculation;
					unitsSoldDataField.ShowDataAs = ShowDataAs.PercentOfRow;
					foreach (var field in pivotTable.Fields)
					{
						field.SubTotalFunctions = eSubTotalFunctions.None;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B46:F67"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 47, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 47, 3, "Chicago"),
					new ExpectedCellValue(sheetName, 47, 4, "Nashville"),
					new ExpectedCellValue(sheetName, 47, 5, "San Francisco"),
					new ExpectedCellValue(sheetName, 47, 6, "Grand Total"),
					new ExpectedCellValue(sheetName, 48, 2, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 48, 3, null),
					new ExpectedCellValue(sheetName, 48, 6, null),
					new ExpectedCellValue(sheetName, 49, 2, "January"),
					new ExpectedCellValue(sheetName, 49, 3, null),
					new ExpectedCellValue(sheetName, 49, 6, null),
					new ExpectedCellValue(sheetName, 50, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 50, 3, .4),
					new ExpectedCellValue(sheetName, 50, 4, .4),
					new ExpectedCellValue(sheetName, 50, 5, .2),
					new ExpectedCellValue(sheetName, 50, 6, 1),
					new ExpectedCellValue(sheetName, 51, 2, "February"),
					new ExpectedCellValue(sheetName, 51, 4, null),
					new ExpectedCellValue(sheetName, 51, 5, null),
					new ExpectedCellValue(sheetName, 52, 2, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 52, 3, 0),
					new ExpectedCellValue(sheetName, 52, 4, 0),
					new ExpectedCellValue(sheetName, 52, 5, 1),
					new ExpectedCellValue(sheetName, 52, 6, 1),
					new ExpectedCellValue(sheetName, 53, 2, "Tent"),
					new ExpectedCellValue(sheetName, 53, 3, 0),
					new ExpectedCellValue(sheetName, 53, 4, 1),
					new ExpectedCellValue(sheetName, 53, 5, 0),
					new ExpectedCellValue(sheetName, 53, 6, 1),
					new ExpectedCellValue(sheetName, 54, 2, "March"),
					new ExpectedCellValue(sheetName, 54, 3, null),
					new ExpectedCellValue(sheetName, 54, 6, null),
					new ExpectedCellValue(sheetName, 55, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 55, 3, 0),
					new ExpectedCellValue(sheetName, 55, 4, 1),
					new ExpectedCellValue(sheetName, 55, 5, 0),
					new ExpectedCellValue(sheetName, 55, 6, 1),
					new ExpectedCellValue(sheetName, 56, 2, "Headlamp"),
					new ExpectedCellValue(sheetName, 56, 3, 1),
					new ExpectedCellValue(sheetName, 56, 4, 0),
					new ExpectedCellValue(sheetName, 56, 5, 0),
					new ExpectedCellValue(sheetName, 56, 6, 1),
					new ExpectedCellValue(sheetName, 57, 2, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 57, 3, null),
					new ExpectedCellValue(sheetName, 57, 6, null),
					new ExpectedCellValue(sheetName, 58, 2, "January"),
					new ExpectedCellValue(sheetName, 58, 3, null),
					new ExpectedCellValue(sheetName, 58, 4, null),
					new ExpectedCellValue(sheetName, 59, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 59, 3, 415.75),
					new ExpectedCellValue(sheetName, 59, 4, 415.75),
					new ExpectedCellValue(sheetName, 59, 5, 415.75),
					new ExpectedCellValue(sheetName, 59, 6, 1247.25),
					new ExpectedCellValue(sheetName, 60, 2, "February"),
					new ExpectedCellValue(sheetName, 60, 3, null),
					new ExpectedCellValue(sheetName, 60, 5, null),
					new ExpectedCellValue(sheetName, 61, 2, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 61, 3, null),
					new ExpectedCellValue(sheetName, 61, 4, null),
					new ExpectedCellValue(sheetName, 61, 5, 99),
					new ExpectedCellValue(sheetName, 61, 6, 99),
					new ExpectedCellValue(sheetName, 62, 2, "Tent"),
					new ExpectedCellValue(sheetName, 62, 3, null),
					new ExpectedCellValue(sheetName, 62, 4, 199),
					new ExpectedCellValue(sheetName, 62, 5, null),
					new ExpectedCellValue(sheetName, 62, 6, 199),
					new ExpectedCellValue(sheetName, 63, 2, "March"),
					new ExpectedCellValue(sheetName, 63, 5, null),
					new ExpectedCellValue(sheetName, 63, 6, null),
					new ExpectedCellValue(sheetName, 64, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 64, 3, null),
					new ExpectedCellValue(sheetName, 64, 4, 415.75),
					new ExpectedCellValue(sheetName, 64, 5, null),
					new ExpectedCellValue(sheetName, 64, 6, 415.75),
					new ExpectedCellValue(sheetName, 65, 2, "Headlamp"),
					new ExpectedCellValue(sheetName, 65, 3, 24.99),
					new ExpectedCellValue(sheetName, 65, 4, null),
					new ExpectedCellValue(sheetName, 65, 5, null),
					new ExpectedCellValue(sheetName, 65, 6, 24.99),
					new ExpectedCellValue(sheetName, 66, 2, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 66, 3, .2),
					new ExpectedCellValue(sheetName, 66, 4, .6667),
					new ExpectedCellValue(sheetName, 66, 5, .1333),
					new ExpectedCellValue(sheetName, 66, 6, 1),
					new ExpectedCellValue(sheetName, 67, 2, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 67, 3, 440.74),
					new ExpectedCellValue(sheetName, 67, 4, 1030.5),
					new ExpectedCellValue(sheetName, 67, 5, 514.75),
					new ExpectedCellValue(sheetName, 67, 6, 1985.99)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableShowDataAs.xlsx")]
		public void PivotTableRefreshShowDataAsPercentOfRowWithRowFieldsColumnFieldsAndColumnDataFields()
		{
			var file = new FileInfo("PivotTableShowDataAs.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable3"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					wholesalePriceDataField.ShowDataAs = ShowDataAs.PercentOfRow;
					unitsSoldDataField.ShowDataAs = ShowDataAs.NoCalculation;
					foreach (var field in pivotTable.Fields)
					{
						field.SubTotalFunctions = eSubTotalFunctions.None;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B30:J41"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 31, 2, null),
					new ExpectedCellValue(sheetName, 31, 3, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 31, 4, null),
					new ExpectedCellValue(sheetName, 31, 5, null),
					new ExpectedCellValue(sheetName, 31, 6, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 31, 7, null),
					new ExpectedCellValue(sheetName, 31, 8, null),
					new ExpectedCellValue(sheetName, 31, 9, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 31, 10, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 32, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 32, 3, "Chicago"),
					new ExpectedCellValue(sheetName, 32, 4, "Nashville"),
					new ExpectedCellValue(sheetName, 32, 5, "San Francisco"),
					new ExpectedCellValue(sheetName, 32, 6, "Chicago"),
					new ExpectedCellValue(sheetName, 32, 7, "Nashville"),
					new ExpectedCellValue(sheetName, 32, 8, "San Francisco"),
					new ExpectedCellValue(sheetName, 32, 9, null),
					new ExpectedCellValue(sheetName, 32, 10, null),
					new ExpectedCellValue(sheetName, 33, 2, "January"),
					new ExpectedCellValue(sheetName, 33, 3, null),
					new ExpectedCellValue(sheetName, 33, 5, null),
					new ExpectedCellValue(sheetName, 33, 9, null),
					new ExpectedCellValue(sheetName, 33, 10, null),
					new ExpectedCellValue(sheetName, 34, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 34, 3, .3333),
					new ExpectedCellValue(sheetName, 34, 4, .3333),
					new ExpectedCellValue(sheetName, 34, 5, .3333),
					new ExpectedCellValue(sheetName, 34, 6, 2),
					new ExpectedCellValue(sheetName, 34, 7, 2),
					new ExpectedCellValue(sheetName, 34, 8, 1),
					new ExpectedCellValue(sheetName, 34, 9, 1),
					new ExpectedCellValue(sheetName, 34, 10, 5),
					new ExpectedCellValue(sheetName, 35, 2, "February"),
					new ExpectedCellValue(sheetName, 35, 3, null),
					new ExpectedCellValue(sheetName, 35, 6, null),
					new ExpectedCellValue(sheetName, 35, 10, null),
					new ExpectedCellValue(sheetName, 36, 2, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 36, 3, 0),
					new ExpectedCellValue(sheetName, 36, 4, 0),
					new ExpectedCellValue(sheetName, 36, 5, 1),
					new ExpectedCellValue(sheetName, 36, 6, null),
					new ExpectedCellValue(sheetName, 36, 7, null),
					new ExpectedCellValue(sheetName, 36, 8, 1),
					new ExpectedCellValue(sheetName, 36, 9, 1),
					new ExpectedCellValue(sheetName, 36, 10, 1),
					new ExpectedCellValue(sheetName, 37, 2, "Tent"),
					new ExpectedCellValue(sheetName, 37, 3, 0),
					new ExpectedCellValue(sheetName, 37, 4, 1),
					new ExpectedCellValue(sheetName, 37, 5, 0),
					new ExpectedCellValue(sheetName, 37, 6, null),
					new ExpectedCellValue(sheetName, 37, 7, 6),
					new ExpectedCellValue(sheetName, 37, 8, null),
					new ExpectedCellValue(sheetName, 37, 9, 1),
					new ExpectedCellValue(sheetName, 37, 10, 6),
					new ExpectedCellValue(sheetName, 38, 2, "March"),
					new ExpectedCellValue(sheetName, 38, 3, null),
					new ExpectedCellValue(sheetName, 38, 5, null),
					new ExpectedCellValue(sheetName, 38, 7, null),
					new ExpectedCellValue(sheetName, 38, 9, null),
					new ExpectedCellValue(sheetName, 38, 10, null),
					new ExpectedCellValue(sheetName, 39, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 39, 3, 0),
					new ExpectedCellValue(sheetName, 39, 4, 1),
					new ExpectedCellValue(sheetName, 39, 5, 0),
					new ExpectedCellValue(sheetName, 39, 6, null),
					new ExpectedCellValue(sheetName, 39, 7, 2),
					new ExpectedCellValue(sheetName, 39, 8, null),
					new ExpectedCellValue(sheetName, 39, 9, 1),
					new ExpectedCellValue(sheetName, 39, 10, 2),
					new ExpectedCellValue(sheetName, 40, 2, "Headlamp"),
					new ExpectedCellValue(sheetName, 40, 3, 1),
					new ExpectedCellValue(sheetName, 40, 4, 0),
					new ExpectedCellValue(sheetName, 40, 5, 0),
					new ExpectedCellValue(sheetName, 40, 6, 1),
					new ExpectedCellValue(sheetName, 40, 7, null),
					new ExpectedCellValue(sheetName, 40, 8, null),
					new ExpectedCellValue(sheetName, 40, 9, 1),
					new ExpectedCellValue(sheetName, 40, 10, 1),
					new ExpectedCellValue(sheetName, 41, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 41, 3, .2219),
					new ExpectedCellValue(sheetName, 41, 4, .5189),
					new ExpectedCellValue(sheetName, 41, 5, .2592),
					new ExpectedCellValue(sheetName, 41, 6, 3),
					new ExpectedCellValue(sheetName, 41, 7, 10),
					new ExpectedCellValue(sheetName, 41, 8, 2),
					new ExpectedCellValue(sheetName, 41, 9, 1),
					new ExpectedCellValue(sheetName, 41, 10, 15)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\ShowDataAsComplex.xlsx")]
		public void PivotTableRefreshShowDataAsPercentOfRowTotalWithErrorValues()
		{
			var file = new FileInfo("ShowDataAsComplex.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "Sheet1";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable2"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B21:E31"), pivotTable.Address);
					Assert.AreEqual(15, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 21, 2, "Sum of Net Change"),
					new ExpectedCellValue(sheetName, 21, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 22, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 22, 3, "FALSE"),
					new ExpectedCellValue(sheetName, 22, 4, "TRUE"),
					new ExpectedCellValue(sheetName, 22, 5, "Grand Total"),
					new ExpectedCellValue(sheetName, 23, 2, "Accounts Receivable"),
					new ExpectedCellValue(sheetName, 23, 3, ExcelErrorValue.Create(eErrorType.Div0)),
					new ExpectedCellValue(sheetName, 23, 4, ExcelErrorValue.Create(eErrorType.Div0)),
					new ExpectedCellValue(sheetName, 23, 5, ExcelErrorValue.Create(eErrorType.Div0)),
					new ExpectedCellValue(sheetName, 24, 2, "Accounts Receivable, Total"),
					new ExpectedCellValue(sheetName, 24, 3, 0),
					new ExpectedCellValue(sheetName, 24, 4, 1),
					new ExpectedCellValue(sheetName, 24, 5, 1),
					new ExpectedCellValue(sheetName, 25, 2, "ASSETS"),
					new ExpectedCellValue(sheetName, 25, 3, ExcelErrorValue.Create(eErrorType.Div0)),
					new ExpectedCellValue(sheetName, 25, 4, ExcelErrorValue.Create(eErrorType.Div0)),
					new ExpectedCellValue(sheetName, 25, 5, ExcelErrorValue.Create(eErrorType.Div0)),
					new ExpectedCellValue(sheetName, 26, 2, "Cash"),
					new ExpectedCellValue(sheetName, 26, 3, ExcelErrorValue.Create(eErrorType.Div0)),
					new ExpectedCellValue(sheetName, 26, 4, ExcelErrorValue.Create(eErrorType.Div0)),
					new ExpectedCellValue(sheetName, 26, 5, ExcelErrorValue.Create(eErrorType.Div0)),
					new ExpectedCellValue(sheetName, 27, 2, "Customers, EU"),
					new ExpectedCellValue(sheetName, 27, 3, 1),
					new ExpectedCellValue(sheetName, 27, 4, 0),
					new ExpectedCellValue(sheetName, 27, 5, 1),
					new ExpectedCellValue(sheetName, 28, 2, "Customers, North America"),
					new ExpectedCellValue(sheetName, 28, 3, 0),
					new ExpectedCellValue(sheetName, 28, 4, 1),
					new ExpectedCellValue(sheetName, 28, 5, 1),
					new ExpectedCellValue(sheetName, 29, 2, "Liquid Assets, Total"),
					new ExpectedCellValue(sheetName, 29, 3, ExcelErrorValue.Create(eErrorType.Div0)),
					new ExpectedCellValue(sheetName, 29, 4, ExcelErrorValue.Create(eErrorType.Div0)),
					new ExpectedCellValue(sheetName, 29, 5, ExcelErrorValue.Create(eErrorType.Div0)),
					new ExpectedCellValue(sheetName, 30, 2, "Securities, Total"),
					new ExpectedCellValue(sheetName, 30, 3, ExcelErrorValue.Create(eErrorType.Div0)),
					new ExpectedCellValue(sheetName, 30, 4, ExcelErrorValue.Create(eErrorType.Div0)),
					new ExpectedCellValue(sheetName, 30, 5, ExcelErrorValue.Create(eErrorType.Div0)),
					new ExpectedCellValue(sheetName, 31, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 31, 3, 0.0770052210411582),
					new ExpectedCellValue(sheetName, 31, 4, 0.922994778958842),
					new ExpectedCellValue(sheetName, 31, 5, 1),
				});
			}
		}
		#endregion

		#region PercentOf Tests
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableShowDataAs.xlsx")]
		public void PivotTableRefreshShowDataAsRowFieldsColumnDataFieldsPercentOfMonthMarch()
		{
			var file = new FileInfo("PivotTableShowDataAs.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				void validateSheet() => TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 2, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 2, 3, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 2, 4, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 3, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 4, 2, "January"),
					new ExpectedCellValue(sheetName, 4, 3, 16.6367),
					new ExpectedCellValue(sheetName, 4, 4, 2),
					new ExpectedCellValue(sheetName, 5, 2, "March"),
					new ExpectedCellValue(sheetName, 5, 3, 1),
					new ExpectedCellValue(sheetName, 5, 4, 1),
					new ExpectedCellValue(sheetName, 6, 2, "Nashville"),
					new ExpectedCellValue(sheetName, 7, 2, "January"),
					new ExpectedCellValue(sheetName, 7, 3, 1),
					new ExpectedCellValue(sheetName, 7, 4, 2),
					new ExpectedCellValue(sheetName, 8, 2, "February"),
					new ExpectedCellValue(sheetName, 8, 3, .4787),
					new ExpectedCellValue(sheetName, 8, 4, 6),
					new ExpectedCellValue(sheetName, 9, 2, "March"),
					new ExpectedCellValue(sheetName, 9, 3, 1),
					new ExpectedCellValue(sheetName, 9, 4, 2),
					new ExpectedCellValue(sheetName, 10, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 11, 2, "January"),
					new ExpectedCellValue(sheetName, 11, 3, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 11, 4, 1),
					new ExpectedCellValue(sheetName, 12, 2, "February"),
					new ExpectedCellValue(sheetName, 12, 3, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 12, 4, 1),
					new ExpectedCellValue(sheetName, 13, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 13, 3, null),
					new ExpectedCellValue(sheetName, 13, 4, 15)
				});

				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable1"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					// Show 'Wholesale Price' data as the percentage of Month: March values.
					wholesalePriceDataField.ShowDataAs = ShowDataAs.Percent;
					wholesalePriceDataField.BaseField = pivotTable.CacheDefinition.GetCacheFieldIndex("Month");
					wholesalePriceDataField.BaseItem = this.GetCacheFieldSharedItemIndex(pivotTable, wholesalePriceDataField.BaseField, "March");
					unitsSoldDataField.ShowDataAs = ShowDataAs.NoCalculation;
					foreach (var field in pivotTable.Fields)
					{
						field.SubtotalLocation = SubtotalLocation.Top;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B2:D13"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				validateSheet();
				// Validate that top subtotals were written in.
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 3, 3, null),
					new ExpectedCellValue(sheetName, 3, 4, 3),
					new ExpectedCellValue(sheetName, 6, 3, null),
					new ExpectedCellValue(sheetName, 6, 4, 10),
					new ExpectedCellValue(sheetName, 10, 3, null),
					new ExpectedCellValue(sheetName, 10, 4, 2),
				});

				// Test again with subtotals turned off.
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable1"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					// Show 'Wholesale Price' data as the percentage of Month: January values.
					wholesalePriceDataField.ShowDataAs = ShowDataAs.Percent;
					wholesalePriceDataField.BaseField = pivotTable.CacheDefinition.GetCacheFieldIndex("Month");
					wholesalePriceDataField.BaseItem = this.GetCacheFieldSharedItemIndex(pivotTable, wholesalePriceDataField.BaseField, "March");
					unitsSoldDataField.ShowDataAs = ShowDataAs.NoCalculation;
					foreach (var field in pivotTable.Fields)
					{
						field.SubTotalFunctions = eSubTotalFunctions.None;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B2:D13"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				validateSheet();
				// Validate that no subtotals were written in.
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 3, 3, null),
					new ExpectedCellValue(sheetName, 3, 4, null),
					new ExpectedCellValue(sheetName, 6, 3, null),
					new ExpectedCellValue(sheetName, 6, 4, null),
					new ExpectedCellValue(sheetName, 10, 3, null),
					new ExpectedCellValue(sheetName, 10, 4, null),
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableShowDataAs.xlsx")]
		public void PivotTableRefreshShowDataAsRowFieldsColumnDataFieldsPercentOfLocationSanFrancisco()
		{
			var file = new FileInfo("PivotTableShowDataAs.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				void validateSheet() => TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 2, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 2, 3, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 2, 4, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 3, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 4, 2, "January"),
					new ExpectedCellValue(sheetName, 4, 3, 1),
					new ExpectedCellValue(sheetName, 4, 4, 2),
					new ExpectedCellValue(sheetName, 5, 2, "March"),
					new ExpectedCellValue(sheetName, 5, 3, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 5, 4, 1),
					new ExpectedCellValue(sheetName, 6, 2, "Nashville"),
					new ExpectedCellValue(sheetName, 7, 2, "January"),
					new ExpectedCellValue(sheetName, 7, 3, 1),
					new ExpectedCellValue(sheetName, 7, 4, 2),
					new ExpectedCellValue(sheetName, 8, 2, "February"),
					new ExpectedCellValue(sheetName, 8, 3, 2.0101),
					new ExpectedCellValue(sheetName, 8, 4, 6),
					new ExpectedCellValue(sheetName, 9, 2, "March"),
					new ExpectedCellValue(sheetName, 9, 3, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 9, 4, 2),
					new ExpectedCellValue(sheetName, 10, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 11, 2, "January"),
					new ExpectedCellValue(sheetName, 11, 3, 1),
					new ExpectedCellValue(sheetName, 11, 4, 1),
					new ExpectedCellValue(sheetName, 12, 2, "February"),
					new ExpectedCellValue(sheetName, 12, 3, 1),
					new ExpectedCellValue(sheetName, 12, 4, 1),
					new ExpectedCellValue(sheetName, 13, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 13, 3, null),
					new ExpectedCellValue(sheetName, 13, 4, 15)
				});

				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable1"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					// Show 'Wholesale Price' data as the percentage of Location: San Francisco values.
					wholesalePriceDataField.ShowDataAs = ShowDataAs.Percent;
					wholesalePriceDataField.BaseField = pivotTable.CacheDefinition.GetCacheFieldIndex("Location");
					wholesalePriceDataField.BaseItem = this.GetCacheFieldSharedItemIndex(pivotTable, wholesalePriceDataField.BaseField, "San Francisco");
					unitsSoldDataField.ShowDataAs = ShowDataAs.NoCalculation;
					foreach (var field in pivotTable.Fields)
					{
						field.SubtotalLocation = SubtotalLocation.Top;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B2:D13"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				validateSheet();
				// Validate that top subtotals were written in.
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 3, 3, .8562),
					new ExpectedCellValue(sheetName, 3, 4, 3),
					new ExpectedCellValue(sheetName, 6, 3, 2.0019),
					new ExpectedCellValue(sheetName, 6, 4, 10),
					new ExpectedCellValue(sheetName, 10, 3, 1),
					new ExpectedCellValue(sheetName, 10, 4, 2),
				});

				// Test again with subtotals turned off.
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable1"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					// Show 'Wholesale Price' data as the percentage of Location: San Francisco values.
					wholesalePriceDataField.ShowDataAs = ShowDataAs.Percent;
					wholesalePriceDataField.BaseField = pivotTable.CacheDefinition.GetCacheFieldIndex("Location");
					wholesalePriceDataField.BaseItem = this.GetCacheFieldSharedItemIndex(pivotTable, wholesalePriceDataField.BaseField, "San Francisco");
					unitsSoldDataField.ShowDataAs = ShowDataAs.NoCalculation;
					foreach (var field in pivotTable.Fields)
					{
						field.SubTotalFunctions = eSubTotalFunctions.None;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B2:D13"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				validateSheet();
				// Validate that no subtotals were written in.
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 3, 3, null),
					new ExpectedCellValue(sheetName, 3, 4, null),
					new ExpectedCellValue(sheetName, 6, 3, null),
					new ExpectedCellValue(sheetName, 6, 4, null),
					new ExpectedCellValue(sheetName, 10, 3, null),
					new ExpectedCellValue(sheetName, 10, 4, null),
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableShowDataAs.xlsx")]
		public void PivotTableRefreshShowDataAsColumnFieldsRowDataFieldsPercentOfItemCarRackSubtotalOn()
		{
			var file = new FileInfo("PivotTableShowDataAs.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable2"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					// Show 'Wholesale Price' data as the percentage of Location: San Francisco values.
					wholesalePriceDataField.ShowDataAs = ShowDataAs.Percent;
					wholesalePriceDataField.BaseField = pivotTable.CacheDefinition.GetCacheFieldIndex("Item");
					wholesalePriceDataField.BaseItem = this.GetCacheFieldSharedItemIndex(pivotTable, wholesalePriceDataField.BaseField, "Car Rack");
					unitsSoldDataField.ShowDataAs = ShowDataAs.NoCalculation;
					foreach (var field in pivotTable.Fields)
					{
						field.SubtotalLocation = SubtotalLocation.Top;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B18:K22"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 21, 2, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 21, 3, 1),
					new ExpectedCellValue(sheetName, 21, 4, null),
					new ExpectedCellValue(sheetName, 21, 5, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 21, 6, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 21, 7, null),
					new ExpectedCellValue(sheetName, 21, 8, 1),
					new ExpectedCellValue(sheetName, 21, 9, .0601),
					new ExpectedCellValue(sheetName, 21, 10, null),
					new ExpectedCellValue(sheetName, 21, 11, null),

					new ExpectedCellValue(sheetName, 22, 2, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 22, 3, 5),
					new ExpectedCellValue(sheetName, 22, 4, 5),
					new ExpectedCellValue(sheetName, 22, 5, 1),
					new ExpectedCellValue(sheetName, 22, 6, 6),
					new ExpectedCellValue(sheetName, 22, 7, 7),
					new ExpectedCellValue(sheetName, 22, 8, 2),
					new ExpectedCellValue(sheetName, 22, 9, 1),
					new ExpectedCellValue(sheetName, 22, 10, 3),
					new ExpectedCellValue(sheetName, 22, 11, 15)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableShowDataAs.xlsx")]
		public void PivotTableRefreshShowDataAsColumnFieldsRowDataFieldsPercentOfItemCarRackSubtotalOff()
		{
			var file = new FileInfo("PivotTableShowDataAs.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable2"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					// Show 'Wholesale Price' data as the percentage of Location: San Francisco values.
					wholesalePriceDataField.ShowDataAs = ShowDataAs.Percent;
					wholesalePriceDataField.BaseField = pivotTable.CacheDefinition.GetCacheFieldIndex("Item");
					wholesalePriceDataField.BaseItem = this.GetCacheFieldSharedItemIndex(pivotTable, wholesalePriceDataField.BaseField, "Car Rack");
					unitsSoldDataField.ShowDataAs = ShowDataAs.NoCalculation;
					foreach (var field in pivotTable.Fields)
					{
						field.SubTotalFunctions = eSubTotalFunctions.None;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B18:H22"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 21, 2, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 21, 3, 1),
					new ExpectedCellValue(sheetName, 21, 4, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 21, 5, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 21, 6, 1),
					new ExpectedCellValue(sheetName, 21, 7, .0601),
					new ExpectedCellValue(sheetName, 21, 8, null),

					new ExpectedCellValue(sheetName, 22, 2, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 22, 3, 5),
					new ExpectedCellValue(sheetName, 22, 4, 1),
					new ExpectedCellValue(sheetName, 22, 5, 6),
					new ExpectedCellValue(sheetName, 22, 6, 2),
					new ExpectedCellValue(sheetName, 22, 7, 1),
					new ExpectedCellValue(sheetName, 22, 8, 15)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableShowDataAs.xlsx")]
		public void PivotTableRefreshShowDataAsColumnFieldsRowDataFieldsPercentOfMonthFebruaryRackSubtotalOn()
		{
			var file = new FileInfo("PivotTableShowDataAs.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable2"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					// Show 'Wholesale Price' data as the percentage of Month: February values.
					wholesalePriceDataField.ShowDataAs = ShowDataAs.Percent;
					wholesalePriceDataField.BaseField = pivotTable.CacheDefinition.GetCacheFieldIndex("Month");
					wholesalePriceDataField.BaseItem = this.GetCacheFieldSharedItemIndex(pivotTable, wholesalePriceDataField.BaseField, "February");
					unitsSoldDataField.ShowDataAs = ShowDataAs.NoCalculation;
					foreach (var field in pivotTable.Fields)
					{
						field.SubtotalLocation = SubtotalLocation.Top;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B18:K22"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 21, 2, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 21, 3, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 21, 4, 4.1854),
					new ExpectedCellValue(sheetName, 21, 5, 1),
					new ExpectedCellValue(sheetName, 21, 6, 1),
					new ExpectedCellValue(sheetName, 21, 7, 1),
					new ExpectedCellValue(sheetName, 21, 8, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 21, 9, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 21, 10, 1.4790),
					new ExpectedCellValue(sheetName, 21, 11, null),

					new ExpectedCellValue(sheetName, 22, 2, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 22, 3, 5),
					new ExpectedCellValue(sheetName, 22, 4, 5),
					new ExpectedCellValue(sheetName, 22, 5, 1),
					new ExpectedCellValue(sheetName, 22, 6, 6),
					new ExpectedCellValue(sheetName, 22, 7, 7),
					new ExpectedCellValue(sheetName, 22, 8, 2),
					new ExpectedCellValue(sheetName, 22, 9, 1),
					new ExpectedCellValue(sheetName, 22, 10, 3),
					new ExpectedCellValue(sheetName, 22, 11, 15)
				});
			}
		}


		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableShowDataAs.xlsx")]
		public void PivotTableRefreshShowDataAsColumnFieldsRowDataFieldsPercentOfMonthFebruarySubtotalOff()
		{
			var file = new FileInfo("PivotTableShowDataAs.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable2"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					// Show 'Wholesale Price' data as the percentage of Month: February values.
					wholesalePriceDataField.ShowDataAs = ShowDataAs.Percent;
					wholesalePriceDataField.BaseField = pivotTable.CacheDefinition.GetCacheFieldIndex("Month");
					wholesalePriceDataField.BaseItem = this.GetCacheFieldSharedItemIndex(pivotTable, wholesalePriceDataField.BaseField, "February");
					unitsSoldDataField.ShowDataAs = ShowDataAs.NoCalculation;
					foreach (var field in pivotTable.Fields)
					{
						field.SubTotalFunctions = eSubTotalFunctions.None;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B18:H22"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 21, 2, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 21, 3, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 21, 4, 1),
					new ExpectedCellValue(sheetName, 21, 5, 1),
					new ExpectedCellValue(sheetName, 21, 6, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 21, 7, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 21, 8, null),

					new ExpectedCellValue(sheetName, 22, 2, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 22, 3, 5),
					new ExpectedCellValue(sheetName, 22, 4, 1),
					new ExpectedCellValue(sheetName, 22, 5, 6),
					new ExpectedCellValue(sheetName, 22, 6, 2),
					new ExpectedCellValue(sheetName, 22, 7, 1),
					new ExpectedCellValue(sheetName, 22, 8, 15)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableShowDataAs.xlsx")]
		public void PivotTableRefreshShowDataAsColumnFieldsRowFieldsAndColumnDataFieldsPercentOfMonthJanuarySubtotalTopOnOff()
		{
			var file = new FileInfo("PivotTableShowDataAs.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				void validateWorksheet() => TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 33, 2, "January"),

					new ExpectedCellValue(sheetName, 34, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 34, 3, 1),
					new ExpectedCellValue(sheetName, 34, 4, 1),
					new ExpectedCellValue(sheetName, 34, 5, 1),
					new ExpectedCellValue(sheetName, 34, 6, 2),
					new ExpectedCellValue(sheetName, 34, 7, 2),
					new ExpectedCellValue(sheetName, 34, 8, 1),
					new ExpectedCellValue(sheetName, 34, 9, 1),
					new ExpectedCellValue(sheetName, 34, 10, 5),

					new ExpectedCellValue(sheetName, 35, 2, "February"),

					new ExpectedCellValue(sheetName, 36, 2, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 36, 3, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 36, 4, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 36, 5, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 36, 6, null),
					new ExpectedCellValue(sheetName, 36, 7, null),
					new ExpectedCellValue(sheetName, 36, 8, 1),
					new ExpectedCellValue(sheetName, 36, 9, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 36, 10, 1),

					new ExpectedCellValue(sheetName, 37, 2, "Tent"),
					new ExpectedCellValue(sheetName, 37, 3, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 37, 4, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 37, 5, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 37, 6, null),
					new ExpectedCellValue(sheetName, 37, 7, 6),
					new ExpectedCellValue(sheetName, 37, 8, null),
					new ExpectedCellValue(sheetName, 37, 9, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 37, 10, 6),

					new ExpectedCellValue(sheetName, 38, 2, "March"),

					new ExpectedCellValue(sheetName, 39, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 39, 3, ExcelErrorValue.Create(eErrorType.Null)),
					new ExpectedCellValue(sheetName, 39, 4, 1),
					new ExpectedCellValue(sheetName, 39, 5, ExcelErrorValue.Create(eErrorType.Null)),
					new ExpectedCellValue(sheetName, 39, 6, null),
					new ExpectedCellValue(sheetName, 39, 7, 2),
					new ExpectedCellValue(sheetName, 39, 8, null),
					new ExpectedCellValue(sheetName, 39, 9, .3333),
					new ExpectedCellValue(sheetName, 39, 10, 2),

					new ExpectedCellValue(sheetName, 40, 2, "Headlamp"),
					new ExpectedCellValue(sheetName, 40, 3, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 40, 4, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 40, 5, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 40, 6, 1),
					new ExpectedCellValue(sheetName, 40, 7, null),
					new ExpectedCellValue(sheetName, 40, 8, null),
					new ExpectedCellValue(sheetName, 40, 9, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 40, 10, 1),

					new ExpectedCellValue(sheetName, 41, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 41, 3, null),
					new ExpectedCellValue(sheetName, 41, 4, null),
					new ExpectedCellValue(sheetName, 41, 5, null),
					new ExpectedCellValue(sheetName, 41, 6, 3),
					new ExpectedCellValue(sheetName, 41, 7, 10),
					new ExpectedCellValue(sheetName, 41, 8, 2),
					new ExpectedCellValue(sheetName, 41, 9, null),
					new ExpectedCellValue(sheetName, 41, 10, 15),
				});

				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable3"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					// Show 'Wholesale Price' data as the percentage of Month: January values.
					wholesalePriceDataField.ShowDataAs = ShowDataAs.Percent;
					wholesalePriceDataField.BaseField = pivotTable.CacheDefinition.GetCacheFieldIndex("Month");
					wholesalePriceDataField.BaseItem = this.GetCacheFieldSharedItemIndex(pivotTable, wholesalePriceDataField.BaseField, "January");
					unitsSoldDataField.ShowDataAs = ShowDataAs.NoCalculation;
					foreach (var field in pivotTable.Fields)
					{
						field.SubtotalLocation = SubtotalLocation.Top;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B30:J41"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				validateWorksheet();
				// Validate subtotal top is on.
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					// January
					new ExpectedCellValue(sheetName, 33, 3, 1),
					new ExpectedCellValue(sheetName, 33, 4, 1),
					new ExpectedCellValue(sheetName, 33, 5, 1),
					new ExpectedCellValue(sheetName, 33, 6, 2),
					new ExpectedCellValue(sheetName, 33, 7, 2),
					new ExpectedCellValue(sheetName, 33, 8, 1),
					new ExpectedCellValue(sheetName, 33, 9, 1),
					new ExpectedCellValue(sheetName, 33, 10, 5),

					// February
					new ExpectedCellValue(sheetName, 35, 3, ExcelErrorValue.Create(eErrorType.Null)),
					new ExpectedCellValue(sheetName, 35, 4, .4787),
					new ExpectedCellValue(sheetName, 35, 5, .2381),
					new ExpectedCellValue(sheetName, 35, 6, null),
					new ExpectedCellValue(sheetName, 35, 7, 6),
					new ExpectedCellValue(sheetName, 35, 8, 1),
					new ExpectedCellValue(sheetName, 35, 9, .2389),
					new ExpectedCellValue(sheetName, 35, 10, 7),

					// March
					new ExpectedCellValue(sheetName, 38, 3, .0601),
					new ExpectedCellValue(sheetName, 38, 4, 1),
					new ExpectedCellValue(sheetName, 38, 5, ExcelErrorValue.Create(eErrorType.Null)),
					new ExpectedCellValue(sheetName, 38, 6, 1),
					new ExpectedCellValue(sheetName, 38, 7, 2),
					new ExpectedCellValue(sheetName, 38, 8, null),
					new ExpectedCellValue(sheetName, 38, 9, .3534),
					new ExpectedCellValue(sheetName, 38, 10, 3),
				});

				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable3"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					// Show 'Wholesale Price' data as the percentage of Month: January values.
					wholesalePriceDataField.ShowDataAs = ShowDataAs.Percent;
					wholesalePriceDataField.BaseField = pivotTable.CacheDefinition.GetCacheFieldIndex("Month");
					wholesalePriceDataField.BaseItem = this.GetCacheFieldSharedItemIndex(pivotTable, wholesalePriceDataField.BaseField, "January");
					unitsSoldDataField.ShowDataAs = ShowDataAs.NoCalculation;
					foreach (var field in pivotTable.Fields)
					{
						field.SubTotalFunctions = eSubTotalFunctions.None;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B30:J41"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				validateWorksheet();
				// Validate subtotal top is off.
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					// January
					new ExpectedCellValue(sheetName, 33, 3, null),
					new ExpectedCellValue(sheetName, 33, 4, null),
					new ExpectedCellValue(sheetName, 33, 5, null),
					new ExpectedCellValue(sheetName, 33, 6, null),
					new ExpectedCellValue(sheetName, 33, 7, null),
					new ExpectedCellValue(sheetName, 33, 8, null),
					new ExpectedCellValue(sheetName, 33, 9, null),
					new ExpectedCellValue(sheetName, 33, 10, null),

					// February
					new ExpectedCellValue(sheetName, 35, 3, null),
					new ExpectedCellValue(sheetName, 35, 4, null),
					new ExpectedCellValue(sheetName, 35, 5, null),
					new ExpectedCellValue(sheetName, 35, 6, null),
					new ExpectedCellValue(sheetName, 35, 7, null),
					new ExpectedCellValue(sheetName, 35, 8, null),
					new ExpectedCellValue(sheetName, 35, 9, null),
					new ExpectedCellValue(sheetName, 35, 10, null),

					// March
					new ExpectedCellValue(sheetName, 38, 3, null),
					new ExpectedCellValue(sheetName, 38, 4, null),
					new ExpectedCellValue(sheetName, 38, 5, null),
					new ExpectedCellValue(sheetName, 38, 6, null),
					new ExpectedCellValue(sheetName, 38, 7, null),
					new ExpectedCellValue(sheetName, 38, 8, null),
					new ExpectedCellValue(sheetName, 38, 9, null),
					new ExpectedCellValue(sheetName, 38, 10, null),
				});
			}
		}


		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableShowDataAs.xlsx")]
		public void PivotTableRefreshShowDataAsColumnFieldsRowFieldsAndColumnDataFieldsPercentOfMonthMarchSubtotalTopOnOffDifferentDataField()
		{
			var file = new FileInfo("PivotTableShowDataAs.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				void validateWorksheet() => TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 33, 2, "January"),

					new ExpectedCellValue(sheetName, 34, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 34, 3, 415.75),
					new ExpectedCellValue(sheetName, 34, 4, 415.75),
					new ExpectedCellValue(sheetName, 34, 5, 415.75),
					new ExpectedCellValue(sheetName, 34, 6, null),
					new ExpectedCellValue(sheetName, 34, 7, 1),
					new ExpectedCellValue(sheetName, 34, 8, null),
					new ExpectedCellValue(sheetName, 34, 9, 1247.25),
					new ExpectedCellValue(sheetName, 34, 10, 2.5),

					new ExpectedCellValue(sheetName, 35, 2, "February"),

					new ExpectedCellValue(sheetName, 36, 2, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 36, 3, null),
					new ExpectedCellValue(sheetName, 36, 4, null),
					new ExpectedCellValue(sheetName, 36, 5, 99),
					new ExpectedCellValue(sheetName, 36, 6, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 36, 7, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 36, 8, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 36, 9, 99),
					new ExpectedCellValue(sheetName, 36, 10, ExcelErrorValue.Create(eErrorType.NA)),

					new ExpectedCellValue(sheetName, 37, 2, "Tent"),
					new ExpectedCellValue(sheetName, 37, 3, null),
					new ExpectedCellValue(sheetName, 37, 4, 199),
					new ExpectedCellValue(sheetName, 37, 5, null),
					new ExpectedCellValue(sheetName, 37, 6, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 37, 7, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 37, 8, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 37, 9, 199),
					new ExpectedCellValue(sheetName, 37, 10, ExcelErrorValue.Create(eErrorType.NA)),

					new ExpectedCellValue(sheetName, 38, 2, "March"),

					new ExpectedCellValue(sheetName, 39, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 39, 3, null),
					new ExpectedCellValue(sheetName, 39, 4, 415.75),
					new ExpectedCellValue(sheetName, 39, 5, null),
					new ExpectedCellValue(sheetName, 39, 6, null),
					new ExpectedCellValue(sheetName, 39, 7, 1),
					new ExpectedCellValue(sheetName, 39, 8, null),
					new ExpectedCellValue(sheetName, 39, 9, 415.75),
					new ExpectedCellValue(sheetName, 39, 10, 1),

					new ExpectedCellValue(sheetName, 40, 2, "Headlamp"),
					new ExpectedCellValue(sheetName, 40, 3, 24.99),
					new ExpectedCellValue(sheetName, 40, 4, null),
					new ExpectedCellValue(sheetName, 40, 5, null),
					new ExpectedCellValue(sheetName, 40, 6, 1),
					new ExpectedCellValue(sheetName, 40, 7, null),
					new ExpectedCellValue(sheetName, 40, 8, null),
					new ExpectedCellValue(sheetName, 40, 9, 24.99),
					new ExpectedCellValue(sheetName, 40, 10, 1),

					new ExpectedCellValue(sheetName, 41, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 41, 3, 440.74),
					new ExpectedCellValue(sheetName, 41, 4, 1030.5),
					new ExpectedCellValue(sheetName, 41, 5, 514.75),
					new ExpectedCellValue(sheetName, 41, 6, null),
					new ExpectedCellValue(sheetName, 41, 7, null),
					new ExpectedCellValue(sheetName, 41, 8, null),
					new ExpectedCellValue(sheetName, 41, 9, 1985.99),
					new ExpectedCellValue(sheetName, 41, 10, null),
				});

				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable3"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					// Show 'Units Sold' data as the percentage of Month: March values.
					unitsSoldDataField.ShowDataAs = ShowDataAs.Percent;
					unitsSoldDataField.BaseField = pivotTable.CacheDefinition.GetCacheFieldIndex("Month");
					unitsSoldDataField.BaseItem = this.GetCacheFieldSharedItemIndex(pivotTable, unitsSoldDataField.BaseField, "March");
					wholesalePriceDataField.ShowDataAs = ShowDataAs.NoCalculation;
					foreach (var field in pivotTable.Fields)
					{
						field.SubtotalLocation = SubtotalLocation.Top;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B30:J41"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				validateWorksheet();
				// Validate subtotal top is on.
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					// January
					new ExpectedCellValue(sheetName, 33, 3, 415.75),
					new ExpectedCellValue(sheetName, 33, 4, 415.75),
					new ExpectedCellValue(sheetName, 33, 5, 415.75),
					new ExpectedCellValue(sheetName, 33, 6, 2),
					new ExpectedCellValue(sheetName, 33, 7, 1),
					new ExpectedCellValue(sheetName, 33, 8, null),
					new ExpectedCellValue(sheetName, 33, 9, 1247.25),
					new ExpectedCellValue(sheetName, 33, 10, 1.6667),

					// February
					new ExpectedCellValue(sheetName, 35, 3, null),
					new ExpectedCellValue(sheetName, 35, 4, 199),
					new ExpectedCellValue(sheetName, 35, 5, 99),
					new ExpectedCellValue(sheetName, 35, 6, ExcelErrorValue.Create(eErrorType.Null)),
					new ExpectedCellValue(sheetName, 35, 7, 3),
					new ExpectedCellValue(sheetName, 35, 8, null),
					new ExpectedCellValue(sheetName, 35, 9, 298),
					new ExpectedCellValue(sheetName, 35, 10, 2.3333),

					// March
					new ExpectedCellValue(sheetName, 38, 3, 24.99),
					new ExpectedCellValue(sheetName, 38, 4, 415.75),
					new ExpectedCellValue(sheetName, 38, 5, null),
					new ExpectedCellValue(sheetName, 38, 6, 1),
					new ExpectedCellValue(sheetName, 38, 7, 1),
					new ExpectedCellValue(sheetName, 38, 8, null),
					new ExpectedCellValue(sheetName, 38, 9, 440.74),
					new ExpectedCellValue(sheetName, 38, 10, 1),
				});

				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable3"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					// Show 'Units Sold' data as the percentage of Month: March values.
					unitsSoldDataField.ShowDataAs = ShowDataAs.Percent;
					unitsSoldDataField.BaseField = pivotTable.CacheDefinition.GetCacheFieldIndex("Month");
					unitsSoldDataField.BaseItem = this.GetCacheFieldSharedItemIndex(pivotTable, unitsSoldDataField.BaseField, "March");
					wholesalePriceDataField.ShowDataAs = ShowDataAs.NoCalculation;
					foreach (var field in pivotTable.Fields)
					{
						field.SubTotalFunctions = eSubTotalFunctions.None;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B30:J41"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				validateWorksheet();
				// Validate subtotal top is off.
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					// January
					new ExpectedCellValue(sheetName, 33, 3, null),
					new ExpectedCellValue(sheetName, 33, 4, null),
					new ExpectedCellValue(sheetName, 33, 5, null),
					new ExpectedCellValue(sheetName, 33, 6, null),
					new ExpectedCellValue(sheetName, 33, 7, null),
					new ExpectedCellValue(sheetName, 33, 8, null),
					new ExpectedCellValue(sheetName, 33, 9, null),
					new ExpectedCellValue(sheetName, 33, 10, null),

					// February
					new ExpectedCellValue(sheetName, 35, 3, null),
					new ExpectedCellValue(sheetName, 35, 4, null),
					new ExpectedCellValue(sheetName, 35, 5, null),
					new ExpectedCellValue(sheetName, 35, 6, null),
					new ExpectedCellValue(sheetName, 35, 7, null),
					new ExpectedCellValue(sheetName, 35, 8, null),
					new ExpectedCellValue(sheetName, 35, 9, null),
					new ExpectedCellValue(sheetName, 35, 10, null),

					// March
					new ExpectedCellValue(sheetName, 38, 3, null),
					new ExpectedCellValue(sheetName, 38, 4, null),
					new ExpectedCellValue(sheetName, 38, 5, null),
					new ExpectedCellValue(sheetName, 38, 6, null),
					new ExpectedCellValue(sheetName, 38, 7, null),
					new ExpectedCellValue(sheetName, 38, 8, null),
					new ExpectedCellValue(sheetName, 38, 9, null),
					new ExpectedCellValue(sheetName, 38, 10, null),
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableShowDataAs.xlsx")]
		public void PivotTableRefreshShowDataAsColumnFieldsRowFieldsAndColumnDataFieldsPercentOfLocationNashvilleSubtotalTopOnOff()
		{
			var file = new FileInfo("PivotTableShowDataAs.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				void validateWorksheet() => TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 33, 2, "January"),

					new ExpectedCellValue(sheetName, 34, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 34, 3, 1),
					new ExpectedCellValue(sheetName, 34, 4, 1),
					new ExpectedCellValue(sheetName, 34, 5, 1),
					new ExpectedCellValue(sheetName, 34, 6, 2),
					new ExpectedCellValue(sheetName, 34, 7, 2),
					new ExpectedCellValue(sheetName, 34, 8, 1),
					new ExpectedCellValue(sheetName, 34, 9, null),
					new ExpectedCellValue(sheetName, 34, 10, 5),

					new ExpectedCellValue(sheetName, 35, 2, "February"),

					new ExpectedCellValue(sheetName, 36, 2, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 36, 3, ExcelErrorValue.Create(eErrorType.Null)),
					new ExpectedCellValue(sheetName, 36, 4, null),
					new ExpectedCellValue(sheetName, 36, 5, null),
					new ExpectedCellValue(sheetName, 36, 6, null),
					new ExpectedCellValue(sheetName, 36, 7, null),
					new ExpectedCellValue(sheetName, 36, 8, 1),
					new ExpectedCellValue(sheetName, 36, 9, null),
					new ExpectedCellValue(sheetName, 36, 10, 1),

					new ExpectedCellValue(sheetName, 37, 2, "Tent"),
					new ExpectedCellValue(sheetName, 37, 3, ExcelErrorValue.Create(eErrorType.Null)),
					new ExpectedCellValue(sheetName, 37, 4, 1),
					new ExpectedCellValue(sheetName, 37, 5, ExcelErrorValue.Create(eErrorType.Null)),
					new ExpectedCellValue(sheetName, 37, 6, null),
					new ExpectedCellValue(sheetName, 37, 7, 6),
					new ExpectedCellValue(sheetName, 37, 8, null),
					new ExpectedCellValue(sheetName, 37, 9, null),
					new ExpectedCellValue(sheetName, 37, 10, 6),

					new ExpectedCellValue(sheetName, 38, 2, "March"),

					new ExpectedCellValue(sheetName, 39, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 39, 3, ExcelErrorValue.Create(eErrorType.Null)),
					new ExpectedCellValue(sheetName, 39, 4, 1),
					new ExpectedCellValue(sheetName, 39, 5, ExcelErrorValue.Create(eErrorType.Null)),
					new ExpectedCellValue(sheetName, 39, 6, null),
					new ExpectedCellValue(sheetName, 39, 7, 2),
					new ExpectedCellValue(sheetName, 39, 8, null),
					new ExpectedCellValue(sheetName, 39, 9, null),
					new ExpectedCellValue(sheetName, 39, 10, 2),

					new ExpectedCellValue(sheetName, 40, 2, "Headlamp"),
					new ExpectedCellValue(sheetName, 40, 3, null),
					new ExpectedCellValue(sheetName, 40, 4, null),
					new ExpectedCellValue(sheetName, 40, 5, ExcelErrorValue.Create(eErrorType.Null)),
					new ExpectedCellValue(sheetName, 40, 6, 1),
					new ExpectedCellValue(sheetName, 40, 7, null),
					new ExpectedCellValue(sheetName, 40, 8, null),
					new ExpectedCellValue(sheetName, 40, 9, null),
					new ExpectedCellValue(sheetName, 40, 10, 1),

					new ExpectedCellValue(sheetName, 41, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 41, 3, .4277),
					new ExpectedCellValue(sheetName, 41, 4, 1),
					new ExpectedCellValue(sheetName, 41, 5, .4995),
					new ExpectedCellValue(sheetName, 41, 6, 3),
					new ExpectedCellValue(sheetName, 41, 7, 10),
					new ExpectedCellValue(sheetName, 41, 8, 2),
					new ExpectedCellValue(sheetName, 41, 9, null),
					new ExpectedCellValue(sheetName, 41, 10, 15),
				});

				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable3"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					// Show 'Wholesale Price' data as the percentage of Location: Nashville values.
					wholesalePriceDataField.ShowDataAs = ShowDataAs.Percent;
					wholesalePriceDataField.BaseField = pivotTable.CacheDefinition.GetCacheFieldIndex("Location");
					wholesalePriceDataField.BaseItem = this.GetCacheFieldSharedItemIndex(pivotTable, wholesalePriceDataField.BaseField, "Nashville");
					unitsSoldDataField.ShowDataAs = ShowDataAs.NoCalculation;
					foreach (var field in pivotTable.Fields)
					{
						field.SubtotalLocation = SubtotalLocation.Top;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B30:J41"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				validateWorksheet();
				// Validate subtotal top is on.
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					// January
					new ExpectedCellValue(sheetName, 33, 3, 1),
					new ExpectedCellValue(sheetName, 33, 4, 1),
					new ExpectedCellValue(sheetName, 33, 5, 1),
					new ExpectedCellValue(sheetName, 33, 6, 2),
					new ExpectedCellValue(sheetName, 33, 7, 2),
					new ExpectedCellValue(sheetName, 33, 8, 1),
					new ExpectedCellValue(sheetName, 33, 9, null),
					new ExpectedCellValue(sheetName, 33, 10, 5),

					// February
					new ExpectedCellValue(sheetName, 35, 3, ExcelErrorValue.Create(eErrorType.Null)),
					new ExpectedCellValue(sheetName, 35, 4, 1),
					new ExpectedCellValue(sheetName, 35, 5, .4975),
					new ExpectedCellValue(sheetName, 35, 6, null),
					new ExpectedCellValue(sheetName, 35, 7, 6),
					new ExpectedCellValue(sheetName, 35, 8, 1),
					new ExpectedCellValue(sheetName, 35, 9, null),
					new ExpectedCellValue(sheetName, 35, 10, 7),

					// March
					new ExpectedCellValue(sheetName, 38, 3, .0601),
					new ExpectedCellValue(sheetName, 38, 4, 1),
					new ExpectedCellValue(sheetName, 38, 5, ExcelErrorValue.Create(eErrorType.Null)),
					new ExpectedCellValue(sheetName, 38, 6, 1),
					new ExpectedCellValue(sheetName, 38, 7, 2),
					new ExpectedCellValue(sheetName, 38, 8, null),
					new ExpectedCellValue(sheetName, 38, 9, null),
					new ExpectedCellValue(sheetName, 38, 10, 3),
				});

				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable3"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					// Show 'Wholesale Price' data as the percentage of Location: Nashville values.
					wholesalePriceDataField.ShowDataAs = ShowDataAs.Percent;
					wholesalePriceDataField.BaseField = pivotTable.CacheDefinition.GetCacheFieldIndex("Location");
					wholesalePriceDataField.BaseItem = this.GetCacheFieldSharedItemIndex(pivotTable, wholesalePriceDataField.BaseField, "Nashville");
					unitsSoldDataField.ShowDataAs = ShowDataAs.NoCalculation;
					foreach (var field in pivotTable.Fields)
					{
						field.SubTotalFunctions = eSubTotalFunctions.None;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B30:J41"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				validateWorksheet();
				// Validate subtotal top is off.
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					// January
					new ExpectedCellValue(sheetName, 33, 3, null),
					new ExpectedCellValue(sheetName, 33, 4, null),
					new ExpectedCellValue(sheetName, 33, 5, null),
					new ExpectedCellValue(sheetName, 33, 6, null),
					new ExpectedCellValue(sheetName, 33, 7, null),
					new ExpectedCellValue(sheetName, 33, 8, null),
					new ExpectedCellValue(sheetName, 33, 9, null),
					new ExpectedCellValue(sheetName, 33, 10, null),

					// February
					new ExpectedCellValue(sheetName, 35, 3, null),
					new ExpectedCellValue(sheetName, 35, 4, null),
					new ExpectedCellValue(sheetName, 35, 5, null),
					new ExpectedCellValue(sheetName, 35, 6, null),
					new ExpectedCellValue(sheetName, 35, 7, null),
					new ExpectedCellValue(sheetName, 35, 8, null),
					new ExpectedCellValue(sheetName, 35, 9, null),
					new ExpectedCellValue(sheetName, 35, 10, null),

					// March
					new ExpectedCellValue(sheetName, 38, 3, null),
					new ExpectedCellValue(sheetName, 38, 4, null),
					new ExpectedCellValue(sheetName, 38, 5, null),
					new ExpectedCellValue(sheetName, 38, 6, null),
					new ExpectedCellValue(sheetName, 38, 7, null),
					new ExpectedCellValue(sheetName, 38, 8, null),
					new ExpectedCellValue(sheetName, 38, 9, null),
					new ExpectedCellValue(sheetName, 38, 10, null),
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableShowDataAs.xlsx")]
		public void PivotTableRefreshShowDataAsColumnFieldsRowFieldsAndColumnDataFieldsPercentOfUnusedFieldSubtotalTopOnOff()
		{
			var file = new FileInfo("PivotTableShowDataAs.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				void validateWorksheet() => TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 33, 2, "January"),

					new ExpectedCellValue(sheetName, 34, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 34, 3, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 34, 4, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 34, 5, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 34, 6, 2),
					new ExpectedCellValue(sheetName, 34, 7, 2),
					new ExpectedCellValue(sheetName, 34, 8, 1),
					new ExpectedCellValue(sheetName, 34, 9, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 34, 10, 5),

					new ExpectedCellValue(sheetName, 35, 2, "February"),

					new ExpectedCellValue(sheetName, 36, 2, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 36, 3, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 36, 4, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 36, 5, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 36, 6, null),
					new ExpectedCellValue(sheetName, 36, 7, null),
					new ExpectedCellValue(sheetName, 36, 8, 1),
					new ExpectedCellValue(sheetName, 36, 9, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 36, 10, 1),

					new ExpectedCellValue(sheetName, 37, 2, "Tent"),
					new ExpectedCellValue(sheetName, 37, 3, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 37, 4, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 37, 5, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 37, 6, null),
					new ExpectedCellValue(sheetName, 37, 7, 6),
					new ExpectedCellValue(sheetName, 37, 8, null),
					new ExpectedCellValue(sheetName, 37, 9, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 37, 10, 6),

					new ExpectedCellValue(sheetName, 38, 2, "March"),

					new ExpectedCellValue(sheetName, 39, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 39, 3, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 39, 4, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 39, 5, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 39, 6, null),
					new ExpectedCellValue(sheetName, 39, 7, 2),
					new ExpectedCellValue(sheetName, 39, 8, null),
					new ExpectedCellValue(sheetName, 39, 9, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 39, 10, 2),

					new ExpectedCellValue(sheetName, 40, 2, "Headlamp"),
					new ExpectedCellValue(sheetName, 40, 3, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 40, 4, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 40, 5, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 40, 6, 1),
					new ExpectedCellValue(sheetName, 40, 7, null),
					new ExpectedCellValue(sheetName, 40, 8, null),
					new ExpectedCellValue(sheetName, 40, 9, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 40, 10, 1),

					new ExpectedCellValue(sheetName, 41, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 41, 3, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 41, 4, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 41, 5, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 41, 6, 3),
					new ExpectedCellValue(sheetName, 41, 7, 10),
					new ExpectedCellValue(sheetName, 41, 8, 2),
					new ExpectedCellValue(sheetName, 41, 9, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 41, 10, 15),
				});

				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable3"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					// Show 'Wholesale Price' data as the percentage of Transaction[1] values.
					wholesalePriceDataField.ShowDataAs = ShowDataAs.Percent;
					wholesalePriceDataField.BaseField = pivotTable.CacheDefinition.GetCacheFieldIndex("Transaction");
					wholesalePriceDataField.BaseItem = 1;
					unitsSoldDataField.ShowDataAs = ShowDataAs.NoCalculation;
					foreach (var field in pivotTable.Fields)
					{
						field.SubtotalLocation = SubtotalLocation.Top;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B30:J41"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				validateWorksheet();
				// Validate subtotal top is on.
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					// January
					new ExpectedCellValue(sheetName, 33, 3, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 33, 4, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 33, 5, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 33, 6, 2),
					new ExpectedCellValue(sheetName, 33, 7, 2),
					new ExpectedCellValue(sheetName, 33, 8, 1),
					new ExpectedCellValue(sheetName, 33, 9, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 33, 10, 5),

					// February
					new ExpectedCellValue(sheetName, 35, 3, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 35, 4, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 35, 5, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 35, 6, null),
					new ExpectedCellValue(sheetName, 35, 7, 6),
					new ExpectedCellValue(sheetName, 35, 8, 1),
					new ExpectedCellValue(sheetName, 35, 9, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 35, 10, 7),

					// March
					new ExpectedCellValue(sheetName, 38, 3, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 38, 4, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 38, 5, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 38, 6, 1),
					new ExpectedCellValue(sheetName, 38, 7, 2),
					new ExpectedCellValue(sheetName, 38, 8, null),
					new ExpectedCellValue(sheetName, 38, 9, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 38, 10, 3),
				});

				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable3"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					// Show 'Wholesale Price' data as the percentage of Location[1] values.
					wholesalePriceDataField.ShowDataAs = ShowDataAs.Percent;
					wholesalePriceDataField.BaseField = pivotTable.CacheDefinition.GetCacheFieldIndex("Transaction");
					wholesalePriceDataField.BaseItem = 1;
					unitsSoldDataField.ShowDataAs = ShowDataAs.NoCalculation;
					foreach (var field in pivotTable.Fields)
					{
						field.SubTotalFunctions = eSubTotalFunctions.None;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B30:J41"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				validateWorksheet();
				// Validate subtotal top is off.
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					// January
					new ExpectedCellValue(sheetName, 33, 3, null),
					new ExpectedCellValue(sheetName, 33, 4, null),
					new ExpectedCellValue(sheetName, 33, 5, null),
					new ExpectedCellValue(sheetName, 33, 6, null),
					new ExpectedCellValue(sheetName, 33, 7, null),
					new ExpectedCellValue(sheetName, 33, 8, null),
					new ExpectedCellValue(sheetName, 33, 9, null),
					new ExpectedCellValue(sheetName, 33, 10, null),

					// February
					new ExpectedCellValue(sheetName, 35, 3, null),
					new ExpectedCellValue(sheetName, 35, 4, null),
					new ExpectedCellValue(sheetName, 35, 5, null),
					new ExpectedCellValue(sheetName, 35, 6, null),
					new ExpectedCellValue(sheetName, 35, 7, null),
					new ExpectedCellValue(sheetName, 35, 8, null),
					new ExpectedCellValue(sheetName, 35, 9, null),
					new ExpectedCellValue(sheetName, 35, 10, null),

					// March
					new ExpectedCellValue(sheetName, 38, 3, null),
					new ExpectedCellValue(sheetName, 38, 4, null),
					new ExpectedCellValue(sheetName, 38, 5, null),
					new ExpectedCellValue(sheetName, 38, 6, null),
					new ExpectedCellValue(sheetName, 38, 7, null),
					new ExpectedCellValue(sheetName, 38, 8, null),
					new ExpectedCellValue(sheetName, 38, 9, null),
					new ExpectedCellValue(sheetName, 38, 10, null),
				});
			}
		}


		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableShowDataAs.xlsx")]
		public void PivotTableRefreshShowDataAsColumnFieldsRowFieldsAndColumnDataFieldsPercentOfDataFieldSubtotalTopOnOff()
		{
			var file = new FileInfo("PivotTableShowDataAs.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				void validateWorksheet() => TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 33, 2, "January"),

					new ExpectedCellValue(sheetName, 34, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 34, 3, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 34, 4, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 34, 5, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 34, 6, 2),
					new ExpectedCellValue(sheetName, 34, 7, 2),
					new ExpectedCellValue(sheetName, 34, 8, 1),
					new ExpectedCellValue(sheetName, 34, 9, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 34, 10, 5),

					new ExpectedCellValue(sheetName, 35, 2, "February"),

					new ExpectedCellValue(sheetName, 36, 2, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 36, 3, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 36, 4, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 36, 5, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 36, 6, null),
					new ExpectedCellValue(sheetName, 36, 7, null),
					new ExpectedCellValue(sheetName, 36, 8, 1),
					new ExpectedCellValue(sheetName, 36, 9, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 36, 10, 1),

					new ExpectedCellValue(sheetName, 37, 2, "Tent"),
					new ExpectedCellValue(sheetName, 37, 3, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 37, 4, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 37, 5, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 37, 6, null),
					new ExpectedCellValue(sheetName, 37, 7, 6),
					new ExpectedCellValue(sheetName, 37, 8, null),
					new ExpectedCellValue(sheetName, 37, 9, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 37, 10, 6),

					new ExpectedCellValue(sheetName, 38, 2, "March"),

					new ExpectedCellValue(sheetName, 39, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 39, 3, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 39, 4, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 39, 5, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 39, 6, null),
					new ExpectedCellValue(sheetName, 39, 7, 2),
					new ExpectedCellValue(sheetName, 39, 8, null),
					new ExpectedCellValue(sheetName, 39, 9, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 39, 10, 2),

					new ExpectedCellValue(sheetName, 40, 2, "Headlamp"),
					new ExpectedCellValue(sheetName, 40, 3, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 40, 4, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 40, 5, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 40, 6, 1),
					new ExpectedCellValue(sheetName, 40, 7, null),
					new ExpectedCellValue(sheetName, 40, 8, null),
					new ExpectedCellValue(sheetName, 40, 9, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 40, 10, 1),

					new ExpectedCellValue(sheetName, 41, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 41, 3, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 41, 4, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 41, 5, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 41, 6, 3),
					new ExpectedCellValue(sheetName, 41, 7, 10),
					new ExpectedCellValue(sheetName, 41, 8, 2),
					new ExpectedCellValue(sheetName, 41, 9, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 41, 10, 15),
				});

				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable3"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					// Show 'Wholesale Price' data as the percentage of Wholesale Price[1] values.
					wholesalePriceDataField.ShowDataAs = ShowDataAs.Percent;
					wholesalePriceDataField.BaseField = pivotTable.CacheDefinition.GetCacheFieldIndex("Wholesale Price");
					wholesalePriceDataField.BaseItem = 1;
					unitsSoldDataField.ShowDataAs = ShowDataAs.NoCalculation;
					foreach (var field in pivotTable.Fields)
					{
						field.SubtotalLocation = SubtotalLocation.Top;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B30:J41"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				validateWorksheet();
				// Validate subtotal top is on.
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					// January
					new ExpectedCellValue(sheetName, 33, 3, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 33, 4, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 33, 5, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 33, 6, 2),
					new ExpectedCellValue(sheetName, 33, 7, 2),
					new ExpectedCellValue(sheetName, 33, 8, 1),
					new ExpectedCellValue(sheetName, 33, 9, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 33, 10, 5),

					// February
					new ExpectedCellValue(sheetName, 35, 3, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 35, 4, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 35, 5, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 35, 6, null),
					new ExpectedCellValue(sheetName, 35, 7, 6),
					new ExpectedCellValue(sheetName, 35, 8, 1),
					new ExpectedCellValue(sheetName, 35, 9, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 35, 10, 7),

					// March
					new ExpectedCellValue(sheetName, 38, 3, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 38, 4, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 38, 5, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 38, 6, 1),
					new ExpectedCellValue(sheetName, 38, 7, 2),
					new ExpectedCellValue(sheetName, 38, 8, null),
					new ExpectedCellValue(sheetName, 38, 9, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 38, 10, 3),
				});

				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable3"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					// Show 'Wholesale Price' data as the percentage of Wholesale Price[1] values.
					wholesalePriceDataField.ShowDataAs = ShowDataAs.Percent;
					wholesalePriceDataField.BaseField = pivotTable.CacheDefinition.GetCacheFieldIndex("Wholesale Price");
					wholesalePriceDataField.BaseItem = 1;
					unitsSoldDataField.ShowDataAs = ShowDataAs.NoCalculation;
					foreach (var field in pivotTable.Fields)
					{
						field.SubTotalFunctions = eSubTotalFunctions.None;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B30:J41"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				validateWorksheet();
				// Validate subtotal top is off.
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					// January
					new ExpectedCellValue(sheetName, 33, 3, null),
					new ExpectedCellValue(sheetName, 33, 4, null),
					new ExpectedCellValue(sheetName, 33, 5, null),
					new ExpectedCellValue(sheetName, 33, 6, null),
					new ExpectedCellValue(sheetName, 33, 7, null),
					new ExpectedCellValue(sheetName, 33, 8, null),
					new ExpectedCellValue(sheetName, 33, 9, null),
					new ExpectedCellValue(sheetName, 33, 10, null),

					// February
					new ExpectedCellValue(sheetName, 35, 3, null),
					new ExpectedCellValue(sheetName, 35, 4, null),
					new ExpectedCellValue(sheetName, 35, 5, null),
					new ExpectedCellValue(sheetName, 35, 6, null),
					new ExpectedCellValue(sheetName, 35, 7, null),
					new ExpectedCellValue(sheetName, 35, 8, null),
					new ExpectedCellValue(sheetName, 35, 9, null),
					new ExpectedCellValue(sheetName, 35, 10, null),

					// March
					new ExpectedCellValue(sheetName, 38, 3, null),
					new ExpectedCellValue(sheetName, 38, 4, null),
					new ExpectedCellValue(sheetName, 38, 5, null),
					new ExpectedCellValue(sheetName, 38, 6, null),
					new ExpectedCellValue(sheetName, 38, 7, null),
					new ExpectedCellValue(sheetName, 38, 8, null),
					new ExpectedCellValue(sheetName, 38, 9, null),
					new ExpectedCellValue(sheetName, 38, 10, null),
				});
			}
		}


		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableShowDataAs.xlsx")]
		public void PivotTableRefreshShowDataAsPercentOfWithErrorStates()
		{
			var file = new FileInfo("PivotTableShowDataAs.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				using (var package = new ExcelPackage(file))
				{
					var dataSheet = package.Workbook.Worksheets["Sheet1"];
					dataSheet.Cells["F6"].Value = 0;
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable3"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					wholesalePriceDataField.ShowDataAs = ShowDataAs.Percent;
					wholesalePriceDataField.BaseField = pivotTable.CacheDefinition.GetCacheFieldIndex("Month");
					wholesalePriceDataField.BaseItem = 0;
					unitsSoldDataField.ShowDataAs = ShowDataAs.NoCalculation;
					foreach (var field in pivotTable.Fields)
					{
						field.SubtotalLocation = SubtotalLocation.Top;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B30:J41"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 31, 2, null),
					new ExpectedCellValue(sheetName, 31, 3, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 31, 4, null),
					new ExpectedCellValue(sheetName, 31, 5, null),
					new ExpectedCellValue(sheetName, 31, 6, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 31, 7, null),
					new ExpectedCellValue(sheetName, 31, 8, null),
					new ExpectedCellValue(sheetName, 31, 9, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 31, 10, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 32, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 32, 3, "Chicago"),
					new ExpectedCellValue(sheetName, 32, 4, "Nashville"),
					new ExpectedCellValue(sheetName, 32, 5, "San Francisco"),
					new ExpectedCellValue(sheetName, 32, 6, "Chicago"),
					new ExpectedCellValue(sheetName, 32, 7, "Nashville"),
					new ExpectedCellValue(sheetName, 32, 8, "San Francisco"),
					new ExpectedCellValue(sheetName, 32, 9, null),
					new ExpectedCellValue(sheetName, 32, 10, null),
					new ExpectedCellValue(sheetName, 33, 2, "January"),
					new ExpectedCellValue(sheetName, 33, 3, ExcelErrorValue.Create(eErrorType.Div0)),
					new ExpectedCellValue(sheetName, 33, 4, 1),
					new ExpectedCellValue(sheetName, 33, 5, 1),
					new ExpectedCellValue(sheetName, 33, 9, 1),
					new ExpectedCellValue(sheetName, 33, 10, 5),
					new ExpectedCellValue(sheetName, 34, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 34, 3, ExcelErrorValue.Create(eErrorType.Div0)),
					new ExpectedCellValue(sheetName, 34, 4, 1),
					new ExpectedCellValue(sheetName, 34, 5, 1),
					new ExpectedCellValue(sheetName, 34, 6, 2),
					new ExpectedCellValue(sheetName, 34, 7, 2),
					new ExpectedCellValue(sheetName, 34, 8, 1),
					new ExpectedCellValue(sheetName, 34, 9, 1),
					new ExpectedCellValue(sheetName, 34, 10, 5),
					new ExpectedCellValue(sheetName, 35, 2, "February"),
					new ExpectedCellValue(sheetName, 35, 3, ExcelErrorValue.Create(eErrorType.Null)),
					new ExpectedCellValue(sheetName, 35, 4, .4787),
					new ExpectedCellValue(sheetName, 35, 5, .2381),
					new ExpectedCellValue(sheetName, 35, 6, null),
					new ExpectedCellValue(sheetName, 35, 7, 6),
					new ExpectedCellValue(sheetName, 35, 8, 1),
					new ExpectedCellValue(sheetName, 35, 9, .3584),
					new ExpectedCellValue(sheetName, 35, 10, 7),
					new ExpectedCellValue(sheetName, 36, 2, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 36, 3, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 36, 4, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 36, 5, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 36, 6, null),
					new ExpectedCellValue(sheetName, 36, 7, null),
					new ExpectedCellValue(sheetName, 36, 8, 1),
					new ExpectedCellValue(sheetName, 36, 9, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 36, 10, 1),
					new ExpectedCellValue(sheetName, 37, 2, "Tent"),
					new ExpectedCellValue(sheetName, 37, 3, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 37, 4, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 37, 5, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 37, 6, null),
					new ExpectedCellValue(sheetName, 37, 7, 6),
					new ExpectedCellValue(sheetName, 37, 8, null),
					new ExpectedCellValue(sheetName, 37, 9, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 37, 10, 6),
					new ExpectedCellValue(sheetName, 38, 2, "March"),
					new ExpectedCellValue(sheetName, 38, 3, ExcelErrorValue.Create(eErrorType.Div0)),
					new ExpectedCellValue(sheetName, 38, 4, 1),
					new ExpectedCellValue(sheetName, 38, 5, ExcelErrorValue.Create(eErrorType.Null)),
					new ExpectedCellValue(sheetName, 38, 6, 1),
					new ExpectedCellValue(sheetName, 38, 7, 2),
					new ExpectedCellValue(sheetName, 38, 8, null),
					new ExpectedCellValue(sheetName, 38, 9, .5301),
					new ExpectedCellValue(sheetName, 38, 10, 3),
					new ExpectedCellValue(sheetName, 39, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 39, 3, ExcelErrorValue.Create(eErrorType.Null)),
					new ExpectedCellValue(sheetName, 39, 4, 1),
					new ExpectedCellValue(sheetName, 39, 5, ExcelErrorValue.Create(eErrorType.Null)),
					new ExpectedCellValue(sheetName, 39, 6, null),
					new ExpectedCellValue(sheetName, 39, 7, 2),
					new ExpectedCellValue(sheetName, 39, 8, null),
					new ExpectedCellValue(sheetName, 39, 9, .5000),
					new ExpectedCellValue(sheetName, 39, 10, 2),
					new ExpectedCellValue(sheetName, 40, 2, "Headlamp"),
					new ExpectedCellValue(sheetName, 40, 3, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 40, 4, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 40, 5, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 40, 6, 1),
					new ExpectedCellValue(sheetName, 40, 7, null),
					new ExpectedCellValue(sheetName, 40, 8, null),
					new ExpectedCellValue(sheetName, 40, 9, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 40, 10, 1),
					new ExpectedCellValue(sheetName, 41, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 41, 3, null),
					new ExpectedCellValue(sheetName, 41, 4, null),
					new ExpectedCellValue(sheetName, 41, 5, null),
					new ExpectedCellValue(sheetName, 41, 6, 3),
					new ExpectedCellValue(sheetName, 41, 7, 10),
					new ExpectedCellValue(sheetName, 41, 8, 2),
					new ExpectedCellValue(sheetName, 41, 9, null),
					new ExpectedCellValue(sheetName, 41, 10, 15)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\ShowDataAsComplex.xlsx")]
		public void PivotTableRefreshShowDataAsPercentOfTabularFormColumnDataFields()
		{
			var file = new FileInfo("ShowDataAsComplex.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "Sheet1";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable1"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B3:H17"), pivotTable.Address);
					Assert.AreEqual(15, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 3, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 4, 3, "Count of Net Change"),
					new ExpectedCellValue(sheetName, 4, 5, "Sum of Balance at End Date"),
					new ExpectedCellValue(sheetName, 4, 7, "Total Count of Net Change"),
					new ExpectedCellValue(sheetName, 4, 8, "Total Sum of Balance at End Date"),
					new ExpectedCellValue(sheetName, 5, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 5, 3, "FALSE"),
					new ExpectedCellValue(sheetName, 5, 4, "TRUE"),
					new ExpectedCellValue(sheetName, 5, 5, "FALSE"),
					new ExpectedCellValue(sheetName, 5, 6, "TRUE"),
					new ExpectedCellValue(sheetName, 6, 2, "Begin-Total"),
					new ExpectedCellValue(sheetName, 6, 3, 1),
					new ExpectedCellValue(sheetName, 6, 4, 0),
					new ExpectedCellValue(sheetName, 6, 5, null),
					new ExpectedCellValue(sheetName, 6, 6, null),
					new ExpectedCellValue(sheetName, 6, 7, 1),
					new ExpectedCellValue(sheetName, 6, 8, null),
					new ExpectedCellValue(sheetName, 7, 2, "Accounts Receivable"),
					new ExpectedCellValue(sheetName, 7, 3, 1),
					new ExpectedCellValue(sheetName, 7, 4, 0),
					new ExpectedCellValue(sheetName, 7, 5, null),
					new ExpectedCellValue(sheetName, 7, 6, null),
					new ExpectedCellValue(sheetName, 7, 7, 1),
					new ExpectedCellValue(sheetName, 7, 8, null),
					new ExpectedCellValue(sheetName, 8, 2, "ASSETS"),
					new ExpectedCellValue(sheetName, 8, 3, 1),
					new ExpectedCellValue(sheetName, 8, 4, 0),
					new ExpectedCellValue(sheetName, 8, 5, null),
					new ExpectedCellValue(sheetName, 8, 6, null),
					new ExpectedCellValue(sheetName, 8, 7, 1),
					new ExpectedCellValue(sheetName, 8, 8, null),
					new ExpectedCellValue(sheetName, 9, 2, "End-Total"),
					new ExpectedCellValue(sheetName, 9, 3, 0.333333333333333),
					new ExpectedCellValue(sheetName, 9, 4, 0.666666666666667),
					new ExpectedCellValue(sheetName, 9, 5, 2.12307706591973),
					new ExpectedCellValue(sheetName, 9, 6, 1),
					new ExpectedCellValue(sheetName, 9, 7, 1),
					new ExpectedCellValue(sheetName, 9, 8, null),
					new ExpectedCellValue(sheetName, 10, 2, "Accounts Receivable, Total"),
					new ExpectedCellValue(sheetName, 10, 3, 0),
					new ExpectedCellValue(sheetName, 10, 4, 1),
					new ExpectedCellValue(sheetName, 10, 5, ExcelErrorValue.Create(eErrorType.Null)),
					new ExpectedCellValue(sheetName, 10, 6, 1),
					new ExpectedCellValue(sheetName, 10, 7, 1),
					new ExpectedCellValue(sheetName, 10, 8, null),
					new ExpectedCellValue(sheetName, 11, 2, "Liquid Assets, Total"),
					new ExpectedCellValue(sheetName, 11, 3, 1),
					new ExpectedCellValue(sheetName, 11, 4, 0),
					new ExpectedCellValue(sheetName, 11, 5, null),
					new ExpectedCellValue(sheetName, 11, 6, null),
					new ExpectedCellValue(sheetName, 11, 7, 1),
					new ExpectedCellValue(sheetName, 11, 8, null),
					new ExpectedCellValue(sheetName, 12, 2, "Securities, Total"),
					new ExpectedCellValue(sheetName, 12, 3, 0),
					new ExpectedCellValue(sheetName, 12, 4, 1),
					new ExpectedCellValue(sheetName, 12, 5, ExcelErrorValue.Create(eErrorType.Null)),
					new ExpectedCellValue(sheetName, 12, 6, ExcelErrorValue.Create(eErrorType.Div0)),
					new ExpectedCellValue(sheetName, 12, 7, 1),
					new ExpectedCellValue(sheetName, 12, 8, null),
					new ExpectedCellValue(sheetName, 13, 2, "Posting"),
					new ExpectedCellValue(sheetName, 13, 3, 0.666666666666667),
					new ExpectedCellValue(sheetName, 13, 4, 0.333333333333333),
					new ExpectedCellValue(sheetName, 13, 5, 4.74378108319677),
					new ExpectedCellValue(sheetName, 13, 6, 1),
					new ExpectedCellValue(sheetName, 13, 7, 1),
					new ExpectedCellValue(sheetName, 13, 8, null),
					new ExpectedCellValue(sheetName, 14, 2, "Cash"),
					new ExpectedCellValue(sheetName, 14, 3, 1),
					new ExpectedCellValue(sheetName, 14, 4, 0),
					new ExpectedCellValue(sheetName, 14, 5, null),
					new ExpectedCellValue(sheetName, 14, 6, null),
					new ExpectedCellValue(sheetName, 14, 7, 1),
					new ExpectedCellValue(sheetName, 14, 8, null),
					new ExpectedCellValue(sheetName, 15, 2, "Customers, EU"),
					new ExpectedCellValue(sheetName, 15, 3, 1),
					new ExpectedCellValue(sheetName, 15, 4, 0),
					new ExpectedCellValue(sheetName, 15, 5, null),
					new ExpectedCellValue(sheetName, 15, 6, null),
					new ExpectedCellValue(sheetName, 15, 7, 1),
					new ExpectedCellValue(sheetName, 15, 8, null),
					new ExpectedCellValue(sheetName, 16, 2, "Customers, North America"),
					new ExpectedCellValue(sheetName, 16, 3, 0),
					new ExpectedCellValue(sheetName, 16, 4, 1),
					new ExpectedCellValue(sheetName, 16, 5, ExcelErrorValue.Create(eErrorType.Null)),
					new ExpectedCellValue(sheetName, 16, 6, 1),
					new ExpectedCellValue(sheetName, 16, 7, 1),
					new ExpectedCellValue(sheetName, 16, 8, null),
					new ExpectedCellValue(sheetName, 17, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 17, 3, 0.625),
					new ExpectedCellValue(sheetName, 17, 4, 0.375),
					new ExpectedCellValue(sheetName, 17, 5, 3.04613915570146),
					new ExpectedCellValue(sheetName, 17, 6, 1),
					new ExpectedCellValue(sheetName, 17, 7, 1),
					new ExpectedCellValue(sheetName, 17, 8, null),
				});
			}
		}
		#endregion

		#region PercentOfParentRow Tests
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableShowDataAs.xlsx")]
		public void PivotTableRefreshShowDataAsColumnFieldsRowFieldsAndColumnDataFieldsPercentOfParentRowDataFieldSubtotalTopOnOff()
		{
			var file = new FileInfo("PivotTableShowDataAs.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				void validateWorksheet() => TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 33, 2, "January"),

					new ExpectedCellValue(sheetName, 34, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 34, 3, 1),
					new ExpectedCellValue(sheetName, 34, 4, 1),
					new ExpectedCellValue(sheetName, 34, 5, 1),
					new ExpectedCellValue(sheetName, 34, 6, 2),
					new ExpectedCellValue(sheetName, 34, 7, 2),
					new ExpectedCellValue(sheetName, 34, 8, 1),
					new ExpectedCellValue(sheetName, 34, 9, 1),
					new ExpectedCellValue(sheetName, 34, 10, 5),

					new ExpectedCellValue(sheetName, 35, 2, "February"),

					new ExpectedCellValue(sheetName, 36, 2, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 36, 3, null),
					new ExpectedCellValue(sheetName, 36, 4, 0),
					new ExpectedCellValue(sheetName, 36, 5, 1),
					new ExpectedCellValue(sheetName, 36, 6, null),
					new ExpectedCellValue(sheetName, 36, 7, null),
					new ExpectedCellValue(sheetName, 36, 8, 1),
					new ExpectedCellValue(sheetName, 36, 9, .3322),
					new ExpectedCellValue(sheetName, 36, 10, 1),

					new ExpectedCellValue(sheetName, 37, 2, "Tent"),
					new ExpectedCellValue(sheetName, 37, 3, null),
					new ExpectedCellValue(sheetName, 37, 4, 1),
					new ExpectedCellValue(sheetName, 37, 5, 0),
					new ExpectedCellValue(sheetName, 37, 6, null),
					new ExpectedCellValue(sheetName, 37, 7, 6),
					new ExpectedCellValue(sheetName, 37, 8, null),
					new ExpectedCellValue(sheetName, 37, 9, .6678),
					new ExpectedCellValue(sheetName, 37, 10, 6),

					new ExpectedCellValue(sheetName, 38, 2, "March"),

					new ExpectedCellValue(sheetName, 39, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 39, 3, 0),
					new ExpectedCellValue(sheetName, 39, 4, 1),
					new ExpectedCellValue(sheetName, 39, 5, null),
					new ExpectedCellValue(sheetName, 39, 6, null),
					new ExpectedCellValue(sheetName, 39, 7, 2),
					new ExpectedCellValue(sheetName, 39, 8, null),
					new ExpectedCellValue(sheetName, 39, 9, .9433),
					new ExpectedCellValue(sheetName, 39, 10, 2),

					new ExpectedCellValue(sheetName, 40, 2, "Headlamp"),
					new ExpectedCellValue(sheetName, 40, 3, 1),
					new ExpectedCellValue(sheetName, 40, 4, 0),
					new ExpectedCellValue(sheetName, 40, 5, null),
					new ExpectedCellValue(sheetName, 40, 6, 1),
					new ExpectedCellValue(sheetName, 40, 7, null),
					new ExpectedCellValue(sheetName, 40, 8, null),
					new ExpectedCellValue(sheetName, 40, 9, .0567),
					new ExpectedCellValue(sheetName, 40, 10, 1),

					new ExpectedCellValue(sheetName, 41, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 41, 3, 1),
					new ExpectedCellValue(sheetName, 41, 4, 1),
					new ExpectedCellValue(sheetName, 41, 5, 1),
					new ExpectedCellValue(sheetName, 41, 6, 3),
					new ExpectedCellValue(sheetName, 41, 7, 10),
					new ExpectedCellValue(sheetName, 41, 8, 2),
					new ExpectedCellValue(sheetName, 41, 9, 1),
					new ExpectedCellValue(sheetName, 41, 10, 15),
				});

				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable3"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					// Show 'Wholesale Price' data as the percentage of its parent row.
					wholesalePriceDataField.ShowDataAs = ShowDataAs.PercentOfParentRow;
					unitsSoldDataField.ShowDataAs = ShowDataAs.NoCalculation;
					foreach (var field in pivotTable.Fields)
					{
						field.SubtotalLocation = SubtotalLocation.Top;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B30:J41"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				validateWorksheet();
				// Validate subtotal top is on.
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					// January
					new ExpectedCellValue(sheetName, 33, 3, .9433),
					new ExpectedCellValue(sheetName, 33, 4, .4034),
					new ExpectedCellValue(sheetName, 33, 5, .8077),
					new ExpectedCellValue(sheetName, 33, 6, 2),
					new ExpectedCellValue(sheetName, 33, 7, 2),
					new ExpectedCellValue(sheetName, 33, 8, 1),
					new ExpectedCellValue(sheetName, 33, 9, .6280),
					new ExpectedCellValue(sheetName, 33, 10, 5),

					// February
					new ExpectedCellValue(sheetName, 35, 3, 0),
					new ExpectedCellValue(sheetName, 35, 4, .1931),
					new ExpectedCellValue(sheetName, 35, 5, .1923),
					new ExpectedCellValue(sheetName, 35, 6, null),
					new ExpectedCellValue(sheetName, 35, 7, 6),
					new ExpectedCellValue(sheetName, 35, 8, 1),
					new ExpectedCellValue(sheetName, 35, 9, .1501),
					new ExpectedCellValue(sheetName, 35, 10, 7),

					// March
					new ExpectedCellValue(sheetName, 38, 3, .0567),
					new ExpectedCellValue(sheetName, 38, 4, .4034),
					new ExpectedCellValue(sheetName, 38, 5, 0),
					new ExpectedCellValue(sheetName, 38, 6, 1),
					new ExpectedCellValue(sheetName, 38, 7, 2),
					new ExpectedCellValue(sheetName, 38, 8, null),
					new ExpectedCellValue(sheetName, 38, 9, .2219),
					new ExpectedCellValue(sheetName, 38, 10, 3),
				});

				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable3"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					// Show 'Wholesale Price' data as the percentage of its parent row.
					wholesalePriceDataField.ShowDataAs = ShowDataAs.PercentOfParentRow;
					unitsSoldDataField.ShowDataAs = ShowDataAs.NoCalculation;
					foreach (var field in pivotTable.Fields)
					{
						field.SubTotalFunctions = eSubTotalFunctions.None;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B30:J41"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				validateWorksheet();
				// Validate subtotal top is off.
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					// January
					new ExpectedCellValue(sheetName, 33, 3, null),
					new ExpectedCellValue(sheetName, 33, 4, null),
					new ExpectedCellValue(sheetName, 33, 5, null),
					new ExpectedCellValue(sheetName, 33, 6, null),
					new ExpectedCellValue(sheetName, 33, 7, null),
					new ExpectedCellValue(sheetName, 33, 8, null),
					new ExpectedCellValue(sheetName, 33, 9, null),
					new ExpectedCellValue(sheetName, 33, 10, null),

					// February
					new ExpectedCellValue(sheetName, 35, 3, null),
					new ExpectedCellValue(sheetName, 35, 4, null),
					new ExpectedCellValue(sheetName, 35, 5, null),
					new ExpectedCellValue(sheetName, 35, 6, null),
					new ExpectedCellValue(sheetName, 35, 7, null),
					new ExpectedCellValue(sheetName, 35, 8, null),
					new ExpectedCellValue(sheetName, 35, 9, null),
					new ExpectedCellValue(sheetName, 35, 10, null),

					// March
					new ExpectedCellValue(sheetName, 38, 3, null),
					new ExpectedCellValue(sheetName, 38, 4, null),
					new ExpectedCellValue(sheetName, 38, 5, null),
					new ExpectedCellValue(sheetName, 38, 6, null),
					new ExpectedCellValue(sheetName, 38, 7, null),
					new ExpectedCellValue(sheetName, 38, 8, null),
					new ExpectedCellValue(sheetName, 38, 9, null),
					new ExpectedCellValue(sheetName, 38, 10, null),
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableShowDataAs.xlsx")]
		public void PivotTableRefreshShowDataAsColumnFieldsRowFieldsAndRowDataFieldsPercentOfParentRowDataFieldSubtotalTopOnOff()
		{
			var file = new FileInfo("PivotTableShowDataAs.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				void validateWorksheet() => TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
					{
						new ExpectedCellValue(sheetName, 47, 2, "Row Labels"),
						new ExpectedCellValue(sheetName, 47, 3, "Chicago"),
						new ExpectedCellValue(sheetName, 47, 4, "Nashville"),
						new ExpectedCellValue(sheetName, 47, 5, "San Francisco"),
						new ExpectedCellValue(sheetName, 47, 6, "Grand Total"),

						new ExpectedCellValue(sheetName, 48, 2, "Sum of Units Sold"),
						new ExpectedCellValue(sheetName, 48, 3, null),
						new ExpectedCellValue(sheetName, 48, 6, null),

						new ExpectedCellValue(sheetName, 50, 2, "Car Rack"),
						new ExpectedCellValue(sheetName, 50, 3, 1),
						new ExpectedCellValue(sheetName, 50, 4, 1),
						new ExpectedCellValue(sheetName, 50, 5, 1),
						new ExpectedCellValue(sheetName, 50, 6, 1),

						new ExpectedCellValue(sheetName, 52, 2, "Sleeping Bag"),
						new ExpectedCellValue(sheetName, 52, 3, null),
						new ExpectedCellValue(sheetName, 52, 4, 0),
						new ExpectedCellValue(sheetName, 52, 5, 1),
						new ExpectedCellValue(sheetName, 52, 6, .1429),

						new ExpectedCellValue(sheetName, 53, 2, "Tent"),
						new ExpectedCellValue(sheetName, 53, 3, null),
						new ExpectedCellValue(sheetName, 53, 4, 1),
						new ExpectedCellValue(sheetName, 53, 5, 0),
						new ExpectedCellValue(sheetName, 53, 6, .8571),

						new ExpectedCellValue(sheetName, 55, 2, "Car Rack"),
						new ExpectedCellValue(sheetName, 55, 3, 0),
						new ExpectedCellValue(sheetName, 55, 4, 1),
						new ExpectedCellValue(sheetName, 55, 5, null),
						new ExpectedCellValue(sheetName, 55, 6, .6667),

						new ExpectedCellValue(sheetName, 56, 2, "Headlamp"),
						new ExpectedCellValue(sheetName, 56, 3, 1),
						new ExpectedCellValue(sheetName, 56, 4, 0),
						new ExpectedCellValue(sheetName, 56, 5, null),
						new ExpectedCellValue(sheetName, 56, 6, .3333),

						new ExpectedCellValue(sheetName, 57, 2, "Sum of Wholesale Price"),
						new ExpectedCellValue(sheetName, 57, 3, null),
						new ExpectedCellValue(sheetName, 57, 6, null),

						new ExpectedCellValue(sheetName, 59, 2, "Car Rack"),
						new ExpectedCellValue(sheetName, 59, 3, 415.75),
						new ExpectedCellValue(sheetName, 59, 4, 415.75),
						new ExpectedCellValue(sheetName, 59, 5, 415.75),
						new ExpectedCellValue(sheetName, 59, 6, 1247.25),

						new ExpectedCellValue(sheetName, 61, 2, "Sleeping Bag"),
						new ExpectedCellValue(sheetName, 61, 3, null),
						new ExpectedCellValue(sheetName, 61, 4, null),
						new ExpectedCellValue(sheetName, 61, 5, 99),
						new ExpectedCellValue(sheetName, 61, 6, 99),
						new ExpectedCellValue(sheetName, 62, 2, "Tent"),
						new ExpectedCellValue(sheetName, 62, 3, null),
						new ExpectedCellValue(sheetName, 62, 4, 199),
						new ExpectedCellValue(sheetName, 62, 5, null),
						new ExpectedCellValue(sheetName, 62, 6, 199),

						new ExpectedCellValue(sheetName, 64, 2, "Car Rack"),
						new ExpectedCellValue(sheetName, 64, 3, null),
						new ExpectedCellValue(sheetName, 64, 4, 415.75),
						new ExpectedCellValue(sheetName, 64, 5, null),
						new ExpectedCellValue(sheetName, 64, 6, 415.75),
						new ExpectedCellValue(sheetName, 65, 2, "Headlamp"),
						new ExpectedCellValue(sheetName, 65, 3, 24.99),
						new ExpectedCellValue(sheetName, 65, 4, null),
						new ExpectedCellValue(sheetName, 65, 5, null),
						new ExpectedCellValue(sheetName, 65, 6, 24.99),
						new ExpectedCellValue(sheetName, 66, 2, "Total Sum of Units Sold"),
						new ExpectedCellValue(sheetName, 66, 3, 1),
						new ExpectedCellValue(sheetName, 66, 4, 1),
						new ExpectedCellValue(sheetName, 66, 5, 1),
						new ExpectedCellValue(sheetName, 66, 6, 1),
						new ExpectedCellValue(sheetName, 67, 2, "Total Sum of Wholesale Price"),
						new ExpectedCellValue(sheetName, 67, 3, 440.74),
						new ExpectedCellValue(sheetName, 67, 4, 1030.5),
						new ExpectedCellValue(sheetName, 67, 5, 514.75),
						new ExpectedCellValue(sheetName, 67, 6, 1985.99)
					});
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable4"];
					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					// Show 'Units Sold' data as the percentage of its parent row.
					unitsSoldDataField.ShowDataAs = ShowDataAs.PercentOfParentRow;
					wholesalePriceDataField.ShowDataAs = ShowDataAs.NoCalculation;
					foreach (var field in pivotTable.Fields)
					{
						field.SubTotalFunctions = eSubTotalFunctions.Default;
						field.SubtotalLocation = SubtotalLocation.Top;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B46:F67"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				validateWorksheet();
				// Validate subtotal top is on.
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					// January
					new ExpectedCellValue(sheetName, 49, 3, .6667),
					new ExpectedCellValue(sheetName, 49, 4, .2),
					new ExpectedCellValue(sheetName, 49, 5, .5),
					new ExpectedCellValue(sheetName, 49, 6, .3333),
					// February
					new ExpectedCellValue(sheetName, 51, 3, 0),
					new ExpectedCellValue(sheetName, 51, 4, .6),
					new ExpectedCellValue(sheetName, 51, 5, .5),
					new ExpectedCellValue(sheetName, 51, 6, .4667),
					// March
					new ExpectedCellValue(sheetName, 54, 3, .3333),
					new ExpectedCellValue(sheetName, 54, 4, .2),
					new ExpectedCellValue(sheetName, 54, 5, 0),
					new ExpectedCellValue(sheetName, 54, 6, .2),
					// January
					new ExpectedCellValue(sheetName, 58, 3, 415.75),
					new ExpectedCellValue(sheetName, 58, 4, 415.75),
					new ExpectedCellValue(sheetName, 58, 5, 415.75),
					new ExpectedCellValue(sheetName, 58, 6, 1247.25),
					// February
					new ExpectedCellValue(sheetName, 60, 3, null),
					new ExpectedCellValue(sheetName, 60, 4, 199),
					new ExpectedCellValue(sheetName, 60, 5, 99),
					new ExpectedCellValue(sheetName, 60, 6, 298),
					// March
					new ExpectedCellValue(sheetName, 63, 3, 24.99),
					new ExpectedCellValue(sheetName, 63, 4, 415.75),
					new ExpectedCellValue(sheetName, 63, 5, null),
					new ExpectedCellValue(sheetName, 63, 6, 440.74)
				});

				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable4"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					// Show 'Units Sold' data as the percentage of its parent row.
					unitsSoldDataField.ShowDataAs = ShowDataAs.PercentOfParentRow;
					wholesalePriceDataField.ShowDataAs = ShowDataAs.NoCalculation;
					foreach (var field in pivotTable.Fields)
					{
						field.SubTotalFunctions = eSubTotalFunctions.None;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B46:F67"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				validateWorksheet();
				// Validate subtotal top is off.
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					// January
					new ExpectedCellValue(sheetName, 49, 3, null),
					new ExpectedCellValue(sheetName, 49, 4, null),
					new ExpectedCellValue(sheetName, 49, 5, null),
					new ExpectedCellValue(sheetName, 49, 6, null),
					// February
					new ExpectedCellValue(sheetName, 51, 3, null),
					new ExpectedCellValue(sheetName, 51, 4, null),
					new ExpectedCellValue(sheetName, 51, 5, null),
					new ExpectedCellValue(sheetName, 51, 6, null),
					// March
					new ExpectedCellValue(sheetName, 54, 3, null),
					new ExpectedCellValue(sheetName, 54, 4, null),
					new ExpectedCellValue(sheetName, 54, 5, null),
					new ExpectedCellValue(sheetName, 54, 6, null),
					// January
					new ExpectedCellValue(sheetName, 58, 3, null),
					new ExpectedCellValue(sheetName, 58, 4, null),
					new ExpectedCellValue(sheetName, 58, 5, null),
					new ExpectedCellValue(sheetName, 58, 6, null),
					// February
					new ExpectedCellValue(sheetName, 60, 3, null),
					new ExpectedCellValue(sheetName, 60, 4, null),
					new ExpectedCellValue(sheetName, 60, 5, null),
					new ExpectedCellValue(sheetName, 60, 6, null),
					// March
					new ExpectedCellValue(sheetName, 63, 3, null),
					new ExpectedCellValue(sheetName, 63, 4, null),
					new ExpectedCellValue(sheetName, 63, 5, null),
					new ExpectedCellValue(sheetName, 63, 6, null)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableShowDataAs.xlsx")]
		public void PivotTableRefreshShowDataAsColumnFieldsRowFieldsAndRowDataFieldsPercentOfParentRowDataFieldSubtotalBottom()
		{
			var file = new FileInfo("PivotTableShowDataAs.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable4"];
					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					// Show 'Units Sold' data as the percentage of its parent row.
					unitsSoldDataField.ShowDataAs = ShowDataAs.PercentOfParentRow;
					wholesalePriceDataField.ShowDataAs = ShowDataAs.NoCalculation;
					foreach (var field in pivotTable.Fields)
					{
						field.SubTotalFunctions = eSubTotalFunctions.Default;
						field.SubtotalLocation = SubtotalLocation.Bottom;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B46:F73"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 47, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 47, 3, "Chicago"),
					new ExpectedCellValue(sheetName, 47, 4, "Nashville"),
					new ExpectedCellValue(sheetName, 47, 5, "San Francisco"),
					new ExpectedCellValue(sheetName, 47, 6, "Grand Total"),

					new ExpectedCellValue(sheetName, 48, 2, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 48, 3, null),
					new ExpectedCellValue(sheetName, 48, 6, null),

					new ExpectedCellValue(sheetName, 49, 2, "January"),
					new ExpectedCellValue(sheetName, 49, 3, null),
					new ExpectedCellValue(sheetName, 49, 4, null),
					new ExpectedCellValue(sheetName, 49, 5, null),
					new ExpectedCellValue(sheetName, 49, 6, null),

					new ExpectedCellValue(sheetName, 50, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 50, 3, 1),
					new ExpectedCellValue(sheetName, 50, 4, 1),
					new ExpectedCellValue(sheetName, 50, 5, 1),
					new ExpectedCellValue(sheetName, 50, 6, 1),

					new ExpectedCellValue(sheetName, 51, 2, "January Total"),
					new ExpectedCellValue(sheetName, 51, 3, .6667),
					new ExpectedCellValue(sheetName, 51, 4, .2),
					new ExpectedCellValue(sheetName, 51, 5, .5),
					new ExpectedCellValue(sheetName, 51, 6, .3333),

					new ExpectedCellValue(sheetName, 52, 2, "February"),
					new ExpectedCellValue(sheetName, 52, 3, null),
					new ExpectedCellValue(sheetName, 52, 4, null),
					new ExpectedCellValue(sheetName, 52, 5, null),
					new ExpectedCellValue(sheetName, 52, 6, null),

					new ExpectedCellValue(sheetName, 53, 2, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 53, 3, null),
					new ExpectedCellValue(sheetName, 53, 4, 0),
					new ExpectedCellValue(sheetName, 53, 5, 1),
					new ExpectedCellValue(sheetName, 53, 6, .1429),

					new ExpectedCellValue(sheetName, 54, 2, "Tent"),
					new ExpectedCellValue(sheetName, 54, 3, null),
					new ExpectedCellValue(sheetName, 54, 4, 1),
					new ExpectedCellValue(sheetName, 54, 5, 0),
					new ExpectedCellValue(sheetName, 54, 6, .8571),

					new ExpectedCellValue(sheetName, 55, 2, "February Total"),
					new ExpectedCellValue(sheetName, 55, 3, 0),
					new ExpectedCellValue(sheetName, 55, 4, .6),
					new ExpectedCellValue(sheetName, 55, 5, .5),
					new ExpectedCellValue(sheetName, 55, 6, .4667),

					new ExpectedCellValue(sheetName, 56, 2, "March"),
					new ExpectedCellValue(sheetName, 56, 3, null),
					new ExpectedCellValue(sheetName, 56, 4, null),
					new ExpectedCellValue(sheetName, 56, 5, null),
					new ExpectedCellValue(sheetName, 56, 6, null),

					new ExpectedCellValue(sheetName, 57, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 57, 3, 0),
					new ExpectedCellValue(sheetName, 57, 4, 1),
					new ExpectedCellValue(sheetName, 57, 5, null),
					new ExpectedCellValue(sheetName, 57, 6, .6667),

					new ExpectedCellValue(sheetName, 58, 2, "Headlamp"),
					new ExpectedCellValue(sheetName, 58, 3, 1),
					new ExpectedCellValue(sheetName, 58, 4, 0),
					new ExpectedCellValue(sheetName, 58, 5, null),
					new ExpectedCellValue(sheetName, 58, 6, .3333),

					new ExpectedCellValue(sheetName, 59, 2, "March Total"),
					new ExpectedCellValue(sheetName, 59, 3, .3333),
					new ExpectedCellValue(sheetName, 59, 4, .2),
					new ExpectedCellValue(sheetName, 59, 5, 0),
					new ExpectedCellValue(sheetName, 59, 6, .2),

					new ExpectedCellValue(sheetName, 60, 2, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 60, 3, null),
					new ExpectedCellValue(sheetName, 60, 6, null),

					new ExpectedCellValue(sheetName, 61, 2, "January"),
					new ExpectedCellValue(sheetName, 61, 3, null),
					new ExpectedCellValue(sheetName, 61, 6, null),

					new ExpectedCellValue(sheetName, 62, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 62, 3, 415.75),
					new ExpectedCellValue(sheetName, 62, 4, 415.75),
					new ExpectedCellValue(sheetName, 62, 5, 415.75),
					new ExpectedCellValue(sheetName, 62, 6, 1247.25),

					new ExpectedCellValue(sheetName, 63, 2, "January Total"),
					new ExpectedCellValue(sheetName, 63, 3, 415.75),
					new ExpectedCellValue(sheetName, 63, 4, 415.75),
					new ExpectedCellValue(sheetName, 63, 5, 415.75),
					new ExpectedCellValue(sheetName, 63, 6, 1247.25),

					new ExpectedCellValue(sheetName, 64, 2, "February"),
					new ExpectedCellValue(sheetName, 64, 3, null),
					new ExpectedCellValue(sheetName, 64, 6, null),

					new ExpectedCellValue(sheetName, 65, 2, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 65, 3, null),
					new ExpectedCellValue(sheetName, 65, 4, null),
					new ExpectedCellValue(sheetName, 65, 5, 99),
					new ExpectedCellValue(sheetName, 65, 6, 99),

					new ExpectedCellValue(sheetName, 66, 2, "Tent"),
					new ExpectedCellValue(sheetName, 66, 3, null),
					new ExpectedCellValue(sheetName, 66, 4, 199),
					new ExpectedCellValue(sheetName, 66, 5, null),
					new ExpectedCellValue(sheetName, 66, 6, 199),

					new ExpectedCellValue(sheetName, 67, 2, "February Total"),
					new ExpectedCellValue(sheetName, 67, 3, null),
					new ExpectedCellValue(sheetName, 67, 4, 199),
					new ExpectedCellValue(sheetName, 67, 5, 99),
					new ExpectedCellValue(sheetName, 67, 6, 298),

					new ExpectedCellValue(sheetName, 68, 2, "March"),
					new ExpectedCellValue(sheetName, 68, 3, null),
					new ExpectedCellValue(sheetName, 68, 6, null),

					new ExpectedCellValue(sheetName, 69, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 69, 3, null),
					new ExpectedCellValue(sheetName, 69, 4, 415.75),
					new ExpectedCellValue(sheetName, 69, 5, null),
					new ExpectedCellValue(sheetName, 69, 6, 415.75),

					new ExpectedCellValue(sheetName, 70, 2, "Headlamp"),
					new ExpectedCellValue(sheetName, 70, 3, 24.99),
					new ExpectedCellValue(sheetName, 70, 4, null),
					new ExpectedCellValue(sheetName, 70, 5, null),
					new ExpectedCellValue(sheetName, 70, 6, 24.99),

					new ExpectedCellValue(sheetName, 71, 2, "March Total"),
					new ExpectedCellValue(sheetName, 71, 3, 24.99),
					new ExpectedCellValue(sheetName, 71, 4, 415.75),
					new ExpectedCellValue(sheetName, 71, 5, null),
					new ExpectedCellValue(sheetName, 71, 6, 440.74),

					new ExpectedCellValue(sheetName, 72, 2, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 72, 3, 1),
					new ExpectedCellValue(sheetName, 72, 4, 1),
					new ExpectedCellValue(sheetName, 72, 5, 1),
					new ExpectedCellValue(sheetName, 72, 6, 1),
					new ExpectedCellValue(sheetName, 73, 2, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 73, 3, 440.74),
					new ExpectedCellValue(sheetName, 73, 4, 1030.5),
					new ExpectedCellValue(sheetName, 73, 5, 514.75),
					new ExpectedCellValue(sheetName, 73, 6, 1985.99)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableShowDataAs.xlsx")]
		public void PivotTableRefreshShowDataAsPercentOfParentRowWithColumnFieldsAndRowDataFieldsSubtotalsTopAndBottom()
		{
			var file = new FileInfo("PivotTableShowDataAs.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				void validateWorksheet() => TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 19, 2, null),
					new ExpectedCellValue(sheetName, 19, 3, "January"),
					new ExpectedCellValue(sheetName, 19, 4, "January Total"),
					new ExpectedCellValue(sheetName, 19, 5, "February"),
					new ExpectedCellValue(sheetName, 19, 6, null),
					new ExpectedCellValue(sheetName, 19, 7, "February Total"),
					new ExpectedCellValue(sheetName, 19, 8, "March"),
					new ExpectedCellValue(sheetName, 19, 9, null),
					new ExpectedCellValue(sheetName, 19, 10, "March Total"),
					new ExpectedCellValue(sheetName, 19, 11, "Grand Total"),

					new ExpectedCellValue(sheetName, 20, 2, "Values"),
					new ExpectedCellValue(sheetName, 20, 3, "Car Rack"),
					new ExpectedCellValue(sheetName, 20, 4, null),
					new ExpectedCellValue(sheetName, 20, 5, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 20, 6, "Tent"),
					new ExpectedCellValue(sheetName, 20, 7, null),
					new ExpectedCellValue(sheetName, 20, 8, "Car Rack"),
					new ExpectedCellValue(sheetName, 20, 9, "Headlamp"),
					new ExpectedCellValue(sheetName, 20, 10, null),
					new ExpectedCellValue(sheetName, 20, 11, null),

					new ExpectedCellValue(sheetName, 21, 2, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 21, 3, null),
					new ExpectedCellValue(sheetName, 21, 4, null),
					new ExpectedCellValue(sheetName, 21, 5, null),
					new ExpectedCellValue(sheetName, 21, 6, null),
					new ExpectedCellValue(sheetName, 21, 7, null),
					new ExpectedCellValue(sheetName, 21, 8, null),
					new ExpectedCellValue(sheetName, 21, 9, null),
					new ExpectedCellValue(sheetName, 21, 10, null),
					new ExpectedCellValue(sheetName, 21, 11, null),

					new ExpectedCellValue(sheetName, 22, 2, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 22, 3, 5),
					new ExpectedCellValue(sheetName, 22, 4, 5),
					new ExpectedCellValue(sheetName, 22, 5, 1),
					new ExpectedCellValue(sheetName, 22, 6, 6),
					new ExpectedCellValue(sheetName, 22, 7, 7),
					new ExpectedCellValue(sheetName, 22, 8, 2),
					new ExpectedCellValue(sheetName, 22, 9, 1),
					new ExpectedCellValue(sheetName, 22, 10, 3),
					new ExpectedCellValue(sheetName, 22, 11, 15)
				});

				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable2"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					// Show 'Wholesale Price' data as the percentage of its parent row.
					wholesalePriceDataField.ShowDataAs = ShowDataAs.PercentOfParentRow;
					unitsSoldDataField.ShowDataAs = ShowDataAs.NoCalculation;
					foreach (var field in pivotTable.Fields)
					{
						field.SubtotalLocation = SubtotalLocation.Top;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B18:K22"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				validateWorksheet();

				// Verify that subtotal top provides the same result.
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable2"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					// Show 'Wholesale Price' data as the percentage of its parent row.
					wholesalePriceDataField.ShowDataAs = ShowDataAs.PercentOfParentRow;
					unitsSoldDataField.ShowDataAs = ShowDataAs.NoCalculation;
					foreach (var field in pivotTable.Fields)
					{
						field.SubtotalLocation = SubtotalLocation.Bottom;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B18:K22"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				validateWorksheet();
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableShowDataAs.xlsx")]
		public void PivotTableRefreshShowDataAsPercentOfParentRowWithColumnFieldsAndRowDataFieldsSubtotalsOff()
		{
			var file = new FileInfo("PivotTableShowDataAs.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable2"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					// Show 'Wholesale Price' data as the percentage of its parent row.
					wholesalePriceDataField.ShowDataAs = ShowDataAs.PercentOfParentRow;
					unitsSoldDataField.ShowDataAs = ShowDataAs.NoCalculation;
					foreach (var field in pivotTable.Fields)
					{
						field.SubtotalLocation = SubtotalLocation.Off;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B18:H22"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}

				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 19, 2, null),
					new ExpectedCellValue(sheetName, 19, 3, "January"),
					new ExpectedCellValue(sheetName, 19, 4, "February"),
					new ExpectedCellValue(sheetName, 19, 5, null),
					new ExpectedCellValue(sheetName, 19, 6, "March"),
					new ExpectedCellValue(sheetName, 19, 7, null),
					new ExpectedCellValue(sheetName, 19, 8, "Grand Total"),

					new ExpectedCellValue(sheetName, 20, 2, "Values"),
					new ExpectedCellValue(sheetName, 20, 3, "Car Rack"),
					new ExpectedCellValue(sheetName, 20, 4, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 20, 5, "Tent"),
					new ExpectedCellValue(sheetName, 20, 6, "Car Rack"),
					new ExpectedCellValue(sheetName, 20, 7, "Headlamp"),
					new ExpectedCellValue(sheetName, 20, 8, null),

					new ExpectedCellValue(sheetName, 21, 2, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 21, 3, null),
					new ExpectedCellValue(sheetName, 21, 4, null),
					new ExpectedCellValue(sheetName, 21, 5, null),
					new ExpectedCellValue(sheetName, 21, 6, null),
					new ExpectedCellValue(sheetName, 21, 7, null),
					new ExpectedCellValue(sheetName, 21, 8, null),

					new ExpectedCellValue(sheetName, 22, 2, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 22, 3, 5),
					new ExpectedCellValue(sheetName, 22, 4, 1),
					new ExpectedCellValue(sheetName, 22, 5, 6),
					new ExpectedCellValue(sheetName, 22, 6, 2),
					new ExpectedCellValue(sheetName, 22, 7, 1),
					new ExpectedCellValue(sheetName, 22, 8, 15)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableTabularShowDataAs.xlsx")]
		public void PivotTableRefreshShowDataAsPercentOfParentRowWithColumnFieldsAndRowDataFieldsTabularForm()
		{
			var file = new FileInfo("PivotTableTabularShowDataAs.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable1"];
					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					// Show 'Wholesale Price' data as the percentage of its parent row.
					wholesalePriceDataField.ShowDataAs = ShowDataAs.PercentOfParentRow;
					unitsSoldDataField.ShowDataAs = ShowDataAs.NoCalculation;

					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B2:E13"), pivotTable.Address);
					package.SaveAs(newFile.File);
				}

				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 2, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 2, 3, "Month"),
					new ExpectedCellValue(sheetName, 2, 4, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 2, 5, "Sum of Units Sold"),

					new ExpectedCellValue(sheetName, 3, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 3, 3, "January"),
					new ExpectedCellValue(sheetName, 3, 4, .9433),
					new ExpectedCellValue(sheetName, 3, 5, 2),

					new ExpectedCellValue(sheetName, 4, 2, null),
					new ExpectedCellValue(sheetName, 4, 3, "March"),
					new ExpectedCellValue(sheetName, 4, 4, .0567),
					new ExpectedCellValue(sheetName, 4, 5, 1),

					new ExpectedCellValue(sheetName, 5, 2, "Chicago Total"),
					new ExpectedCellValue(sheetName, 5, 3, null),
					new ExpectedCellValue(sheetName, 5, 4, .2219),
					new ExpectedCellValue(sheetName, 5, 5, 3),

					new ExpectedCellValue(sheetName, 6, 2, "Nashville"),
					new ExpectedCellValue(sheetName, 6, 3, "January"),
					new ExpectedCellValue(sheetName, 6, 4, .4034),
					new ExpectedCellValue(sheetName, 6, 5, 2),

					new ExpectedCellValue(sheetName, 7, 2, null),
					new ExpectedCellValue(sheetName, 7, 3, "February"),
					new ExpectedCellValue(sheetName, 7, 4, .1931),
					new ExpectedCellValue(sheetName, 7, 5, 6),

					new ExpectedCellValue(sheetName, 8, 2, null),
					new ExpectedCellValue(sheetName, 8, 3, "March"),
					new ExpectedCellValue(sheetName, 8, 4, .4034),
					new ExpectedCellValue(sheetName, 8, 5, 2),

					new ExpectedCellValue(sheetName, 9, 2, "Nashville Total"),
					new ExpectedCellValue(sheetName, 9, 3, null),
					new ExpectedCellValue(sheetName, 9, 4, .5189),
					new ExpectedCellValue(sheetName, 9, 5, 10),

					new ExpectedCellValue(sheetName, 10, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 10, 3, "January"),
					new ExpectedCellValue(sheetName, 10, 4, .8077),
					new ExpectedCellValue(sheetName, 10, 5, 1),

					new ExpectedCellValue(sheetName, 11, 2, null),
					new ExpectedCellValue(sheetName, 11, 3, "February"),
					new ExpectedCellValue(sheetName, 11, 4, .1923),
					new ExpectedCellValue(sheetName, 11, 5, 1),

					new ExpectedCellValue(sheetName, 12, 2, "San Francisco Total"),
					new ExpectedCellValue(sheetName, 12, 3, null),
					new ExpectedCellValue(sheetName, 12, 4, .2592),
					new ExpectedCellValue(sheetName, 12, 5, 2),

					new ExpectedCellValue(sheetName, 13, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 13, 3, null),
					new ExpectedCellValue(sheetName, 13, 4, 1),
					new ExpectedCellValue(sheetName, 13, 5, 15),
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableMixedDataFieldFunctionTypes.xlsx")]
		public void PivotTableRefreshDataFieldCountFunctionAndShowDataAsPercentOfParentRowTotal()
		{
			var file = new FileInfo("PivotTableMixedDataFieldFunctionTypes.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable1"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("A1:K15"), pivotTable.Address);
					Assert.AreEqual(16, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 1, 1, null),
					new ExpectedCellValue(sheetName, 2, 1, null),
					new ExpectedCellValue(sheetName, 3, 1, "Values"),
					new ExpectedCellValue(sheetName, 4, 1, "Count of Net Change"),
					new ExpectedCellValue(sheetName, 5, 1, null),
					new ExpectedCellValue(sheetName, 6, 1, null),
					new ExpectedCellValue(sheetName, 7, 1, "Count of Balance at End Date"),
					new ExpectedCellValue(sheetName, 8, 1, null),
					new ExpectedCellValue(sheetName, 9, 1, null),
					new ExpectedCellValue(sheetName, 10, 1, "Count of Balance at End date (prior year)"),
					new ExpectedCellValue(sheetName, 11, 1, null),
					new ExpectedCellValue(sheetName, 12, 1, null),
					new ExpectedCellValue(sheetName, 13, 1, "Total Count of Net Change"),
					new ExpectedCellValue(sheetName, 14, 1, "Total Count of Balance at End Date"),
					new ExpectedCellValue(sheetName, 15, 1, "Total Count of Balance at End date (prior year)"),
					new ExpectedCellValue(sheetName, 1, 2, null),
					new ExpectedCellValue(sheetName, 2, 2, null),
					new ExpectedCellValue(sheetName, 3, 2, "Account Type"),
					new ExpectedCellValue(sheetName, 4, 2, "Begin-Total"),
					new ExpectedCellValue(sheetName, 5, 2, "End-Total"),
					new ExpectedCellValue(sheetName, 6, 2, "Posting"),
					new ExpectedCellValue(sheetName, 7, 2, "Begin-Total"),
					new ExpectedCellValue(sheetName, 8, 2, "End-Total"),
					new ExpectedCellValue(sheetName, 9, 2, "Posting"),
					new ExpectedCellValue(sheetName, 10, 2, "Begin-Total"),
					new ExpectedCellValue(sheetName, 11, 2, "End-Total"),
					new ExpectedCellValue(sheetName, 12, 2, "Posting"),
					new ExpectedCellValue(sheetName, 13, 2, null),
					new ExpectedCellValue(sheetName, 14, 2, null),
					new ExpectedCellValue(sheetName, 15, 2, null),
					new ExpectedCellValue(sheetName, 1, 3, "Quarters"),
					new ExpectedCellValue(sheetName, 2, 3, "Qtr1"),
					new ExpectedCellValue(sheetName, 3, 3, "FALSE"),
					new ExpectedCellValue(sheetName, 4, 3, 0),
					new ExpectedCellValue(sheetName, 5, 3, 1),
					new ExpectedCellValue(sheetName, 6, 3, 0),
					new ExpectedCellValue(sheetName, 7, 3, 0),
					new ExpectedCellValue(sheetName, 8, 3, 1),
					new ExpectedCellValue(sheetName, 9, 3, 0),
					new ExpectedCellValue(sheetName, 10, 3, 0),
					new ExpectedCellValue(sheetName, 11, 3, 1),
					new ExpectedCellValue(sheetName, 12, 3, 0),
					new ExpectedCellValue(sheetName, 13, 3, 1d),
					new ExpectedCellValue(sheetName, 14, 3, 1d),
					new ExpectedCellValue(sheetName, 15, 3, 1d),
					new ExpectedCellValue(sheetName, 1, 4, "Blocked"),
					new ExpectedCellValue(sheetName, 2, 4, null),
					new ExpectedCellValue(sheetName, 3, 4, "TRUE"),
					new ExpectedCellValue(sheetName, 4, 4, 0.5),
					new ExpectedCellValue(sheetName, 5, 4, 0.25),
					new ExpectedCellValue(sheetName, 6, 4, 0.25),
					new ExpectedCellValue(sheetName, 7, 4, 0.5),
					new ExpectedCellValue(sheetName, 8, 4, 0.25),
					new ExpectedCellValue(sheetName, 9, 4, 0.25),
					new ExpectedCellValue(sheetName, 10, 4, 0.5),
					new ExpectedCellValue(sheetName, 11, 4, 0.25),
					new ExpectedCellValue(sheetName, 12, 4, 0.25),
					new ExpectedCellValue(sheetName, 13, 4, 1d),
					new ExpectedCellValue(sheetName, 14, 4, 1d),
					new ExpectedCellValue(sheetName, 15, 4, 1d),
					new ExpectedCellValue(sheetName, 1, 5, null),
					new ExpectedCellValue(sheetName, 2, 5, "Qtr2"),
					new ExpectedCellValue(sheetName, 3, 5, "FALSE"),
					new ExpectedCellValue(sheetName, 4, 5, 0.33),
					new ExpectedCellValue(sheetName, 5, 5, 0.33),
					new ExpectedCellValue(sheetName, 6, 5, 0.33),
					new ExpectedCellValue(sheetName, 7, 5, 0.33),
					new ExpectedCellValue(sheetName, 8, 5, 0.33),
					new ExpectedCellValue(sheetName, 9, 5, 0.33),
					new ExpectedCellValue(sheetName, 10, 5, 0.33),
					new ExpectedCellValue(sheetName, 11, 5, 0.33),
					new ExpectedCellValue(sheetName, 12, 5, 0.33),
					new ExpectedCellValue(sheetName, 13, 5, 1d),
					new ExpectedCellValue(sheetName, 14, 5, 1d),
					new ExpectedCellValue(sheetName, 15, 5, 1d),
					new ExpectedCellValue(sheetName, 1, 6, null),
					new ExpectedCellValue(sheetName, 2, 6, null),
					new ExpectedCellValue(sheetName, 3, 6, "TRUE"),
					new ExpectedCellValue(sheetName, 4, 6, 0),
					new ExpectedCellValue(sheetName, 5, 6, 1d),
					new ExpectedCellValue(sheetName, 6, 6, 0),
					new ExpectedCellValue(sheetName, 7, 6, 0),
					new ExpectedCellValue(sheetName, 8, 6, 1d),
					new ExpectedCellValue(sheetName, 9, 6, 0),
					new ExpectedCellValue(sheetName, 10, 6, 0),
					new ExpectedCellValue(sheetName, 11, 6, 1d),
					new ExpectedCellValue(sheetName, 12, 6, 0),
					new ExpectedCellValue(sheetName, 13, 6, 1d),
					new ExpectedCellValue(sheetName, 14, 6, 1d),
					new ExpectedCellValue(sheetName, 15, 6, 1d),
					new ExpectedCellValue(sheetName, 1, 7, null),
					new ExpectedCellValue(sheetName, 2, 7, "Qtr3"),
					new ExpectedCellValue(sheetName, 3, 7, "FALSE"),
					new ExpectedCellValue(sheetName, 4, 7, 0.5),
					new ExpectedCellValue(sheetName, 5, 7, 0.5),
					new ExpectedCellValue(sheetName, 6, 7, 0),
					new ExpectedCellValue(sheetName, 7, 7, 1d),
					new ExpectedCellValue(sheetName, 8, 7, 0),
					new ExpectedCellValue(sheetName, 9, 7, 0),
					new ExpectedCellValue(sheetName, 10, 7, 0.5),
					new ExpectedCellValue(sheetName, 11, 7, 0.5),
					new ExpectedCellValue(sheetName, 12, 7, 0),
					new ExpectedCellValue(sheetName, 13, 7, 1d),
					new ExpectedCellValue(sheetName, 14, 7, 1d),
					new ExpectedCellValue(sheetName, 15, 7, 1d),
					new ExpectedCellValue(sheetName, 1, 8, null),
					new ExpectedCellValue(sheetName, 2, 8, null),
					new ExpectedCellValue(sheetName, 3, 8, "TRUE"),
					new ExpectedCellValue(sheetName, 4, 8, 0.33),
					new ExpectedCellValue(sheetName, 5, 8, 0.33),
					new ExpectedCellValue(sheetName, 6, 8, 0.33),
					new ExpectedCellValue(sheetName, 7, 8, 0.33),
					new ExpectedCellValue(sheetName, 8, 8, 0.33),
					new ExpectedCellValue(sheetName, 9, 8, 0.33),
					new ExpectedCellValue(sheetName, 10, 8, 0.33),
					new ExpectedCellValue(sheetName, 11, 8, 0.33),
					new ExpectedCellValue(sheetName, 12, 8, 0.33),
					new ExpectedCellValue(sheetName, 13, 8, 1d),
					new ExpectedCellValue(sheetName, 14, 8, 1d),
					new ExpectedCellValue(sheetName, 15, 8, 1d),
					new ExpectedCellValue(sheetName, 1, 9, null),
					new ExpectedCellValue(sheetName, 2, 9, "Qtr4"),
					new ExpectedCellValue(sheetName, 3, 9, "FALSE"),
					new ExpectedCellValue(sheetName, 4, 9, 0.1667),
					new ExpectedCellValue(sheetName, 5, 9, 0),
					new ExpectedCellValue(sheetName, 6, 9, 0.8333),
					new ExpectedCellValue(sheetName, 7, 9, 0.1667),
					new ExpectedCellValue(sheetName, 8, 9, 0),
					new ExpectedCellValue(sheetName, 9, 9, 0.8333),
					new ExpectedCellValue(sheetName, 10, 9, 0.1667),
					new ExpectedCellValue(sheetName, 11, 9, 0),
					new ExpectedCellValue(sheetName, 12, 9, 0.8333),
					new ExpectedCellValue(sheetName, 13, 9, 1d),
					new ExpectedCellValue(sheetName, 14, 9, 1d),
					new ExpectedCellValue(sheetName, 15, 9, 1d),
					new ExpectedCellValue(sheetName, 1, 10, null),
					new ExpectedCellValue(sheetName, 2, 10, null),
					new ExpectedCellValue(sheetName, 3, 10, "TRUE"),
					new ExpectedCellValue(sheetName, 4, 10, 0),
					new ExpectedCellValue(sheetName, 5, 10, 0.5),
					new ExpectedCellValue(sheetName, 6, 10, 0.5),
					new ExpectedCellValue(sheetName, 7, 10, 0),
					new ExpectedCellValue(sheetName, 8, 10, 0.4),
					new ExpectedCellValue(sheetName, 9, 10, 0.6),
					new ExpectedCellValue(sheetName, 10, 10, 0),
					new ExpectedCellValue(sheetName, 11, 10, 0.4),
					new ExpectedCellValue(sheetName, 12, 10, 0.6),
					new ExpectedCellValue(sheetName, 13, 10, 1d),
					new ExpectedCellValue(sheetName, 14, 10, 1d),
					new ExpectedCellValue(sheetName, 15, 10, 1d),
					new ExpectedCellValue(sheetName, 1, 11, null),
					new ExpectedCellValue(sheetName, 2, 11, "Grand Total"),
					new ExpectedCellValue(sheetName, 3, 11, null),
					new ExpectedCellValue(sheetName, 4, 11, 0.25),
					new ExpectedCellValue(sheetName, 5, 11, 0.333),
					new ExpectedCellValue(sheetName, 6, 11, 0.4167),
					new ExpectedCellValue(sheetName, 7, 11, 0.25),
					new ExpectedCellValue(sheetName, 8, 11, 0.2917),
					new ExpectedCellValue(sheetName, 9, 11, 0.4583),
					new ExpectedCellValue(sheetName, 10, 11, 0.2308),
					new ExpectedCellValue(sheetName, 11, 11, 0.3462),
					new ExpectedCellValue(sheetName, 12, 11, 0.4231),
					new ExpectedCellValue(sheetName, 13, 11, 1d),
					new ExpectedCellValue(sheetName, 14, 11, 1d),
					new ExpectedCellValue(sheetName, 15, 11, 1d)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableMixedDataFieldFunctionTypes.xlsx")]
		public void PivotTableRefreshDataFieldMixedFunctionsAndShowDataAsPercentOfParentRowTotal()
		{
			var file = new FileInfo("PivotTableMixedDataFieldFunctionTypes.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable2"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("A19:N36"), pivotTable.Address);
					Assert.AreEqual(16, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 19, 1, null),
					new ExpectedCellValue(sheetName, 20, 1, null),
					new ExpectedCellValue(sheetName, 21, 1, "Row Labels"),
					new ExpectedCellValue(sheetName, 22, 1, "Begin-Total"),
					new ExpectedCellValue(sheetName, 23, 1, "Average of Indentation"),
					new ExpectedCellValue(sheetName, 24, 1, "Count of Balance at End Date"),
					new ExpectedCellValue(sheetName, 25, 1, "Product of Net Change"),
					new ExpectedCellValue(sheetName, 26, 1, "End-Total"),
					new ExpectedCellValue(sheetName, 27, 1, "Average of Indentation"),
					new ExpectedCellValue(sheetName, 28, 1, "Count of Balance at End Date"),
					new ExpectedCellValue(sheetName, 29, 1, "Product of Net Change"),
					new ExpectedCellValue(sheetName, 30, 1, "Posting"),
					new ExpectedCellValue(sheetName, 31, 1, "Average of Indentation"),
					new ExpectedCellValue(sheetName, 32, 1, "Count of Balance at End Date"),
					new ExpectedCellValue(sheetName, 33, 1, "Product of Net Change"),
					new ExpectedCellValue(sheetName, 34, 1, "Total Average of Indentation"),
					new ExpectedCellValue(sheetName, 35, 1, "Total Count of Balance at End Date"),
					new ExpectedCellValue(sheetName, 36, 1, "Total Product of Net Change"),
					new ExpectedCellValue(sheetName, 19, 2, "Column Labels"),
					new ExpectedCellValue(sheetName, 20, 2, "FALSE"),
					new ExpectedCellValue(sheetName, 21, 2, 2010),
					new ExpectedCellValue(sheetName, 22, 2, null),
					new ExpectedCellValue(sheetName, 23, 2, 0.8),
					new ExpectedCellValue(sheetName, 24, 2, 0.5),
					new ExpectedCellValue(sheetName, 25, 2, null),
					new ExpectedCellValue(sheetName, 26, 2, null),
					new ExpectedCellValue(sheetName, 27, 2, 0),
					new ExpectedCellValue(sheetName, 28, 2, 0),
					new ExpectedCellValue(sheetName, 29, 2, null),
					new ExpectedCellValue(sheetName, 30, 2, null),
					new ExpectedCellValue(sheetName, 31, 2, 1.20),
					new ExpectedCellValue(sheetName, 32, 2, 0.5),
					new ExpectedCellValue(sheetName, 33, 2, null),
					new ExpectedCellValue(sheetName, 34, 2, 1d),
					new ExpectedCellValue(sheetName, 35, 2, 1d),
					new ExpectedCellValue(sheetName, 36, 2, null),
					new ExpectedCellValue(sheetName, 19, 3, null),
					new ExpectedCellValue(sheetName, 20, 3, null),
					new ExpectedCellValue(sheetName, 21, 3, 2012),
					new ExpectedCellValue(sheetName, 22, 3, null),
					new ExpectedCellValue(sheetName, 23, 3, 0),
					new ExpectedCellValue(sheetName, 24, 3, 0),
					new ExpectedCellValue(sheetName, 25, 3, null),
					new ExpectedCellValue(sheetName, 26, 3, null),
					new ExpectedCellValue(sheetName, 27, 3, 0),
					new ExpectedCellValue(sheetName, 28, 3, 0),
					new ExpectedCellValue(sheetName, 29, 3, null),
					new ExpectedCellValue(sheetName, 30, 3, null),
					new ExpectedCellValue(sheetName, 31, 3, 1d),
					new ExpectedCellValue(sheetName, 32, 3, 1d),
					new ExpectedCellValue(sheetName, 33, 3, null),
					new ExpectedCellValue(sheetName, 34, 3, 1d),
					new ExpectedCellValue(sheetName, 35, 3, 1d),
					new ExpectedCellValue(sheetName, 36, 3, null),
					new ExpectedCellValue(sheetName, 19, 4, null),
					new ExpectedCellValue(sheetName, 20, 4, null),
					new ExpectedCellValue(sheetName, 21, 4, 2013),
					new ExpectedCellValue(sheetName, 22, 4, null),
					new ExpectedCellValue(sheetName, 23, 4, 0),
					new ExpectedCellValue(sheetName, 24, 4, 0),
					new ExpectedCellValue(sheetName, 25, 4, 0),
					new ExpectedCellValue(sheetName, 26, 4, null),
					new ExpectedCellValue(sheetName, 27, 4, 0),
					new ExpectedCellValue(sheetName, 28, 4, 0),
					new ExpectedCellValue(sheetName, 29, 4, 0),
					new ExpectedCellValue(sheetName, 30, 4, null),
					new ExpectedCellValue(sheetName, 31, 4, 1d),
					new ExpectedCellValue(sheetName, 32, 4, 1d),
					new ExpectedCellValue(sheetName, 33, 4, 1d),
					new ExpectedCellValue(sheetName, 34, 4, 1d),
					new ExpectedCellValue(sheetName, 35, 4, 1d),
					new ExpectedCellValue(sheetName, 36, 4, 1d),
					new ExpectedCellValue(sheetName, 19, 5, null),
					new ExpectedCellValue(sheetName, 20, 5, null),
					new ExpectedCellValue(sheetName, 21, 5, 2014),
					new ExpectedCellValue(sheetName, 22, 5, null),
					new ExpectedCellValue(sheetName, 23, 5, 1d),
					new ExpectedCellValue(sheetName, 24, 5, 1d),
					new ExpectedCellValue(sheetName, 25, 5, null),
					new ExpectedCellValue(sheetName, 26, 5, null),
					new ExpectedCellValue(sheetName, 27, 5, 0),
					new ExpectedCellValue(sheetName, 28, 5, 0),
					new ExpectedCellValue(sheetName, 29, 5, null),
					new ExpectedCellValue(sheetName, 30, 5, null),
					new ExpectedCellValue(sheetName, 31, 5, 0),
					new ExpectedCellValue(sheetName, 32, 5, 0),
					new ExpectedCellValue(sheetName, 33, 5, null),
					new ExpectedCellValue(sheetName, 34, 5, 1d),
					new ExpectedCellValue(sheetName, 35, 5, 1d),
					new ExpectedCellValue(sheetName, 36, 5, null),
					new ExpectedCellValue(sheetName, 19, 6, null),
					new ExpectedCellValue(sheetName, 20, 6, null),
					new ExpectedCellValue(sheetName, 21, 6, 2016),
					new ExpectedCellValue(sheetName, 22, 6, null),
					new ExpectedCellValue(sheetName, 23, 6, 0),
					new ExpectedCellValue(sheetName, 24, 6, 0.5),
					new ExpectedCellValue(sheetName, 25, 6, null),
					new ExpectedCellValue(sheetName, 26, 6, null),
					new ExpectedCellValue(sheetName, 27, 6, 1.5),
					new ExpectedCellValue(sheetName, 28, 6, 0.5),
					new ExpectedCellValue(sheetName, 29, 6, null),
					new ExpectedCellValue(sheetName, 30, 6, null),
					new ExpectedCellValue(sheetName, 31, 6, 0),
					new ExpectedCellValue(sheetName, 32, 6, 0),
					new ExpectedCellValue(sheetName, 33, 6, null),
					new ExpectedCellValue(sheetName, 34, 6, 1d),
					new ExpectedCellValue(sheetName, 35, 6, 1d),
					new ExpectedCellValue(sheetName, 36, 6, null),
					new ExpectedCellValue(sheetName, 19, 7, null),
					new ExpectedCellValue(sheetName, 20, 7, null),
					new ExpectedCellValue(sheetName, 21, 7, 2017),
					new ExpectedCellValue(sheetName, 22, 7, null),
					new ExpectedCellValue(sheetName, 23, 7, 0),
					new ExpectedCellValue(sheetName, 24, 7, 0),
					new ExpectedCellValue(sheetName, 25, 7, null),
					new ExpectedCellValue(sheetName, 26, 7, null),
					new ExpectedCellValue(sheetName, 27, 7, 0),
					new ExpectedCellValue(sheetName, 28, 7, 0),
					new ExpectedCellValue(sheetName, 29, 7, null),
					new ExpectedCellValue(sheetName, 30, 7, null),
					new ExpectedCellValue(sheetName, 31, 7, 1d),
					new ExpectedCellValue(sheetName, 32, 7, 1d),
					new ExpectedCellValue(sheetName, 33, 7, null),
					new ExpectedCellValue(sheetName, 34, 7, 1d),
					new ExpectedCellValue(sheetName, 35, 7, 1d),
					new ExpectedCellValue(sheetName, 36, 7, null),
					new ExpectedCellValue(sheetName, 19, 8, null),
					new ExpectedCellValue(sheetName, 20, 8, null),
					new ExpectedCellValue(sheetName, 21, 8, 2018),
					new ExpectedCellValue(sheetName, 22, 8, null),
					new ExpectedCellValue(sheetName, 23, 8, 0),
					new ExpectedCellValue(sheetName, 24, 8, 0),
					new ExpectedCellValue(sheetName, 25, 8, 0),
					new ExpectedCellValue(sheetName, 26, 8, null),
					new ExpectedCellValue(sheetName, 27, 8, 0.8),
					new ExpectedCellValue(sheetName, 28, 8, 0.5),
					new ExpectedCellValue(sheetName, 29, 8, 0),
					new ExpectedCellValue(sheetName, 30, 8, null),
					new ExpectedCellValue(sheetName, 31, 8, 1.2),
					new ExpectedCellValue(sheetName, 32, 8, 0.5),
					new ExpectedCellValue(sheetName, 33, 8, 1d),
					new ExpectedCellValue(sheetName, 34, 8, 1d),
					new ExpectedCellValue(sheetName, 35, 8, 1d),
					new ExpectedCellValue(sheetName, 36, 8, 1d),
					new ExpectedCellValue(sheetName, 19, 9, null),
					new ExpectedCellValue(sheetName, 20, 9, "TRUE"),
					new ExpectedCellValue(sheetName, 21, 9, 2012),
					new ExpectedCellValue(sheetName, 22, 9, null),
					new ExpectedCellValue(sheetName, 23, 9, 0.8333),
					new ExpectedCellValue(sheetName, 24, 9, 0.25),
					new ExpectedCellValue(sheetName, 25, 9, null),
					new ExpectedCellValue(sheetName, 26, 9, null),
					new ExpectedCellValue(sheetName, 27, 9, 0.8333),
					new ExpectedCellValue(sheetName, 28, 9, 0.25),
					new ExpectedCellValue(sheetName, 29, 9, null),
					new ExpectedCellValue(sheetName, 30, 9, null),
					new ExpectedCellValue(sheetName, 31, 9, 1.25),
					new ExpectedCellValue(sheetName, 32, 9, 0.5),
					new ExpectedCellValue(sheetName, 33, 9, null),
					new ExpectedCellValue(sheetName, 34, 9, 1d),
					new ExpectedCellValue(sheetName, 35, 9, 1d),
					new ExpectedCellValue(sheetName, 36, 9, null),
					new ExpectedCellValue(sheetName, 19, 10, null),
					new ExpectedCellValue(sheetName, 20, 10, null),
					new ExpectedCellValue(sheetName, 21, 10, 2013),
					new ExpectedCellValue(sheetName, 22, 10, null),
					new ExpectedCellValue(sheetName, 23, 10, 1d),
					new ExpectedCellValue(sheetName, 24, 10, 0.5),
					new ExpectedCellValue(sheetName, 25, 10, null),
					new ExpectedCellValue(sheetName, 26, 10, null),
					new ExpectedCellValue(sheetName, 27, 10, 1d),
					new ExpectedCellValue(sheetName, 28, 10, 0.5),
					new ExpectedCellValue(sheetName, 29, 10, null),
					new ExpectedCellValue(sheetName, 30, 10, null),
					new ExpectedCellValue(sheetName, 31, 10, 0),
					new ExpectedCellValue(sheetName, 32, 10, 0),
					new ExpectedCellValue(sheetName, 33, 10, null),
					new ExpectedCellValue(sheetName, 34, 10, 1d),
					new ExpectedCellValue(sheetName, 35, 10, 1d),
					new ExpectedCellValue(sheetName, 36, 10, null),
					new ExpectedCellValue(sheetName, 19, 11, null),
					new ExpectedCellValue(sheetName, 20, 11, null),
					new ExpectedCellValue(sheetName, 21, 11, 2014),
					new ExpectedCellValue(sheetName, 22, 11, null),
					new ExpectedCellValue(sheetName, 23, 11, 2d),
					new ExpectedCellValue(sheetName, 24, 11, 0.3333),
					new ExpectedCellValue(sheetName, 25, 11, null),
					new ExpectedCellValue(sheetName, 26, 11, null),
					new ExpectedCellValue(sheetName, 27, 11, 0.5),
					new ExpectedCellValue(sheetName, 28, 11, 0.6667),
					new ExpectedCellValue(sheetName, 29, 11, null),
					new ExpectedCellValue(sheetName, 30, 11, null),
					new ExpectedCellValue(sheetName, 31, 11, 0),
					new ExpectedCellValue(sheetName, 32, 11, 0),
					new ExpectedCellValue(sheetName, 33, 11, null),
					new ExpectedCellValue(sheetName, 34, 11, 1d),
					new ExpectedCellValue(sheetName, 35, 11, 1d),
					new ExpectedCellValue(sheetName, 36, 11, null),
					new ExpectedCellValue(sheetName, 19, 12, null),
					new ExpectedCellValue(sheetName, 20, 12, null),
					new ExpectedCellValue(sheetName, 21, 12, 2017),
					new ExpectedCellValue(sheetName, 22, 12, null),
					new ExpectedCellValue(sheetName, 23, 12, 0),
					new ExpectedCellValue(sheetName, 24, 12, 0),
					new ExpectedCellValue(sheetName, 25, 12, 0),
					new ExpectedCellValue(sheetName, 26, 12, null),
					new ExpectedCellValue(sheetName, 27, 12, 0),
					new ExpectedCellValue(sheetName, 28, 12, 0),
					new ExpectedCellValue(sheetName, 29, 12, 0),
					new ExpectedCellValue(sheetName, 30, 12, null),
					new ExpectedCellValue(sheetName, 31, 12, 1d),
					new ExpectedCellValue(sheetName, 32, 12, 1d),
					new ExpectedCellValue(sheetName, 33, 12, 1d),
					new ExpectedCellValue(sheetName, 34, 12, 1d),
					new ExpectedCellValue(sheetName, 35, 12, 1d),
					new ExpectedCellValue(sheetName, 36, 12, 1d),
					new ExpectedCellValue(sheetName, 19, 13, null),
					new ExpectedCellValue(sheetName, 20, 13, null),
					new ExpectedCellValue(sheetName, 21, 13, 2018),
					new ExpectedCellValue(sheetName, 22, 13, null),
					new ExpectedCellValue(sheetName, 23, 13, 0),
					new ExpectedCellValue(sheetName, 24, 13, 0),
					new ExpectedCellValue(sheetName, 25, 13, 0),
					new ExpectedCellValue(sheetName, 26, 13, null),
					new ExpectedCellValue(sheetName, 27, 13, 0.4286),
					new ExpectedCellValue(sheetName, 28, 13, 0.5),
					new ExpectedCellValue(sheetName, 29, 13, 0),
					new ExpectedCellValue(sheetName, 30, 13, null),
					new ExpectedCellValue(sheetName, 31, 13, 1.2857),
					new ExpectedCellValue(sheetName, 32, 13, 0.5),
					new ExpectedCellValue(sheetName, 33, 13, 0),
					new ExpectedCellValue(sheetName, 34, 13, 1d),
					new ExpectedCellValue(sheetName, 35, 13, 1d),
					new ExpectedCellValue(sheetName, 36, 13, 1d),
					new ExpectedCellValue(sheetName, 19, 14, null),
					new ExpectedCellValue(sheetName, 20, 14, "Grand Total"),
					new ExpectedCellValue(sheetName, 21, 14, null),
					new ExpectedCellValue(sheetName, 22, 14, null),
					new ExpectedCellValue(sheetName, 23, 14, 0.75),
					new ExpectedCellValue(sheetName, 24, 14, 0.25),
					new ExpectedCellValue(sheetName, 25, 14, null),
					new ExpectedCellValue(sheetName, 26, 14, null),
					new ExpectedCellValue(sheetName, 27, 14, 0.7),
					new ExpectedCellValue(sheetName, 28, 14, 0.2917),
					new ExpectedCellValue(sheetName, 29, 14, null),
					new ExpectedCellValue(sheetName, 30, 14, null),
					new ExpectedCellValue(sheetName, 31, 14, 1.35),
					new ExpectedCellValue(sheetName, 32, 14, 0.4583),
					new ExpectedCellValue(sheetName, 33, 14, null),
					new ExpectedCellValue(sheetName, 34, 14, 1d),
					new ExpectedCellValue(sheetName, 35, 14, 1d),
					new ExpectedCellValue(sheetName, 36, 14, null)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableMixedDataFieldFunctionTypes.xlsx")]
		public void PivotTableRefreshDataFieldShowDataAsPercentOfParentRowTotalColumnDataFields()
		{
			var file = new FileInfo("PivotTableMixedDataFieldFunctionTypes.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable3"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("A40:L66"), pivotTable.Address);
					Assert.AreEqual(16, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 40, 1, null),
					new ExpectedCellValue(sheetName, 41, 1, null),
					new ExpectedCellValue(sheetName, 42, 1, "Posting Date"),
					new ExpectedCellValue(sheetName, 43, 1, "Jan"),
					new ExpectedCellValue(sheetName, 44, 1, null),
					new ExpectedCellValue(sheetName, 45, 1, null),
					new ExpectedCellValue(sheetName, 46, 1, "Feb"),
					new ExpectedCellValue(sheetName, 47, 1, null),
					new ExpectedCellValue(sheetName, 48, 1, "Mar"),
					new ExpectedCellValue(sheetName, 49, 1, "Apr"),
					new ExpectedCellValue(sheetName, 50, 1, "May"),
					new ExpectedCellValue(sheetName, 51, 1, null),
					new ExpectedCellValue(sheetName, 52, 1, "Jun"),
					new ExpectedCellValue(sheetName, 53, 1, "Jul"),
					new ExpectedCellValue(sheetName, 54, 1, null),
					new ExpectedCellValue(sheetName, 55, 1, "Aug"),
					new ExpectedCellValue(sheetName, 56, 1, null),
					new ExpectedCellValue(sheetName, 57, 1, "Sep"),
					new ExpectedCellValue(sheetName, 58, 1, "Oct"),
					new ExpectedCellValue(sheetName, 59, 1, null),
					new ExpectedCellValue(sheetName, 60, 1, null),
					new ExpectedCellValue(sheetName, 61, 1, null),
					new ExpectedCellValue(sheetName, 62, 1, null),
					new ExpectedCellValue(sheetName, 63, 1, "Nov"),
					new ExpectedCellValue(sheetName, 64, 1, "Dec"),
					new ExpectedCellValue(sheetName, 65, 1, null),
					new ExpectedCellValue(sheetName, 66, 1, "Grand Total"),
					new ExpectedCellValue(sheetName, 40, 2, null),
					new ExpectedCellValue(sheetName, 41, 2, null),
					new ExpectedCellValue(sheetName, 42, 2, "Account Type"),
					new ExpectedCellValue(sheetName, 43, 2, "Begin-Total"),
					new ExpectedCellValue(sheetName, 44, 2, "End-Total"),
					new ExpectedCellValue(sheetName, 45, 2, "Posting"),
					new ExpectedCellValue(sheetName, 46, 2, "End-Total"),
					new ExpectedCellValue(sheetName, 47, 2, "Posting"),
					new ExpectedCellValue(sheetName, 48, 2, "Begin-Total"),
					new ExpectedCellValue(sheetName, 49, 2, "End-Total"),
					new ExpectedCellValue(sheetName, 50, 2, "Begin-Total"),
					new ExpectedCellValue(sheetName, 51, 2, "End-Total"),
					new ExpectedCellValue(sheetName, 52, 2, "Posting"),
					new ExpectedCellValue(sheetName, 53, 2, "Begin-Total"),
					new ExpectedCellValue(sheetName, 54, 2, "End-Total"),
					new ExpectedCellValue(sheetName, 55, 2, "End-Total"),
					new ExpectedCellValue(sheetName, 56, 2, "Posting"),
					new ExpectedCellValue(sheetName, 57, 2, "Begin-Total"),
					new ExpectedCellValue(sheetName, 58, 2, "Begin-Total"),
					new ExpectedCellValue(sheetName, 59, 2, "End-Total"),
					new ExpectedCellValue(sheetName, 60, 2, null),
					new ExpectedCellValue(sheetName, 61, 2, "Posting"),
					new ExpectedCellValue(sheetName, 62, 2, null),
					new ExpectedCellValue(sheetName, 63, 2, "Posting"),
					new ExpectedCellValue(sheetName, 64, 2, "Posting"),
					new ExpectedCellValue(sheetName, 65, 2, null),
					new ExpectedCellValue(sheetName, 66, 2, null),
					new ExpectedCellValue(sheetName, 40, 3, null),
					new ExpectedCellValue(sheetName, 41, 3, null),
					new ExpectedCellValue(sheetName, 42, 3, "Debit/Credit"),
					new ExpectedCellValue(sheetName, 43, 3, "Debit"),
					new ExpectedCellValue(sheetName, 44, 3, "Both"),
					new ExpectedCellValue(sheetName, 45, 3, "Debit"),
					new ExpectedCellValue(sheetName, 46, 3, "Both"),
					new ExpectedCellValue(sheetName, 47, 3, "Both"),
					new ExpectedCellValue(sheetName, 48, 3, "Both"),
					new ExpectedCellValue(sheetName, 49, 3, "Both"),
					new ExpectedCellValue(sheetName, 50, 3, "Debit"),
					new ExpectedCellValue(sheetName, 51, 3, "Debit"),
					new ExpectedCellValue(sheetName, 52, 3, "Debit"),
					new ExpectedCellValue(sheetName, 53, 3, "Both"),
					new ExpectedCellValue(sheetName, 54, 3, "Both"),
					new ExpectedCellValue(sheetName, 55, 3, "Both"),
					new ExpectedCellValue(sheetName, 56, 3, "Both"),
					new ExpectedCellValue(sheetName, 57, 3, "Credit"),
					new ExpectedCellValue(sheetName, 58, 3, "Credit"),
					new ExpectedCellValue(sheetName, 59, 3, "Both"),
					new ExpectedCellValue(sheetName, 60, 3, "Credit"),
					new ExpectedCellValue(sheetName, 61, 3, "Both"),
					new ExpectedCellValue(sheetName, 62, 3, "Debit"),
					new ExpectedCellValue(sheetName, 63, 3, "Both"),
					new ExpectedCellValue(sheetName, 64, 3, "Credit"),
					new ExpectedCellValue(sheetName, 65, 3, "Debit"),
					new ExpectedCellValue(sheetName, 66, 3, null),
					new ExpectedCellValue(sheetName, 40, 4, "Blocked"),
					new ExpectedCellValue(sheetName, 41, 4, "FALSE"),
					new ExpectedCellValue(sheetName, 42, 4, "Count of Indentation"),
					new ExpectedCellValue(sheetName, 43, 4, null),
					new ExpectedCellValue(sheetName, 44, 4, null),
					new ExpectedCellValue(sheetName, 45, 4, null),
					new ExpectedCellValue(sheetName, 46, 4, 1d),
					new ExpectedCellValue(sheetName, 47, 4, null),
					new ExpectedCellValue(sheetName, 48, 4, null),
					new ExpectedCellValue(sheetName, 49, 4, 1d),
					new ExpectedCellValue(sheetName, 50, 4, 1d),
					new ExpectedCellValue(sheetName, 51, 4, null),
					new ExpectedCellValue(sheetName, 52, 4, 1d),
					new ExpectedCellValue(sheetName, 53, 4, 1d),
					new ExpectedCellValue(sheetName, 54, 4, null),
					new ExpectedCellValue(sheetName, 55, 4, 1d),
					new ExpectedCellValue(sheetName, 56, 4, null),
					new ExpectedCellValue(sheetName, 57, 4, null),
					new ExpectedCellValue(sheetName, 58, 4, 1d),
					new ExpectedCellValue(sheetName, 59, 4, null),
					new ExpectedCellValue(sheetName, 60, 4, null),
					new ExpectedCellValue(sheetName, 61, 4, 0.5),
					new ExpectedCellValue(sheetName, 62, 4, 0.5),
					new ExpectedCellValue(sheetName, 63, 4, 1d),
					new ExpectedCellValue(sheetName, 64, 4, 0.5),
					new ExpectedCellValue(sheetName, 65, 4, 0.5),
					new ExpectedCellValue(sheetName, 66, 4, 1d),
					new ExpectedCellValue(sheetName, 40, 5, "Values"),
					new ExpectedCellValue(sheetName, 41, 5, null),
					new ExpectedCellValue(sheetName, 42, 5, "Average of Income/Balance"),
					new ExpectedCellValue(sheetName, 43, 5, null),
					new ExpectedCellValue(sheetName, 44, 5, null),
					new ExpectedCellValue(sheetName, 45, 5, null),
					new ExpectedCellValue(sheetName, 46, 5, null),
					new ExpectedCellValue(sheetName, 47, 5, null),
					new ExpectedCellValue(sheetName, 48, 5, null),
					new ExpectedCellValue(sheetName, 49, 5, null),
					new ExpectedCellValue(sheetName, 50, 5, null),
					new ExpectedCellValue(sheetName, 51, 5, null),
					new ExpectedCellValue(sheetName, 52, 5, null),
					new ExpectedCellValue(sheetName, 53, 5, null),
					new ExpectedCellValue(sheetName, 54, 5, null),
					new ExpectedCellValue(sheetName, 55, 5, null),
					new ExpectedCellValue(sheetName, 56, 5, null),
					new ExpectedCellValue(sheetName, 57, 5, null),
					new ExpectedCellValue(sheetName, 58, 5, null),
					new ExpectedCellValue(sheetName, 59, 5, null),
					new ExpectedCellValue(sheetName, 60, 5, null),
					new ExpectedCellValue(sheetName, 61, 5, null),
					new ExpectedCellValue(sheetName, 62, 5, null),
					new ExpectedCellValue(sheetName, 63, 5, null),
					new ExpectedCellValue(sheetName, 64, 5, null),
					new ExpectedCellValue(sheetName, 65, 5, null),
					new ExpectedCellValue(sheetName, 66, 5, null),
					new ExpectedCellValue(sheetName, 40, 6, null),
					new ExpectedCellValue(sheetName, 41, 6, null),
					new ExpectedCellValue(sheetName, 42, 6, "Product of Net Change"),
					new ExpectedCellValue(sheetName, 43, 6, null),
					new ExpectedCellValue(sheetName, 44, 6, null),
					new ExpectedCellValue(sheetName, 45, 6, null),
					new ExpectedCellValue(sheetName, 46, 6, null),
					new ExpectedCellValue(sheetName, 47, 6, null),
					new ExpectedCellValue(sheetName, 48, 6, null),
					new ExpectedCellValue(sheetName, 49, 6, 1d),
					new ExpectedCellValue(sheetName, 50, 6, null),
					new ExpectedCellValue(sheetName, 51, 6, null),
					new ExpectedCellValue(sheetName, 52, 6, null),
					new ExpectedCellValue(sheetName, 53, 6, null),
					new ExpectedCellValue(sheetName, 54, 6, null),
					new ExpectedCellValue(sheetName, 55, 6, null),
					new ExpectedCellValue(sheetName, 56, 6, null),
					new ExpectedCellValue(sheetName, 57, 6, null),
					new ExpectedCellValue(sheetName, 58, 6, null),
					new ExpectedCellValue(sheetName, 59, 6, null),
					new ExpectedCellValue(sheetName, 60, 6, null),
					new ExpectedCellValue(sheetName, 61, 6, 0),
					new ExpectedCellValue(sheetName, 62, 6, 0),
					new ExpectedCellValue(sheetName, 63, 6, 1d),
					new ExpectedCellValue(sheetName, 64, 6, null),
					new ExpectedCellValue(sheetName, 65, 6, null),
					new ExpectedCellValue(sheetName, 66, 6, null),
					new ExpectedCellValue(sheetName, 40, 7, null),
					new ExpectedCellValue(sheetName, 41, 7, "TRUE"),
					new ExpectedCellValue(sheetName, 42, 7, "Count of Indentation"),
					new ExpectedCellValue(sheetName, 44, 7, 1d),
					new ExpectedCellValue(sheetName, 43, 7, 1d),
					new ExpectedCellValue(sheetName, 45, 7, 1d),
					new ExpectedCellValue(sheetName, 46, 7, null),
					new ExpectedCellValue(sheetName, 47, 7, 1d),
					new ExpectedCellValue(sheetName, 48, 7, 1d),
					new ExpectedCellValue(sheetName, 49, 7, null),
					new ExpectedCellValue(sheetName, 50, 7, null),
					new ExpectedCellValue(sheetName, 51, 7, 1d),
					new ExpectedCellValue(sheetName, 52, 7, null),
					new ExpectedCellValue(sheetName, 53, 7, null),
					new ExpectedCellValue(sheetName, 54, 7, 1d),
					new ExpectedCellValue(sheetName, 55, 7, null),
					new ExpectedCellValue(sheetName, 56, 7, 1d),
					new ExpectedCellValue(sheetName, 57, 7, 1d),
					new ExpectedCellValue(sheetName, 58, 7, null),
					new ExpectedCellValue(sheetName, 59, 7, 0.5),
					new ExpectedCellValue(sheetName, 60, 7, 0.5),
					new ExpectedCellValue(sheetName, 61, 7, null),
					new ExpectedCellValue(sheetName, 62, 7, null),
					new ExpectedCellValue(sheetName, 63, 7, 1d),
					new ExpectedCellValue(sheetName, 64, 7, 0),
					new ExpectedCellValue(sheetName, 65, 7, 1d),
					new ExpectedCellValue(sheetName, 66, 7, 1d),
					new ExpectedCellValue(sheetName, 40, 8, null),
					new ExpectedCellValue(sheetName, 41, 8, null),
					new ExpectedCellValue(sheetName, 42, 8, "Average of Income/Balance"),
					new ExpectedCellValue(sheetName, 44, 8, null),
					new ExpectedCellValue(sheetName, 43, 8, null),
					new ExpectedCellValue(sheetName, 45, 8, null),
					new ExpectedCellValue(sheetName, 46, 8, null),
					new ExpectedCellValue(sheetName, 47, 8, null),
					new ExpectedCellValue(sheetName, 48, 8, null),
					new ExpectedCellValue(sheetName, 49, 8, null),
					new ExpectedCellValue(sheetName, 50, 8, null),
					new ExpectedCellValue(sheetName, 51, 8, null),
					new ExpectedCellValue(sheetName, 52, 8, null),
					new ExpectedCellValue(sheetName, 53, 8, null),
					new ExpectedCellValue(sheetName, 54, 8, null),
					new ExpectedCellValue(sheetName, 55, 8, null),
					new ExpectedCellValue(sheetName, 56, 8, null),
					new ExpectedCellValue(sheetName, 57, 8, null),
					new ExpectedCellValue(sheetName, 58, 8, null),
					new ExpectedCellValue(sheetName, 59, 8, null),
					new ExpectedCellValue(sheetName, 60, 8, null),
					new ExpectedCellValue(sheetName, 61, 8, null),
					new ExpectedCellValue(sheetName, 62, 8, null),
					new ExpectedCellValue(sheetName, 63, 8, null),
					new ExpectedCellValue(sheetName, 64, 8, null),
					new ExpectedCellValue(sheetName, 65, 8, null),
					new ExpectedCellValue(sheetName, 66, 8, null),
					new ExpectedCellValue(sheetName, 40, 9, null),
					new ExpectedCellValue(sheetName, 41, 9, null),
					new ExpectedCellValue(sheetName, 42, 9, "Product of Net Change"),
					new ExpectedCellValue(sheetName, 44, 9, null),
					new ExpectedCellValue(sheetName, 43, 9, null),
					new ExpectedCellValue(sheetName, 45, 9, 1d),
					new ExpectedCellValue(sheetName, 46, 9, null),
					new ExpectedCellValue(sheetName, 47, 9, null),
					new ExpectedCellValue(sheetName, 48, 9, null),
					new ExpectedCellValue(sheetName, 49, 9, null),
					new ExpectedCellValue(sheetName, 50, 9, null),
					new ExpectedCellValue(sheetName, 51, 9, 1d),
					new ExpectedCellValue(sheetName, 52, 9, null),
					new ExpectedCellValue(sheetName, 53, 9, null),
					new ExpectedCellValue(sheetName, 54, 9, 1d),
					new ExpectedCellValue(sheetName, 55, 9, null),
					new ExpectedCellValue(sheetName, 56, 9, null),
					new ExpectedCellValue(sheetName, 57, 9, null),
					new ExpectedCellValue(sheetName, 58, 9, null),
					new ExpectedCellValue(sheetName, 59, 9, 0),
					new ExpectedCellValue(sheetName, 60, 9, 0),
					new ExpectedCellValue(sheetName, 61, 9, null),
					new ExpectedCellValue(sheetName, 62, 9, null),
					new ExpectedCellValue(sheetName, 63, 9, null),
					new ExpectedCellValue(sheetName, 64, 9, 0),
					new ExpectedCellValue(sheetName, 65, 9, 1d),
					new ExpectedCellValue(sheetName, 66, 9, null),
					new ExpectedCellValue(sheetName, 40, 10, null),
					new ExpectedCellValue(sheetName, 41, 10, "Total Count of Indentation"),
					new ExpectedCellValue(sheetName, 42, 10, null),
					new ExpectedCellValue(sheetName, 44, 10, 1d),
					new ExpectedCellValue(sheetName, 43, 10, 1d),
					new ExpectedCellValue(sheetName, 45, 10, 1d),
					new ExpectedCellValue(sheetName, 46, 10, 1d),
					new ExpectedCellValue(sheetName, 47, 10, 1d),
					new ExpectedCellValue(sheetName, 48, 10, 1d),
					new ExpectedCellValue(sheetName, 49, 10, 1d),
					new ExpectedCellValue(sheetName, 50, 10, 1d),
					new ExpectedCellValue(sheetName, 51, 10, 1d),
					new ExpectedCellValue(sheetName, 52, 10, 1d),
					new ExpectedCellValue(sheetName, 53, 10, 1d),
					new ExpectedCellValue(sheetName, 54, 10, 1d),
					new ExpectedCellValue(sheetName, 55, 10, 1d),
					new ExpectedCellValue(sheetName, 56, 10, 1d),
					new ExpectedCellValue(sheetName, 57, 10, 1d),
					new ExpectedCellValue(sheetName, 58, 10, 1d),
					new ExpectedCellValue(sheetName, 59, 10, 0.5),
					new ExpectedCellValue(sheetName, 60, 10, 0.5),
					new ExpectedCellValue(sheetName, 61, 10, 0.5),
					new ExpectedCellValue(sheetName, 62, 10, 0.5),
					new ExpectedCellValue(sheetName, 63, 10, 1d),
					new ExpectedCellValue(sheetName, 64, 10, 0.25),
					new ExpectedCellValue(sheetName, 65, 10, 0.75),
					new ExpectedCellValue(sheetName, 66, 10, 1d),
					new ExpectedCellValue(sheetName, 40, 11, null),
					new ExpectedCellValue(sheetName, 41, 11, "Total Average of Income/Balance"),
					new ExpectedCellValue(sheetName, 42, 11, null),
					new ExpectedCellValue(sheetName, 44, 11, null),
					new ExpectedCellValue(sheetName, 43, 11, null),
					new ExpectedCellValue(sheetName, 45, 11, null),
					new ExpectedCellValue(sheetName, 46, 11, null),
					new ExpectedCellValue(sheetName, 47, 11, null),
					new ExpectedCellValue(sheetName, 48, 11, null),
					new ExpectedCellValue(sheetName, 49, 11, null),
					new ExpectedCellValue(sheetName, 50, 11, null),
					new ExpectedCellValue(sheetName, 51, 11, null),
					new ExpectedCellValue(sheetName, 52, 11, null),
					new ExpectedCellValue(sheetName, 53, 11, null),
					new ExpectedCellValue(sheetName, 54, 11, null),
					new ExpectedCellValue(sheetName, 55, 11, null),
					new ExpectedCellValue(sheetName, 56, 11, null),
					new ExpectedCellValue(sheetName, 57, 11, null),
					new ExpectedCellValue(sheetName, 58, 11, null),
					new ExpectedCellValue(sheetName, 59, 11, null),
					new ExpectedCellValue(sheetName, 60, 11, null),
					new ExpectedCellValue(sheetName, 61, 11, null),
					new ExpectedCellValue(sheetName, 62, 11, null),
					new ExpectedCellValue(sheetName, 63, 11, null),
					new ExpectedCellValue(sheetName, 64, 11, null),
					new ExpectedCellValue(sheetName, 65, 11, null),
					new ExpectedCellValue(sheetName, 66, 11, null),
					new ExpectedCellValue(sheetName, 40, 12, null),
					new ExpectedCellValue(sheetName, 41, 12, "Total Product of Net Change"),
					new ExpectedCellValue(sheetName, 42, 12, null),
					new ExpectedCellValue(sheetName, 44, 12, null),
					new ExpectedCellValue(sheetName, 43, 12, null),
					new ExpectedCellValue(sheetName, 45, 12, 1d),
					new ExpectedCellValue(sheetName, 46, 12, null),
					new ExpectedCellValue(sheetName, 47, 12, null),
					new ExpectedCellValue(sheetName, 48, 12, null),
					new ExpectedCellValue(sheetName, 49, 12, 1d),
					new ExpectedCellValue(sheetName, 50, 12, null),
					new ExpectedCellValue(sheetName, 51, 12, 1d),
					new ExpectedCellValue(sheetName, 52, 12, null),
					new ExpectedCellValue(sheetName, 53, 12, null),
					new ExpectedCellValue(sheetName, 54, 12, 1d),
					new ExpectedCellValue(sheetName, 55, 12, null),
					new ExpectedCellValue(sheetName, 56, 12, null),
					new ExpectedCellValue(sheetName, 57, 12, null),
					new ExpectedCellValue(sheetName, 58, 12, null),
					new ExpectedCellValue(sheetName, 59, 12, 0),
					new ExpectedCellValue(sheetName, 60, 12, 0),
					new ExpectedCellValue(sheetName, 61, 12, 0),
					new ExpectedCellValue(sheetName, 62, 12, 0),
					new ExpectedCellValue(sheetName, 63, 12, null),
					new ExpectedCellValue(sheetName, 64, 12, null),
					new ExpectedCellValue(sheetName, 65, 12, null),
					new ExpectedCellValue(sheetName, 66, 12, null),
				});
			}
		}
		#endregion

		#region PercentOfParentColumn Tests
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableShowDataAs.xlsx")]
		public void PivotTableRefreshShowDataAsRowFieldsColumnDataFieldsPercentOfParentColumn()
		{
			var file = new FileInfo("PivotTableShowDataAs.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable1"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					wholesalePriceDataField.ShowDataAs = ShowDataAs.PercentOfParentCol;
					unitsSoldDataField.ShowDataAs = ShowDataAs.NoCalculation;
					foreach (var field in pivotTable.Fields)
					{
						field.SubtotalLocation = SubtotalLocation.Off;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B2:D13"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 2, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 2, 3, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 2, 4, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 3, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 3, 3, null),
					new ExpectedCellValue(sheetName, 3, 4, null),
					new ExpectedCellValue(sheetName, 4, 2, "January"),
					new ExpectedCellValue(sheetName, 4, 3, null),
					new ExpectedCellValue(sheetName, 4, 4, 2),
					new ExpectedCellValue(sheetName, 5, 2, "March"),
					new ExpectedCellValue(sheetName, 5, 3, null),
					new ExpectedCellValue(sheetName, 5, 4, 1),
					new ExpectedCellValue(sheetName, 6, 2, "Nashville"),
					new ExpectedCellValue(sheetName, 6, 3, null),
					new ExpectedCellValue(sheetName, 6, 4, null),
					new ExpectedCellValue(sheetName, 7, 2, "January"),
					new ExpectedCellValue(sheetName, 7, 3, null),
					new ExpectedCellValue(sheetName, 7, 4, 2),
					new ExpectedCellValue(sheetName, 8, 2, "February"),
					new ExpectedCellValue(sheetName, 8, 3, null),
					new ExpectedCellValue(sheetName, 8, 4, 6),
					new ExpectedCellValue(sheetName, 9, 2, "March"),
					new ExpectedCellValue(sheetName, 9, 3, null),
					new ExpectedCellValue(sheetName, 9, 4, 2),
					new ExpectedCellValue(sheetName, 10, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 10, 3, null),
					new ExpectedCellValue(sheetName, 10, 4, null),
					new ExpectedCellValue(sheetName, 11, 2, "January"),
					new ExpectedCellValue(sheetName, 11, 3, null),
					new ExpectedCellValue(sheetName, 11, 4, 1),
					new ExpectedCellValue(sheetName, 12, 2, "February"),
					new ExpectedCellValue(sheetName, 12, 3, null),
					new ExpectedCellValue(sheetName, 12, 4, 1),
					new ExpectedCellValue(sheetName, 13, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 13, 3, null),
					new ExpectedCellValue(sheetName, 13, 4, 15)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableShowDataAs.xlsx")]
		public void PivotTableRefreshShowDataAsPercentOfParentColumnWithColumnFieldsAndRowDataFieldsSubtotalsTopAndBottom()
		{
			var file = new FileInfo("PivotTableShowDataAs.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				void validateWorksheet() => TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 19, 2, null),
					new ExpectedCellValue(sheetName, 19, 3, "January"),
					new ExpectedCellValue(sheetName, 19, 4, "January Total"),
					new ExpectedCellValue(sheetName, 19, 5, "February"),
					new ExpectedCellValue(sheetName, 19, 6, null),
					new ExpectedCellValue(sheetName, 19, 7, "February Total"),
					new ExpectedCellValue(sheetName, 19, 8, "March"),
					new ExpectedCellValue(sheetName, 19, 9, null),
					new ExpectedCellValue(sheetName, 19, 10, "March Total"),
					new ExpectedCellValue(sheetName, 19, 11, "Grand Total"),

					new ExpectedCellValue(sheetName, 20, 2, "Values"),
					new ExpectedCellValue(sheetName, 20, 3, "Car Rack"),
					new ExpectedCellValue(sheetName, 20, 4, null),
					new ExpectedCellValue(sheetName, 20, 5, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 20, 6, "Tent"),
					new ExpectedCellValue(sheetName, 20, 7, null),
					new ExpectedCellValue(sheetName, 20, 8, "Car Rack"),
					new ExpectedCellValue(sheetName, 20, 9, "Headlamp"),
					new ExpectedCellValue(sheetName, 20, 10, null),
					new ExpectedCellValue(sheetName, 20, 11, null),

					new ExpectedCellValue(sheetName, 21, 2, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 21, 3, 1),
					new ExpectedCellValue(sheetName, 21, 4, .6280),
					new ExpectedCellValue(sheetName, 21, 5, .3322),
					new ExpectedCellValue(sheetName, 21, 6, .6678),
					new ExpectedCellValue(sheetName, 21, 7, .1501),
					new ExpectedCellValue(sheetName, 21, 8, .9433),
					new ExpectedCellValue(sheetName, 21, 9, .0567),
					new ExpectedCellValue(sheetName, 21, 10, .2219),
					new ExpectedCellValue(sheetName, 21, 11, 1),

					new ExpectedCellValue(sheetName, 22, 2, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 22, 3, 5),
					new ExpectedCellValue(sheetName, 22, 4, 5),
					new ExpectedCellValue(sheetName, 22, 5, 1),
					new ExpectedCellValue(sheetName, 22, 6, 6),
					new ExpectedCellValue(sheetName, 22, 7, 7),
					new ExpectedCellValue(sheetName, 22, 8, 2),
					new ExpectedCellValue(sheetName, 22, 9, 1),
					new ExpectedCellValue(sheetName, 22, 10, 3),
					new ExpectedCellValue(sheetName, 22, 11, 15)
				});

				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable2"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					// Show 'Wholesale Price' data as the percentage of its parent column.
					wholesalePriceDataField.ShowDataAs = ShowDataAs.PercentOfParentCol;
					unitsSoldDataField.ShowDataAs = ShowDataAs.NoCalculation;
					foreach (var field in pivotTable.Fields)
					{
						field.SubtotalLocation = SubtotalLocation.Top;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B18:K22"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				validateWorksheet();

				// Verify that subtotal top provides the same result.
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable2"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					// Show 'Wholesale Price' data as the percentage of its parent column.
					wholesalePriceDataField.ShowDataAs = ShowDataAs.PercentOfParentCol;
					unitsSoldDataField.ShowDataAs = ShowDataAs.NoCalculation;
					foreach (var field in pivotTable.Fields)
					{
						field.SubtotalLocation = SubtotalLocation.Bottom;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B18:K22"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				validateWorksheet();
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableShowDataAs.xlsx")]
		public void PivotTableRefreshShowDataAsPercentOfParentColumnWithColumnFieldsAndRowDataFieldsSubtotalsOff()
		{
			var file = new FileInfo("PivotTableShowDataAs.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable2"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					// Show 'Wholesale Price' data as the percentage of its parent column.
					wholesalePriceDataField.ShowDataAs = ShowDataAs.PercentOfParentCol;
					unitsSoldDataField.ShowDataAs = ShowDataAs.NoCalculation;
					foreach (var field in pivotTable.Fields)
					{
						field.SubtotalLocation = SubtotalLocation.Off;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B18:H22"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}

				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 19, 2, null),
					new ExpectedCellValue(sheetName, 19, 3, "January"),
					new ExpectedCellValue(sheetName, 19, 4, "February"),
					new ExpectedCellValue(sheetName, 19, 5, null),
					new ExpectedCellValue(sheetName, 19, 6, "March"),
					new ExpectedCellValue(sheetName, 19, 7, null),
					new ExpectedCellValue(sheetName, 19, 8, "Grand Total"),

					new ExpectedCellValue(sheetName, 20, 2, "Values"),
					new ExpectedCellValue(sheetName, 20, 3, "Car Rack"),
					new ExpectedCellValue(sheetName, 20, 4, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 20, 5, "Tent"),
					new ExpectedCellValue(sheetName, 20, 6, "Car Rack"),
					new ExpectedCellValue(sheetName, 20, 7, "Headlamp"),
					new ExpectedCellValue(sheetName, 20, 8, null),

					new ExpectedCellValue(sheetName, 21, 2, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 21, 3, 1),
					new ExpectedCellValue(sheetName, 21, 4, .3322),
					new ExpectedCellValue(sheetName, 21, 5, .6678),
					new ExpectedCellValue(sheetName, 21, 6, .9433),
					new ExpectedCellValue(sheetName, 21, 7, .0567),
					new ExpectedCellValue(sheetName, 21, 8, 1),

					new ExpectedCellValue(sheetName, 22, 2, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 22, 3, 5),
					new ExpectedCellValue(sheetName, 22, 4, 1),
					new ExpectedCellValue(sheetName, 22, 5, 6),
					new ExpectedCellValue(sheetName, 22, 6, 2),
					new ExpectedCellValue(sheetName, 22, 7, 1),
					new ExpectedCellValue(sheetName, 22, 8, 15)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableShowDataAs.xlsx")]
		public void PivotTableRefreshShowDataAsColumnFieldsRowFieldsAndColumnDataFieldsPercentOfParentColumnDataFieldSubtotalTopOnOff()
		{
			var file = new FileInfo("PivotTableShowDataAs.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				void validateWorksheet() => TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 33, 2, "January"),

					new ExpectedCellValue(sheetName, 34, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 34, 3, .3333),
					new ExpectedCellValue(sheetName, 34, 4, .3333),
					new ExpectedCellValue(sheetName, 34, 5, .3333),
					new ExpectedCellValue(sheetName, 34, 6, 2),
					new ExpectedCellValue(sheetName, 34, 7, 2),
					new ExpectedCellValue(sheetName, 34, 8, 1),
					new ExpectedCellValue(sheetName, 34, 9, 1),
					new ExpectedCellValue(sheetName, 34, 10, 5),

					new ExpectedCellValue(sheetName, 35, 2, "February"),

					new ExpectedCellValue(sheetName, 36, 2, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 36, 3, 0),
					new ExpectedCellValue(sheetName, 36, 4, 0),
					new ExpectedCellValue(sheetName, 36, 5, 1),
					new ExpectedCellValue(sheetName, 36, 6, null),
					new ExpectedCellValue(sheetName, 36, 7, null),
					new ExpectedCellValue(sheetName, 36, 8, 1),
					new ExpectedCellValue(sheetName, 36, 9, 1),
					new ExpectedCellValue(sheetName, 36, 10, 1),

					new ExpectedCellValue(sheetName, 37, 2, "Tent"),
					new ExpectedCellValue(sheetName, 37, 3, 0),
					new ExpectedCellValue(sheetName, 37, 4, 1),
					new ExpectedCellValue(sheetName, 37, 5, 0),
					new ExpectedCellValue(sheetName, 37, 6, null),
					new ExpectedCellValue(sheetName, 37, 7, 6),
					new ExpectedCellValue(sheetName, 37, 8, null),
					new ExpectedCellValue(sheetName, 37, 9, 1),
					new ExpectedCellValue(sheetName, 37, 10, 6),

					new ExpectedCellValue(sheetName, 38, 2, "March"),

					new ExpectedCellValue(sheetName, 39, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 39, 3, 0),
					new ExpectedCellValue(sheetName, 39, 4, 1),
					new ExpectedCellValue(sheetName, 39, 5, 0),
					new ExpectedCellValue(sheetName, 39, 6, null),
					new ExpectedCellValue(sheetName, 39, 7, 2),
					new ExpectedCellValue(sheetName, 39, 8, null),
					new ExpectedCellValue(sheetName, 39, 9, 1),
					new ExpectedCellValue(sheetName, 39, 10, 2),

					new ExpectedCellValue(sheetName, 40, 2, "Headlamp"),
					new ExpectedCellValue(sheetName, 40, 3, 1),
					new ExpectedCellValue(sheetName, 40, 4, 0),
					new ExpectedCellValue(sheetName, 40, 5, 0),
					new ExpectedCellValue(sheetName, 40, 6, 1),
					new ExpectedCellValue(sheetName, 40, 7, null),
					new ExpectedCellValue(sheetName, 40, 8, null),
					new ExpectedCellValue(sheetName, 40, 9, 1),
					new ExpectedCellValue(sheetName, 40, 10, 1),

					new ExpectedCellValue(sheetName, 41, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 41, 3, .2219),
					new ExpectedCellValue(sheetName, 41, 4, .5189),
					new ExpectedCellValue(sheetName, 41, 5, .2592),
					new ExpectedCellValue(sheetName, 41, 6, 3),
					new ExpectedCellValue(sheetName, 41, 7, 10),
					new ExpectedCellValue(sheetName, 41, 8, 2),
					new ExpectedCellValue(sheetName, 41, 9, 1),
					new ExpectedCellValue(sheetName, 41, 10, 15),
				});

				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable3"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					// Show 'Wholesale Price' data as the percentage of its parent column.
					wholesalePriceDataField.ShowDataAs = ShowDataAs.PercentOfParentCol;
					unitsSoldDataField.ShowDataAs = ShowDataAs.NoCalculation;
					foreach (var field in pivotTable.Fields)
					{
						field.SubtotalLocation = SubtotalLocation.Top;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B30:J41"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				validateWorksheet();
				// Validate subtotal top is on.
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					// January
					new ExpectedCellValue(sheetName, 33, 3, .3333),
					new ExpectedCellValue(sheetName, 33, 4, .3333),
					new ExpectedCellValue(sheetName, 33, 5, .3333),
					new ExpectedCellValue(sheetName, 33, 6, 2),
					new ExpectedCellValue(sheetName, 33, 7, 2),
					new ExpectedCellValue(sheetName, 33, 8, 1),
					new ExpectedCellValue(sheetName, 33, 9, 1),
					new ExpectedCellValue(sheetName, 33, 10, 5),

					// February
					new ExpectedCellValue(sheetName, 35, 3, 0),
					new ExpectedCellValue(sheetName, 35, 4, .6667),
					new ExpectedCellValue(sheetName, 35, 5, .3322),
					new ExpectedCellValue(sheetName, 35, 6, null),
					new ExpectedCellValue(sheetName, 35, 7, 6),
					new ExpectedCellValue(sheetName, 35, 8, 1),
					new ExpectedCellValue(sheetName, 35, 9, 1),
					new ExpectedCellValue(sheetName, 35, 10, 7),

					// March
					new ExpectedCellValue(sheetName, 38, 3, .0567),
					new ExpectedCellValue(sheetName, 38, 4, .9433),
					new ExpectedCellValue(sheetName, 38, 5, 0),
					new ExpectedCellValue(sheetName, 38, 6, 1),
					new ExpectedCellValue(sheetName, 38, 7, 2),
					new ExpectedCellValue(sheetName, 38, 8, null),
					new ExpectedCellValue(sheetName, 38, 9, 1),
					new ExpectedCellValue(sheetName, 38, 10, 3),
				});

				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable3"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					// Show 'Wholesale Price' data as the percentage of its parent column.
					wholesalePriceDataField.ShowDataAs = ShowDataAs.PercentOfParentCol;
					unitsSoldDataField.ShowDataAs = ShowDataAs.NoCalculation;
					foreach (var field in pivotTable.Fields)
					{
						field.SubtotalLocation = SubtotalLocation.Off;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B30:J41"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				validateWorksheet();
				// Validate subtotal top is off.
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					// January
					new ExpectedCellValue(sheetName, 33, 3, null),
					new ExpectedCellValue(sheetName, 33, 4, null),
					new ExpectedCellValue(sheetName, 33, 5, null),
					new ExpectedCellValue(sheetName, 33, 6, null),
					new ExpectedCellValue(sheetName, 33, 7, null),
					new ExpectedCellValue(sheetName, 33, 8, null),
					new ExpectedCellValue(sheetName, 33, 9, null),
					new ExpectedCellValue(sheetName, 33, 10, null),

					// February
					new ExpectedCellValue(sheetName, 35, 3, null),
					new ExpectedCellValue(sheetName, 35, 4, null),
					new ExpectedCellValue(sheetName, 35, 5, null),
					new ExpectedCellValue(sheetName, 35, 6, null),
					new ExpectedCellValue(sheetName, 35, 7, null),
					new ExpectedCellValue(sheetName, 35, 8, null),
					new ExpectedCellValue(sheetName, 35, 9, null),
					new ExpectedCellValue(sheetName, 35, 10, null),

					// March
					new ExpectedCellValue(sheetName, 38, 3, null),
					new ExpectedCellValue(sheetName, 38, 4, null),
					new ExpectedCellValue(sheetName, 38, 5, null),
					new ExpectedCellValue(sheetName, 38, 6, null),
					new ExpectedCellValue(sheetName, 38, 7, null),
					new ExpectedCellValue(sheetName, 38, 8, null),
					new ExpectedCellValue(sheetName, 38, 9, null),
					new ExpectedCellValue(sheetName, 38, 10, null),
				});
			}
		}
		#endregion

		#region PercentOfParent Tests
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableShowDataAs.xlsx")]
		public void PivotTableRefreshShowDataAsPercentOfParentLocationWithRowFieldsColumnDataFields()
		{
			var file = new FileInfo("PivotTableShowDataAs.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				void validateSheet() => TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 2, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 2, 3, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 2, 4, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 3, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 4, 2, "January"),
					new ExpectedCellValue(sheetName, 4, 3, .9433),
					new ExpectedCellValue(sheetName, 4, 4, 2),
					new ExpectedCellValue(sheetName, 5, 2, "March"),
					new ExpectedCellValue(sheetName, 5, 3, .0567),
					new ExpectedCellValue(sheetName, 5, 4, 1),
					new ExpectedCellValue(sheetName, 6, 2, "Nashville"),
					new ExpectedCellValue(sheetName, 7, 2, "January"),
					new ExpectedCellValue(sheetName, 7, 3, .4034),
					new ExpectedCellValue(sheetName, 7, 4, 2),
					new ExpectedCellValue(sheetName, 8, 2, "February"),
					new ExpectedCellValue(sheetName, 8, 3, .1931),
					new ExpectedCellValue(sheetName, 8, 4, 6),
					new ExpectedCellValue(sheetName, 9, 2, "March"),
					new ExpectedCellValue(sheetName, 9, 3, .4034),
					new ExpectedCellValue(sheetName, 9, 4, 2),
					new ExpectedCellValue(sheetName, 10, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 11, 2, "January"),
					new ExpectedCellValue(sheetName, 11, 3, .8077),
					new ExpectedCellValue(sheetName, 11, 4, 1),
					new ExpectedCellValue(sheetName, 12, 2, "February"),
					new ExpectedCellValue(sheetName, 12, 3, .1923),
					new ExpectedCellValue(sheetName, 12, 4, 1),
					new ExpectedCellValue(sheetName, 13, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 13, 3, null),
					new ExpectedCellValue(sheetName, 13, 4, 15)
				});

				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable1"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					wholesalePriceDataField.ShowDataAs = ShowDataAs.PercentOfParent;
					wholesalePriceDataField.BaseField = pivotTable.CacheDefinition.GetCacheFieldIndex("Location");
					unitsSoldDataField.ShowDataAs = ShowDataAs.NoCalculation;
					foreach (var field in pivotTable.Fields)
					{
						field.SubtotalLocation = SubtotalLocation.Top;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B2:D13"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				validateSheet();
				// Validate that top subtotals were written in.
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 3, 3, 1),
					new ExpectedCellValue(sheetName, 3, 4, 3),
					new ExpectedCellValue(sheetName, 6, 3, 1),
					new ExpectedCellValue(sheetName, 6, 4, 10),
					new ExpectedCellValue(sheetName, 10, 3, 1),
					new ExpectedCellValue(sheetName, 10, 4, 2),
				});

				// Test again with subtotals turned off.
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable1"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					wholesalePriceDataField.ShowDataAs = ShowDataAs.PercentOfParent;
					wholesalePriceDataField.BaseField = pivotTable.CacheDefinition.GetCacheFieldIndex("Location");
					unitsSoldDataField.ShowDataAs = ShowDataAs.NoCalculation;
					foreach (var field in pivotTable.Fields)
					{
						field.SubTotalFunctions = eSubTotalFunctions.None;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B2:D13"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				validateSheet();
				// Validate that no subtotals were written in.
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 3, 3, null),
					new ExpectedCellValue(sheetName, 3, 4, null),
					new ExpectedCellValue(sheetName, 6, 3, null),
					new ExpectedCellValue(sheetName, 6, 4, null),
					new ExpectedCellValue(sheetName, 10, 3, null),
					new ExpectedCellValue(sheetName, 10, 4, null),
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableShowDataAs.xlsx")]
		public void PivotTableRefreshShowDataAsPercentOfParentDataFieldWithRowFieldsColumnDataFields()
		{
			var file = new FileInfo("PivotTableShowDataAs.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				void validateSheet() => TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 2, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 2, 3, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 2, 4, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 3, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 4, 2, "January"),
					new ExpectedCellValue(sheetName, 4, 3, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 4, 4, 2),
					new ExpectedCellValue(sheetName, 5, 2, "March"),
					new ExpectedCellValue(sheetName, 5, 3, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 5, 4, 1),
					new ExpectedCellValue(sheetName, 6, 2, "Nashville"),
					new ExpectedCellValue(sheetName, 7, 2, "January"),
					new ExpectedCellValue(sheetName, 7, 3, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 7, 4, 2),
					new ExpectedCellValue(sheetName, 8, 2, "February"),
					new ExpectedCellValue(sheetName, 8, 3, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 8, 4, 6),
					new ExpectedCellValue(sheetName, 9, 2, "March"),
					new ExpectedCellValue(sheetName, 9, 3, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 9, 4, 2),
					new ExpectedCellValue(sheetName, 10, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 11, 2, "January"),
					new ExpectedCellValue(sheetName, 11, 3, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 11, 4, 1),
					new ExpectedCellValue(sheetName, 12, 2, "February"),
					new ExpectedCellValue(sheetName, 12, 3, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 12, 4, 1),
					new ExpectedCellValue(sheetName, 13, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 13, 3, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 13, 4, 15)
				});

				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable1"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					wholesalePriceDataField.ShowDataAs = ShowDataAs.PercentOfParent;
					wholesalePriceDataField.BaseField = pivotTable.CacheDefinition.GetCacheFieldIndex("Wholesale Price");
					unitsSoldDataField.ShowDataAs = ShowDataAs.NoCalculation;
					foreach (var field in pivotTable.Fields)
					{
						field.SubtotalLocation = SubtotalLocation.Top;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B2:D13"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				validateSheet();
				// Validate that top subtotals were written in.
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 3, 3, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 3, 4, 3),
					new ExpectedCellValue(sheetName, 6, 3, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 6, 4, 10),
					new ExpectedCellValue(sheetName, 10, 3, ExcelErrorValue.Create(eErrorType.NA)),
					new ExpectedCellValue(sheetName, 10, 4, 2),
				});

				// Test again with subtotals turned off.
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable1"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					wholesalePriceDataField.ShowDataAs = ShowDataAs.PercentOfParent;
					wholesalePriceDataField.BaseField = pivotTable.CacheDefinition.GetCacheFieldIndex("Wholesale Price");
					unitsSoldDataField.ShowDataAs = ShowDataAs.NoCalculation;
					foreach (var field in pivotTable.Fields)
					{
						field.SubTotalFunctions = eSubTotalFunctions.None;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B2:D13"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				validateSheet();
				// Validate that no subtotals were written in.
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 3, 3, null),
					new ExpectedCellValue(sheetName, 3, 4, null),
					new ExpectedCellValue(sheetName, 6, 3, null),
					new ExpectedCellValue(sheetName, 6, 4, null),
					new ExpectedCellValue(sheetName, 10, 3, null),
					new ExpectedCellValue(sheetName, 10, 4, null),
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableShowDataAs.xlsx")]
		public void PivotTableRefreshShowDataAsPercentOfParentMonthWithColumnFieldsRowDataFields()
		{
			var file = new FileInfo("PivotTableShowDataAs.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable2"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					wholesalePriceDataField.ShowDataAs = ShowDataAs.PercentOfParent;
					wholesalePriceDataField.BaseField = pivotTable.CacheDefinition.GetCacheFieldIndex("Month");
					unitsSoldDataField.ShowDataAs = ShowDataAs.NoCalculation;
					foreach (var field in pivotTable.Fields)
					{
						field.SubTotalFunctions = eSubTotalFunctions.Default;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B18:K22"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 21, 2, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 21, 3, 1),
					new ExpectedCellValue(sheetName, 21, 4, 1),
					new ExpectedCellValue(sheetName, 21, 5, .3322),
					new ExpectedCellValue(sheetName, 21, 6, .6678),
					new ExpectedCellValue(sheetName, 21, 7, 1),
					new ExpectedCellValue(sheetName, 21, 8, .9433),
					new ExpectedCellValue(sheetName, 21, 9, .0567),
					new ExpectedCellValue(sheetName, 21, 10, 1),
					new ExpectedCellValue(sheetName, 21, 11, null),
					new ExpectedCellValue(sheetName, 22, 2, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 22, 3, 5),
					new ExpectedCellValue(sheetName, 22, 4, 5),
					new ExpectedCellValue(sheetName, 22, 5, 1),
					new ExpectedCellValue(sheetName, 22, 6, 6),
					new ExpectedCellValue(sheetName, 22, 7, 7),
					new ExpectedCellValue(sheetName, 22, 8, 2),
					new ExpectedCellValue(sheetName, 22, 9, 1),
					new ExpectedCellValue(sheetName, 22, 10, 3),
					new ExpectedCellValue(sheetName, 22, 11, 15)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableShowDataAs.xlsx")]
		public void PivotTableRefreshShowDataAsPercentOfParentItemWithColumnFieldsRowDataFields()
		{
			var file = new FileInfo("PivotTableShowDataAs.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable2"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					wholesalePriceDataField.ShowDataAs = ShowDataAs.PercentOfParent;
					wholesalePriceDataField.BaseField = pivotTable.CacheDefinition.GetCacheFieldIndex("Item");
					unitsSoldDataField.ShowDataAs = ShowDataAs.NoCalculation;
					foreach (var field in pivotTable.Fields)
					{
						field.SubTotalFunctions = eSubTotalFunctions.Default;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B18:K22"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 21, 2, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 21, 3, 1),
					new ExpectedCellValue(sheetName, 21, 4, null),
					new ExpectedCellValue(sheetName, 21, 5, 1),
					new ExpectedCellValue(sheetName, 21, 6, 1),
					new ExpectedCellValue(sheetName, 21, 7, null),
					new ExpectedCellValue(sheetName, 21, 8, 1),
					new ExpectedCellValue(sheetName, 21, 9, 1),
					new ExpectedCellValue(sheetName, 21, 10, null),
					new ExpectedCellValue(sheetName, 21, 11, null),
					new ExpectedCellValue(sheetName, 22, 2, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 22, 3, 5),
					new ExpectedCellValue(sheetName, 22, 4, 5),
					new ExpectedCellValue(sheetName, 22, 5, 1),
					new ExpectedCellValue(sheetName, 22, 6, 6),
					new ExpectedCellValue(sheetName, 22, 7, 7),
					new ExpectedCellValue(sheetName, 22, 8, 2),
					new ExpectedCellValue(sheetName, 22, 9, 1),
					new ExpectedCellValue(sheetName, 22, 10, 3),
					new ExpectedCellValue(sheetName, 22, 11, 15)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableShowDataAs.xlsx")]
		public void PivotTableRefreshShowDataAsPercentOfParentLocationRowFieldsColumnFieldsAndColumnDataFieldsSubtotalsTop()
		{
			var file = new FileInfo("PivotTableShowDataAs.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable3"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					wholesalePriceDataField.ShowDataAs = ShowDataAs.PercentOfParent;
					wholesalePriceDataField.BaseField = pivotTable.CacheDefinition.GetCacheFieldIndex("Location");
					unitsSoldDataField.ShowDataAs = ShowDataAs.NoCalculation;
					foreach (var field in pivotTable.Fields)
					{
						field.SubtotalLocation = SubtotalLocation.Top;
						field.SubTotalFunctions = eSubTotalFunctions.Default;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B30:J41"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 31, 2, null),
					new ExpectedCellValue(sheetName, 31, 3, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 31, 4, null),
					new ExpectedCellValue(sheetName, 31, 5, null),
					new ExpectedCellValue(sheetName, 31, 6, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 31, 7, null),
					new ExpectedCellValue(sheetName, 31, 8, null),
					new ExpectedCellValue(sheetName, 31, 9, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 31, 10, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 32, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 32, 3, "Chicago"),
					new ExpectedCellValue(sheetName, 32, 4, "Nashville"),
					new ExpectedCellValue(sheetName, 32, 5, "San Francisco"),
					new ExpectedCellValue(sheetName, 32, 6, "Chicago"),
					new ExpectedCellValue(sheetName, 32, 7, "Nashville"),
					new ExpectedCellValue(sheetName, 32, 8, "San Francisco"),
					new ExpectedCellValue(sheetName, 32, 9, null),
					new ExpectedCellValue(sheetName, 32, 10, null),
					new ExpectedCellValue(sheetName, 33, 2, "January"),
					new ExpectedCellValue(sheetName, 33, 3, 1),
					new ExpectedCellValue(sheetName, 33, 4, 1),
					new ExpectedCellValue(sheetName, 33, 5, 1),
					new ExpectedCellValue(sheetName, 33, 6, 2),
					new ExpectedCellValue(sheetName, 33, 7, 2),
					new ExpectedCellValue(sheetName, 33, 8, 1),
					new ExpectedCellValue(sheetName, 33, 9, null),
					new ExpectedCellValue(sheetName, 33, 10, 5),
					new ExpectedCellValue(sheetName, 34, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 34, 3, 1),
					new ExpectedCellValue(sheetName, 34, 4, 1),
					new ExpectedCellValue(sheetName, 34, 5, 1),
					new ExpectedCellValue(sheetName, 34, 6, 2),
					new ExpectedCellValue(sheetName, 34, 7, 2),
					new ExpectedCellValue(sheetName, 34, 8, 1),
					new ExpectedCellValue(sheetName, 34, 9, null),
					new ExpectedCellValue(sheetName, 34, 10, 5),
					new ExpectedCellValue(sheetName, 35, 2, "February"),
					new ExpectedCellValue(sheetName, 35, 3, null),
					new ExpectedCellValue(sheetName, 35, 4, 1),
					new ExpectedCellValue(sheetName, 35, 5, 1),
					new ExpectedCellValue(sheetName, 35, 6, null),
					new ExpectedCellValue(sheetName, 35, 7, 6),
					new ExpectedCellValue(sheetName, 35, 8, 1),
					new ExpectedCellValue(sheetName, 35, 9, null),
					new ExpectedCellValue(sheetName, 35, 10, 7),
					new ExpectedCellValue(sheetName, 36, 2, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 36, 3, null),
					new ExpectedCellValue(sheetName, 36, 4, null),
					new ExpectedCellValue(sheetName, 36, 5, 1),
					new ExpectedCellValue(sheetName, 36, 6, null),
					new ExpectedCellValue(sheetName, 36, 7, null),
					new ExpectedCellValue(sheetName, 36, 8, 1),
					new ExpectedCellValue(sheetName, 36, 9, null),
					new ExpectedCellValue(sheetName, 36, 10, 1),
					new ExpectedCellValue(sheetName, 37, 2, "Tent"),
					new ExpectedCellValue(sheetName, 37, 3, null),
					new ExpectedCellValue(sheetName, 37, 4, 1),
					new ExpectedCellValue(sheetName, 37, 5, null),
					new ExpectedCellValue(sheetName, 37, 6, null),
					new ExpectedCellValue(sheetName, 37, 7, 6),
					new ExpectedCellValue(sheetName, 37, 8, null),
					new ExpectedCellValue(sheetName, 37, 9, null),
					new ExpectedCellValue(sheetName, 37, 10, 6),
					new ExpectedCellValue(sheetName, 38, 2, "March"),
					new ExpectedCellValue(sheetName, 38, 3, 1),
					new ExpectedCellValue(sheetName, 38, 4, 1),
					new ExpectedCellValue(sheetName, 38, 5, null),
					new ExpectedCellValue(sheetName, 38, 6, 1),
					new ExpectedCellValue(sheetName, 38, 7, 2),
					new ExpectedCellValue(sheetName, 38, 8, null),
					new ExpectedCellValue(sheetName, 38, 9, null),
					new ExpectedCellValue(sheetName, 38, 10, 3),
					new ExpectedCellValue(sheetName, 39, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 39, 3, null),
					new ExpectedCellValue(sheetName, 39, 4, 1),
					new ExpectedCellValue(sheetName, 39, 5, null),
					new ExpectedCellValue(sheetName, 39, 6, null),
					new ExpectedCellValue(sheetName, 39, 7, 2),
					new ExpectedCellValue(sheetName, 39, 8, null),
					new ExpectedCellValue(sheetName, 39, 9, null),
					new ExpectedCellValue(sheetName, 39, 10, 2),
					new ExpectedCellValue(sheetName, 40, 2, "Headlamp"),
					new ExpectedCellValue(sheetName, 40, 3, 1),
					new ExpectedCellValue(sheetName, 40, 4, null),
					new ExpectedCellValue(sheetName, 40, 5, null),
					new ExpectedCellValue(sheetName, 40, 6, 1),
					new ExpectedCellValue(sheetName, 40, 7, null),
					new ExpectedCellValue(sheetName, 40, 8, null),
					new ExpectedCellValue(sheetName, 40, 9, null),
					new ExpectedCellValue(sheetName, 40, 10, 1),
					new ExpectedCellValue(sheetName, 41, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 41, 3, 1),
					new ExpectedCellValue(sheetName, 41, 4, 1),
					new ExpectedCellValue(sheetName, 41, 5, 1),
					new ExpectedCellValue(sheetName, 41, 6, 3),
					new ExpectedCellValue(sheetName, 41, 7, 10),
					new ExpectedCellValue(sheetName, 41, 8, 2),
					new ExpectedCellValue(sheetName, 41, 9, null),
					new ExpectedCellValue(sheetName, 41, 10, 15)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableShowDataAs.xlsx")]
		public void PivotTableRefreshShowDataAsPercentOfParentLocationRowFieldsColumnFieldsAndColumnDataFieldsSubtotalsBottom()
		{
			var file = new FileInfo("PivotTableShowDataAs.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable3"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					wholesalePriceDataField.ShowDataAs = ShowDataAs.PercentOfParent;
					wholesalePriceDataField.BaseField = pivotTable.CacheDefinition.GetCacheFieldIndex("Month");
					unitsSoldDataField.ShowDataAs = ShowDataAs.NoCalculation;
					foreach (var field in pivotTable.Fields)
					{
						field.SubtotalLocation = SubtotalLocation.Bottom;
						field.SubTotalFunctions = eSubTotalFunctions.Default;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B30:J44"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 31, 2, null),
					new ExpectedCellValue(sheetName, 31, 3, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 31, 4, null),
					new ExpectedCellValue(sheetName, 31, 5, null),
					new ExpectedCellValue(sheetName, 31, 6, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 31, 7, null),
					new ExpectedCellValue(sheetName, 31, 8, null),
					new ExpectedCellValue(sheetName, 31, 9, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 31, 10, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 32, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 32, 3, "Chicago"),
					new ExpectedCellValue(sheetName, 32, 4, "Nashville"),
					new ExpectedCellValue(sheetName, 32, 5, "San Francisco"),
					new ExpectedCellValue(sheetName, 32, 6, "Chicago"),
					new ExpectedCellValue(sheetName, 32, 7, "Nashville"),
					new ExpectedCellValue(sheetName, 32, 8, "San Francisco"),
					new ExpectedCellValue(sheetName, 32, 9, null),
					new ExpectedCellValue(sheetName, 32, 10, null),
					new ExpectedCellValue(sheetName, 33, 2, "January"),
					new ExpectedCellValue(sheetName, 33, 3, null),
					new ExpectedCellValue(sheetName, 33, 4, null),
					new ExpectedCellValue(sheetName, 33, 5, null),
					new ExpectedCellValue(sheetName, 33, 6, null),
					new ExpectedCellValue(sheetName, 33, 7, null),
					new ExpectedCellValue(sheetName, 33, 8, null),
					new ExpectedCellValue(sheetName, 33, 9, null),
					new ExpectedCellValue(sheetName, 33, 10, null),
					new ExpectedCellValue(sheetName, 34, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 34, 3, 1),
					new ExpectedCellValue(sheetName, 34, 4, 1),
					new ExpectedCellValue(sheetName, 34, 5, 1),
					new ExpectedCellValue(sheetName, 34, 6, 2),
					new ExpectedCellValue(sheetName, 34, 7, 2),
					new ExpectedCellValue(sheetName, 34, 8, 1),
					new ExpectedCellValue(sheetName, 34, 9, 1),
					new ExpectedCellValue(sheetName, 34, 10, 5),
					new ExpectedCellValue(sheetName, 35, 2, "January Total"),
					new ExpectedCellValue(sheetName, 35, 3, 1),
					new ExpectedCellValue(sheetName, 35, 4, 1),
					new ExpectedCellValue(sheetName, 35, 5, 1),
					new ExpectedCellValue(sheetName, 35, 6, 2),
					new ExpectedCellValue(sheetName, 35, 7, 2),
					new ExpectedCellValue(sheetName, 35, 8, 1),
					new ExpectedCellValue(sheetName, 35, 9, 1),
					new ExpectedCellValue(sheetName, 35, 10, 5),
					new ExpectedCellValue(sheetName, 36, 2, "February"),
					new ExpectedCellValue(sheetName, 36, 3, null),
					new ExpectedCellValue(sheetName, 36, 4, null),
					new ExpectedCellValue(sheetName, 36, 5, null),
					new ExpectedCellValue(sheetName, 36, 6, null),
					new ExpectedCellValue(sheetName, 36, 7, null),
					new ExpectedCellValue(sheetName, 36, 8, null),
					new ExpectedCellValue(sheetName, 36, 9, null),
					new ExpectedCellValue(sheetName, 36, 10, null),
					new ExpectedCellValue(sheetName, 37, 2, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 37, 3, null),
					new ExpectedCellValue(sheetName, 37, 4, 0),
					new ExpectedCellValue(sheetName, 37, 5, 1),
					new ExpectedCellValue(sheetName, 37, 6, null),
					new ExpectedCellValue(sheetName, 37, 7, null),
					new ExpectedCellValue(sheetName, 37, 8, 1),
					new ExpectedCellValue(sheetName, 37, 9, .3322),
					new ExpectedCellValue(sheetName, 37, 10, 1),
					new ExpectedCellValue(sheetName, 38, 2, "Tent"),
					new ExpectedCellValue(sheetName, 38, 3, null),
					new ExpectedCellValue(sheetName, 38, 4, 1),
					new ExpectedCellValue(sheetName, 38, 5, 0),
					new ExpectedCellValue(sheetName, 38, 6, null),
					new ExpectedCellValue(sheetName, 38, 7, 6),
					new ExpectedCellValue(sheetName, 38, 8, null),
					new ExpectedCellValue(sheetName, 38, 9, .6678),
					new ExpectedCellValue(sheetName, 38, 10, 6),
					new ExpectedCellValue(sheetName, 39, 2, "February Total"),
					new ExpectedCellValue(sheetName, 39, 3, null),
					new ExpectedCellValue(sheetName, 39, 4, 1),
					new ExpectedCellValue(sheetName, 39, 5, 1),
					new ExpectedCellValue(sheetName, 39, 6, null),
					new ExpectedCellValue(sheetName, 39, 7, 6),
					new ExpectedCellValue(sheetName, 39, 8, 1),
					new ExpectedCellValue(sheetName, 39, 9, 1),
					new ExpectedCellValue(sheetName, 39, 10, 7),
					new ExpectedCellValue(sheetName, 40, 2, "March"),
					new ExpectedCellValue(sheetName, 40, 3, null),
					new ExpectedCellValue(sheetName, 40, 4, null),
					new ExpectedCellValue(sheetName, 40, 5, null),
					new ExpectedCellValue(sheetName, 40, 6, null),
					new ExpectedCellValue(sheetName, 40, 7, null),
					new ExpectedCellValue(sheetName, 40, 8, null),
					new ExpectedCellValue(sheetName, 40, 9, null),
					new ExpectedCellValue(sheetName, 40, 10, null),
					new ExpectedCellValue(sheetName, 41, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 41, 3, 0),
					new ExpectedCellValue(sheetName, 41, 4, 1),
					new ExpectedCellValue(sheetName, 41, 5, null),
					new ExpectedCellValue(sheetName, 41, 6, null),
					new ExpectedCellValue(sheetName, 41, 7, 2),
					new ExpectedCellValue(sheetName, 41, 8, null),
					new ExpectedCellValue(sheetName, 41, 9, .9433),
					new ExpectedCellValue(sheetName, 41, 10, 2),
					new ExpectedCellValue(sheetName, 42, 2, "Headlamp"),
					new ExpectedCellValue(sheetName, 42, 3, 1),
					new ExpectedCellValue(sheetName, 42, 4, 0),
					new ExpectedCellValue(sheetName, 42, 5, null),
					new ExpectedCellValue(sheetName, 42, 6, 1),
					new ExpectedCellValue(sheetName, 42, 7, null),
					new ExpectedCellValue(sheetName, 42, 8, null),
					new ExpectedCellValue(sheetName, 42, 9, .0567),
					new ExpectedCellValue(sheetName, 42, 10, 1),
					new ExpectedCellValue(sheetName, 43, 2, "March Total"),
					new ExpectedCellValue(sheetName, 43, 3, 1),
					new ExpectedCellValue(sheetName, 43, 4, 1),
					new ExpectedCellValue(sheetName, 43, 5, null),
					new ExpectedCellValue(sheetName, 43, 6, 1),
					new ExpectedCellValue(sheetName, 43, 7, 2),
					new ExpectedCellValue(sheetName, 43, 8, null),
					new ExpectedCellValue(sheetName, 43, 9, 1),
					new ExpectedCellValue(sheetName, 43, 10, 3),
					new ExpectedCellValue(sheetName, 44, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 44, 3, null),
					new ExpectedCellValue(sheetName, 44, 4, null),
					new ExpectedCellValue(sheetName, 44, 5, null),
					new ExpectedCellValue(sheetName, 44, 6, 3),
					new ExpectedCellValue(sheetName, 44, 7, 10),
					new ExpectedCellValue(sheetName, 44, 8, 2),
					new ExpectedCellValue(sheetName, 44, 9, null),
					new ExpectedCellValue(sheetName, 44, 10, 15)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableShowDataAs.xlsx")]
		public void PivotTableRefreshShowDataAsPercentOfParentMonthRowFieldsColumnFieldsAndRowDataFields()
		{
			var file = new FileInfo("PivotTableShowDataAs.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable4"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					wholesalePriceDataField.ShowDataAs = ShowDataAs.NoCalculation;
					unitsSoldDataField.ShowDataAs = ShowDataAs.PercentOfParent;
					wholesalePriceDataField.BaseField = pivotTable.CacheDefinition.GetCacheFieldIndex("Month");
					foreach (var field in pivotTable.Fields)
					{
						field.SubTotalFunctions = eSubTotalFunctions.None;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B46:F67"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 47, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 47, 3, "Chicago"),
					new ExpectedCellValue(sheetName, 47, 4, "Nashville"),
					new ExpectedCellValue(sheetName, 47, 5, "San Francisco"),
					new ExpectedCellValue(sheetName, 47, 6, "Grand Total"),
					new ExpectedCellValue(sheetName, 48, 2, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 48, 3, null),
					new ExpectedCellValue(sheetName, 48, 6, null),
					new ExpectedCellValue(sheetName, 49, 2, "January"),
					new ExpectedCellValue(sheetName, 49, 3, null),
					new ExpectedCellValue(sheetName, 49, 6, null),
					new ExpectedCellValue(sheetName, 50, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 50, 3, 1),
					new ExpectedCellValue(sheetName, 50, 4, 1),
					new ExpectedCellValue(sheetName, 50, 5, 1),
					new ExpectedCellValue(sheetName, 50, 6, 1),
					new ExpectedCellValue(sheetName, 51, 2, "February"),
					new ExpectedCellValue(sheetName, 51, 4, null),
					new ExpectedCellValue(sheetName, 51, 5, null),
					new ExpectedCellValue(sheetName, 51, 6, null),
					new ExpectedCellValue(sheetName, 52, 2, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 52, 3, null),
					new ExpectedCellValue(sheetName, 52, 4, 0),
					new ExpectedCellValue(sheetName, 52, 5, 1),
					new ExpectedCellValue(sheetName, 52, 6, .1429),
					new ExpectedCellValue(sheetName, 53, 2, "Tent"),
					new ExpectedCellValue(sheetName, 53, 3, null),
					new ExpectedCellValue(sheetName, 53, 4, 1),
					new ExpectedCellValue(sheetName, 53, 5, 0),
					new ExpectedCellValue(sheetName, 53, 6, .8571),
					new ExpectedCellValue(sheetName, 54, 2, "March"),
					new ExpectedCellValue(sheetName, 54, 3, null),
					new ExpectedCellValue(sheetName, 54, 6, null),
					new ExpectedCellValue(sheetName, 55, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 55, 3, 0),
					new ExpectedCellValue(sheetName, 55, 4, 1),
					new ExpectedCellValue(sheetName, 55, 5, null),
					new ExpectedCellValue(sheetName, 55, 6, .6667),
					new ExpectedCellValue(sheetName, 56, 2, "Headlamp"),
					new ExpectedCellValue(sheetName, 56, 3, 1),
					new ExpectedCellValue(sheetName, 56, 4, 0),
					new ExpectedCellValue(sheetName, 56, 5, null),
					new ExpectedCellValue(sheetName, 56, 6, .3333),
					new ExpectedCellValue(sheetName, 57, 2, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 57, 3, null),
					new ExpectedCellValue(sheetName, 57, 6, null),
					new ExpectedCellValue(sheetName, 58, 2, "January"),
					new ExpectedCellValue(sheetName, 58, 3, null),
					new ExpectedCellValue(sheetName, 58, 4, null),
					new ExpectedCellValue(sheetName, 58, 6, null),
					new ExpectedCellValue(sheetName, 59, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 59, 3, 415.75),
					new ExpectedCellValue(sheetName, 59, 4, 415.75),
					new ExpectedCellValue(sheetName, 59, 5, 415.75),
					new ExpectedCellValue(sheetName, 59, 6, 1247.25),
					new ExpectedCellValue(sheetName, 60, 2, "February"),
					new ExpectedCellValue(sheetName, 60, 3, null),
					new ExpectedCellValue(sheetName, 60, 5, null),
					new ExpectedCellValue(sheetName, 61, 2, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 61, 3, null),
					new ExpectedCellValue(sheetName, 61, 4, null),
					new ExpectedCellValue(sheetName, 61, 5, 99),
					new ExpectedCellValue(sheetName, 61, 6, 99),
					new ExpectedCellValue(sheetName, 62, 2, "Tent"),
					new ExpectedCellValue(sheetName, 62, 3, null),
					new ExpectedCellValue(sheetName, 62, 4, 199),
					new ExpectedCellValue(sheetName, 62, 5, null),
					new ExpectedCellValue(sheetName, 62, 6, 199),
					new ExpectedCellValue(sheetName, 63, 2, "March"),
					new ExpectedCellValue(sheetName, 63, 5, null),
					new ExpectedCellValue(sheetName, 63, 6, null),
					new ExpectedCellValue(sheetName, 64, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 64, 3, null),
					new ExpectedCellValue(sheetName, 64, 4, 415.75),
					new ExpectedCellValue(sheetName, 64, 5, null),
					new ExpectedCellValue(sheetName, 64, 6, 415.75),
					new ExpectedCellValue(sheetName, 65, 2, "Headlamp"),
					new ExpectedCellValue(sheetName, 65, 3, 24.99),
					new ExpectedCellValue(sheetName, 65, 4, null),
					new ExpectedCellValue(sheetName, 65, 5, null),
					new ExpectedCellValue(sheetName, 65, 6, 24.99),
					new ExpectedCellValue(sheetName, 66, 2, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 66, 3, null),
					new ExpectedCellValue(sheetName, 66, 4, null),
					new ExpectedCellValue(sheetName, 66, 5, null),
					new ExpectedCellValue(sheetName, 66, 6, null),
					new ExpectedCellValue(sheetName, 67, 2, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 67, 3, 440.74),
					new ExpectedCellValue(sheetName, 67, 4, 1030.5),
					new ExpectedCellValue(sheetName, 67, 5, 514.75),
					new ExpectedCellValue(sheetName, 67, 6, 1985.99)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableShowDataAs.xlsx")]
		public void PivotTableRefreshShowDataAsPercentOfParentLocationRowFieldsColumnFieldsAndRowDataFields()
		{
			var file = new FileInfo("PivotTableShowDataAs.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable4"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					wholesalePriceDataField.ShowDataAs = ShowDataAs.NoCalculation;
					unitsSoldDataField.ShowDataAs = ShowDataAs.PercentOfParent;
					unitsSoldDataField.BaseField = pivotTable.CacheDefinition.GetCacheFieldIndex("Location");
					foreach (var field in pivotTable.Fields)
					{
						field.SubTotalFunctions = eSubTotalFunctions.None;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B46:F67"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 47, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 47, 3, "Chicago"),
					new ExpectedCellValue(sheetName, 47, 4, "Nashville"),
					new ExpectedCellValue(sheetName, 47, 5, "San Francisco"),
					new ExpectedCellValue(sheetName, 47, 6, "Grand Total"),
					new ExpectedCellValue(sheetName, 48, 2, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 48, 3, null),
					new ExpectedCellValue(sheetName, 48, 6, null),
					new ExpectedCellValue(sheetName, 49, 2, "January"),
					new ExpectedCellValue(sheetName, 49, 3, null),
					new ExpectedCellValue(sheetName, 49, 6, null),
					new ExpectedCellValue(sheetName, 50, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 50, 3, 1),
					new ExpectedCellValue(sheetName, 50, 4, 1),
					new ExpectedCellValue(sheetName, 50, 5, 1),
					new ExpectedCellValue(sheetName, 50, 6, null),
					new ExpectedCellValue(sheetName, 51, 2, "February"),
					new ExpectedCellValue(sheetName, 51, 4, null),
					new ExpectedCellValue(sheetName, 51, 5, null),
					new ExpectedCellValue(sheetName, 51, 6, null),
					new ExpectedCellValue(sheetName, 52, 2, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 52, 3, null),
					new ExpectedCellValue(sheetName, 52, 4, null),
					new ExpectedCellValue(sheetName, 52, 5, 1),
					new ExpectedCellValue(sheetName, 52, 6, null),
					new ExpectedCellValue(sheetName, 53, 2, "Tent"),
					new ExpectedCellValue(sheetName, 53, 3, null),
					new ExpectedCellValue(sheetName, 53, 4, 1),
					new ExpectedCellValue(sheetName, 53, 5, null),
					new ExpectedCellValue(sheetName, 53, 6, null),
					new ExpectedCellValue(sheetName, 54, 2, "March"),
					new ExpectedCellValue(sheetName, 54, 3, null),
					new ExpectedCellValue(sheetName, 54, 6, null),
					new ExpectedCellValue(sheetName, 55, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 55, 3, null),
					new ExpectedCellValue(sheetName, 55, 4, 1),
					new ExpectedCellValue(sheetName, 55, 5, null),
					new ExpectedCellValue(sheetName, 55, 6, null),
					new ExpectedCellValue(sheetName, 56, 2, "Headlamp"),
					new ExpectedCellValue(sheetName, 56, 3, 1),
					new ExpectedCellValue(sheetName, 56, 4, null),
					new ExpectedCellValue(sheetName, 56, 5, null),
					new ExpectedCellValue(sheetName, 56, 6, null),
					new ExpectedCellValue(sheetName, 57, 2, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 57, 3, null),
					new ExpectedCellValue(sheetName, 57, 6, null),
					new ExpectedCellValue(sheetName, 58, 2, "January"),
					new ExpectedCellValue(sheetName, 58, 3, null),
					new ExpectedCellValue(sheetName, 58, 4, null),
					new ExpectedCellValue(sheetName, 58, 6, null),
					new ExpectedCellValue(sheetName, 59, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 59, 3, 415.75),
					new ExpectedCellValue(sheetName, 59, 4, 415.75),
					new ExpectedCellValue(sheetName, 59, 5, 415.75),
					new ExpectedCellValue(sheetName, 59, 6, 1247.25),
					new ExpectedCellValue(sheetName, 60, 2, "February"),
					new ExpectedCellValue(sheetName, 60, 3, null),
					new ExpectedCellValue(sheetName, 60, 5, null),
					new ExpectedCellValue(sheetName, 61, 2, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 61, 3, null),
					new ExpectedCellValue(sheetName, 61, 4, null),
					new ExpectedCellValue(sheetName, 61, 5, 99),
					new ExpectedCellValue(sheetName, 61, 6, 99),
					new ExpectedCellValue(sheetName, 62, 2, "Tent"),
					new ExpectedCellValue(sheetName, 62, 3, null),
					new ExpectedCellValue(sheetName, 62, 4, 199),
					new ExpectedCellValue(sheetName, 62, 5, null),
					new ExpectedCellValue(sheetName, 62, 6, 199),
					new ExpectedCellValue(sheetName, 63, 2, "March"),
					new ExpectedCellValue(sheetName, 63, 5, null),
					new ExpectedCellValue(sheetName, 63, 6, null),
					new ExpectedCellValue(sheetName, 64, 2, "Car Rack"),
					new ExpectedCellValue(sheetName, 64, 3, null),
					new ExpectedCellValue(sheetName, 64, 4, 415.75),
					new ExpectedCellValue(sheetName, 64, 5, null),
					new ExpectedCellValue(sheetName, 64, 6, 415.75),
					new ExpectedCellValue(sheetName, 65, 2, "Headlamp"),
					new ExpectedCellValue(sheetName, 65, 3, 24.99),
					new ExpectedCellValue(sheetName, 65, 4, null),
					new ExpectedCellValue(sheetName, 65, 5, null),
					new ExpectedCellValue(sheetName, 65, 6, 24.99),
					new ExpectedCellValue(sheetName, 66, 2, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 66, 3, 1),
					new ExpectedCellValue(sheetName, 66, 4, 1),
					new ExpectedCellValue(sheetName, 66, 5, 1),
					new ExpectedCellValue(sheetName, 66, 6, null),
					new ExpectedCellValue(sheetName, 67, 2, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 67, 3, 440.74),
					new ExpectedCellValue(sheetName, 67, 4, 1030.5),
					new ExpectedCellValue(sheetName, 67, 5, 514.75),
					new ExpectedCellValue(sheetName, 67, 6, 1985.99)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableShowDataAs.xlsx")]
		public void PivotTableRefreshShowDataAsPercentOfParentMonthRowDataFieldsAndRowFieldsColumnFieldsSubtotalsBottom()
		{
			var file = new FileInfo("PivotTableShowDataAs.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable6"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					wholesalePriceDataField.ShowDataAs = ShowDataAs.PercentOfParent;
					wholesalePriceDataField.BaseField = pivotTable.CacheDefinition.GetCacheFieldIndex("Month");
					unitsSoldDataField.ShowDataAs = ShowDataAs.NoCalculation;
					foreach (var field in pivotTable.Fields)
					{
						field.SubtotalLocation = SubtotalLocation.Bottom;
						field.SubTotalFunctions = eSubTotalFunctions.Default;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B76:K88"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 79, 2, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 79, 3, null),
					new ExpectedCellValue(sheetName, 79, 4, null),
					new ExpectedCellValue(sheetName, 79, 5, null),
					new ExpectedCellValue(sheetName, 79, 6, null),
					new ExpectedCellValue(sheetName, 79, 7, null),
					new ExpectedCellValue(sheetName, 79, 8, null),
					new ExpectedCellValue(sheetName, 79, 9, null),
					new ExpectedCellValue(sheetName, 79, 10, null),
					new ExpectedCellValue(sheetName, 79, 11, null),

					new ExpectedCellValue(sheetName, 80, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 80, 3, 1),
					new ExpectedCellValue(sheetName, 80, 4, 1),
					new ExpectedCellValue(sheetName, 80, 5, null),
					new ExpectedCellValue(sheetName, 80, 6, null),
					new ExpectedCellValue(sheetName, 80, 7, null),
					new ExpectedCellValue(sheetName, 80, 8, 0),
					new ExpectedCellValue(sheetName, 80, 9, 1),
					new ExpectedCellValue(sheetName, 80, 10, 1),
					new ExpectedCellValue(sheetName, 80, 11, null),

					new ExpectedCellValue(sheetName, 81, 2, "Nashville"),
					new ExpectedCellValue(sheetName, 81, 3, 1),
					new ExpectedCellValue(sheetName, 81, 4, 1),
					new ExpectedCellValue(sheetName, 81, 5, 0),
					new ExpectedCellValue(sheetName, 81, 6, 1),
					new ExpectedCellValue(sheetName, 81, 7, 1),
					new ExpectedCellValue(sheetName, 81, 8, 1),
					new ExpectedCellValue(sheetName, 81, 9, 0),
					new ExpectedCellValue(sheetName, 81, 10, 1),
					new ExpectedCellValue(sheetName, 81, 11, null),

					new ExpectedCellValue(sheetName, 82, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 82, 3, 1),
					new ExpectedCellValue(sheetName, 82, 4, 1),
					new ExpectedCellValue(sheetName, 82, 5, 1),
					new ExpectedCellValue(sheetName, 82, 6, 0),
					new ExpectedCellValue(sheetName, 82, 7, 1),
					new ExpectedCellValue(sheetName, 82, 8, null),
					new ExpectedCellValue(sheetName, 82, 9, null),
					new ExpectedCellValue(sheetName, 82, 10, null),
					new ExpectedCellValue(sheetName, 82, 11, null),

					new ExpectedCellValue(sheetName, 83, 2, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 83, 3, null),
					new ExpectedCellValue(sheetName, 83, 4, null),
					new ExpectedCellValue(sheetName, 83, 5, null),
					new ExpectedCellValue(sheetName, 83, 6, null),
					new ExpectedCellValue(sheetName, 83, 7, null),
					new ExpectedCellValue(sheetName, 83, 8, null),
					new ExpectedCellValue(sheetName, 83, 9, null),
					new ExpectedCellValue(sheetName, 83, 10, null),
					new ExpectedCellValue(sheetName, 83, 11, null),

					new ExpectedCellValue(sheetName, 84, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 84, 3, 2),
					new ExpectedCellValue(sheetName, 84, 4, 2),
					new ExpectedCellValue(sheetName, 84, 5, null),
					new ExpectedCellValue(sheetName, 84, 6, null),
					new ExpectedCellValue(sheetName, 84, 7, null),
					new ExpectedCellValue(sheetName, 84, 8, null),
					new ExpectedCellValue(sheetName, 84, 9, 1),
					new ExpectedCellValue(sheetName, 84, 10, 1),
					new ExpectedCellValue(sheetName, 84, 11, 3),

					new ExpectedCellValue(sheetName, 85, 2, "Nashville"),
					new ExpectedCellValue(sheetName, 85, 3, 2),
					new ExpectedCellValue(sheetName, 85, 4, 2),
					new ExpectedCellValue(sheetName, 85, 5, null),
					new ExpectedCellValue(sheetName, 85, 6, 6),
					new ExpectedCellValue(sheetName, 85, 7, 6),
					new ExpectedCellValue(sheetName, 85, 8, 2),
					new ExpectedCellValue(sheetName, 85, 9, null),
					new ExpectedCellValue(sheetName, 85, 10, 2),
					new ExpectedCellValue(sheetName, 85, 11, 10),

					new ExpectedCellValue(sheetName, 86, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 86, 3, 1),
					new ExpectedCellValue(sheetName, 86, 4, 1),
					new ExpectedCellValue(sheetName, 86, 5, 1),
					new ExpectedCellValue(sheetName, 86, 6, null),
					new ExpectedCellValue(sheetName, 86, 7, 1),
					new ExpectedCellValue(sheetName, 86, 8, null),
					new ExpectedCellValue(sheetName, 86, 9, null),
					new ExpectedCellValue(sheetName, 86, 10, null),
					new ExpectedCellValue(sheetName, 86, 11, 2),

					new ExpectedCellValue(sheetName, 87, 2, "Total Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 87, 3, 1),
					new ExpectedCellValue(sheetName, 87, 4, 1),
					new ExpectedCellValue(sheetName, 87, 5, .3322),
					new ExpectedCellValue(sheetName, 87, 6, .6678),
					new ExpectedCellValue(sheetName, 87, 7, 1),
					new ExpectedCellValue(sheetName, 87, 8, .9433),
					new ExpectedCellValue(sheetName, 87, 9, .0567),
					new ExpectedCellValue(sheetName, 87, 10, 1),
					new ExpectedCellValue(sheetName, 87, 11, null),

					new ExpectedCellValue(sheetName, 88, 2, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 88, 3, 5),
					new ExpectedCellValue(sheetName, 88, 4, 5),
					new ExpectedCellValue(sheetName, 88, 5, 1),
					new ExpectedCellValue(sheetName, 88, 6, 6),
					new ExpectedCellValue(sheetName, 88, 7, 7),
					new ExpectedCellValue(sheetName, 88, 8, 2),
					new ExpectedCellValue(sheetName, 88, 9, 1),
					new ExpectedCellValue(sheetName, 88, 10, 3),
					new ExpectedCellValue(sheetName, 88, 11, 15)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableShowDataAs.xlsx")]
		public void PivotTableRefreshShowDataAsPercentOfParentWithDiv0()
		{
			var file = new FileInfo("PivotTableShowDataAs.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				using (var package = new ExcelPackage(file))
				{
					var dataSheet = package.Workbook.Worksheets["Sheet1"];
					dataSheet.Cells["F5:F6"].Value = 0;

					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable1"];

					var wholesalePriceDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Wholesale Price");
					var unitsSoldDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Units Sold");
					wholesalePriceDataField.ShowDataAs = ShowDataAs.PercentOfParentRow;
					unitsSoldDataField.ShowDataAs = ShowDataAs.NoCalculation;
					foreach (var field in pivotTable.Fields)
					{
						field.SubtotalLocation = SubtotalLocation.Top;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B2:D13"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 2, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 2, 3, "Sum of Wholesale Price"),
					new ExpectedCellValue(sheetName, 2, 4, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 3, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 3, 3, 0),
					new ExpectedCellValue(sheetName, 3, 4, 3),
					new ExpectedCellValue(sheetName, 4, 2, "January"),
					new ExpectedCellValue(sheetName, 4, 3, null),
					new ExpectedCellValue(sheetName, 4, 4, 2),
					new ExpectedCellValue(sheetName, 5, 2, "March"),
					new ExpectedCellValue(sheetName, 5, 3, null),
					new ExpectedCellValue(sheetName, 5, 4, 1),
					new ExpectedCellValue(sheetName, 6, 2, "Nashville"),
					new ExpectedCellValue(sheetName, 6, 3, .6669),
					new ExpectedCellValue(sheetName, 6, 4, 10),
					new ExpectedCellValue(sheetName, 7, 2, "January"),
					new ExpectedCellValue(sheetName, 7, 3, .4034),
					new ExpectedCellValue(sheetName, 7, 4, 2),
					new ExpectedCellValue(sheetName, 8, 2, "February"),
					new ExpectedCellValue(sheetName, 8, 3, .1931),
					new ExpectedCellValue(sheetName, 8, 4, 6),
					new ExpectedCellValue(sheetName, 9, 2, "March"),
					new ExpectedCellValue(sheetName, 9, 3, .4034),
					new ExpectedCellValue(sheetName, 9, 4, 2),
					new ExpectedCellValue(sheetName, 10, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 10, 3, .3331),
					new ExpectedCellValue(sheetName, 10, 4, 2),
					new ExpectedCellValue(sheetName, 11, 2, "January"),
					new ExpectedCellValue(sheetName, 11, 3, .8077),
					new ExpectedCellValue(sheetName, 11, 4, 1),
					new ExpectedCellValue(sheetName, 12, 2, "February"),
					new ExpectedCellValue(sheetName, 12, 3, .1923),
					new ExpectedCellValue(sheetName, 12, 4, 1),
					new ExpectedCellValue(sheetName, 13, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 13, 3, 1),
					new ExpectedCellValue(sheetName, 13, 4, 15)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableShowDataAs.xlsx")]
		public void PivotTableRefreshShowDataAsPercentOfParentColumnTotal()
		{
			var file = new FileInfo("PivotTableShowDataAs.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				using (var package = new ExcelPackage(file))
				{
					var dataSheet = package.Workbook.Worksheets[sheetName];
					dataSheet.Cells["B100:L106"].Value = 0;

					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable8"];

					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B100:L106"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}

				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 100, 3, "Column Labels"),
					new ExpectedCellValue(sheetName, 101, 3, "Sum of Total"),
					new ExpectedCellValue(sheetName, 101, 7, "Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 101, 11, "Total Sum of Total"),
					new ExpectedCellValue(sheetName, 101, 12, "Total Sum of Units Sold"),
					new ExpectedCellValue(sheetName, 102, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 102, 3, "Car Rack"),
					new ExpectedCellValue(sheetName, 102, 4, "Headlamp"),
					new ExpectedCellValue(sheetName, 102, 5, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 102, 6, "Tent"),
					new ExpectedCellValue(sheetName, 102, 7, "Car Rack"),
					new ExpectedCellValue(sheetName, 102, 8, "Headlamp"),
					new ExpectedCellValue(sheetName, 102, 9, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 102, 10, "Tent"),
					new ExpectedCellValue(sheetName, 103, 2, "January"),
					new ExpectedCellValue(sheetName, 103, 3, 1),
					new ExpectedCellValue(sheetName, 103, 4, 0),
					new ExpectedCellValue(sheetName, 103, 5, 0),
					new ExpectedCellValue(sheetName, 103, 6, 0),
					new ExpectedCellValue(sheetName, 103, 7, 5),
					new ExpectedCellValue(sheetName, 103, 8, null),
					new ExpectedCellValue(sheetName, 103, 9, null),
					new ExpectedCellValue(sheetName, 103, 10, null),
					new ExpectedCellValue(sheetName, 103, 11, 1),
					new ExpectedCellValue(sheetName, 103, 12, 5),
					new ExpectedCellValue(sheetName, 104, 2, "February"),
					new ExpectedCellValue(sheetName, 104, 3, 0),
					new ExpectedCellValue(sheetName, 104, 4, 0),
					new ExpectedCellValue(sheetName, 104, 5, 0.0765661252900232),
					new ExpectedCellValue(sheetName, 104, 6, 0.923433874709977),
					new ExpectedCellValue(sheetName, 104, 7, null),
					new ExpectedCellValue(sheetName, 104, 8, null),
					new ExpectedCellValue(sheetName, 104, 9, 1),
					new ExpectedCellValue(sheetName, 104, 10, 6),
					new ExpectedCellValue(sheetName, 104, 11, 1),
					new ExpectedCellValue(sheetName, 104, 12, 7),
					new ExpectedCellValue(sheetName, 105, 2, "March"),
					new ExpectedCellValue(sheetName, 105, 3, 0.970822776681572),
					new ExpectedCellValue(sheetName, 105, 4, 0.0291772233184275),
					new ExpectedCellValue(sheetName, 105, 5, 0),
					new ExpectedCellValue(sheetName, 105, 6, 0),
					new ExpectedCellValue(sheetName, 105, 7, 2),
					new ExpectedCellValue(sheetName, 105, 8, 1),
					new ExpectedCellValue(sheetName, 105, 9, null),
					new ExpectedCellValue(sheetName, 105, 10, null),
					new ExpectedCellValue(sheetName, 105, 11, 1),
					new ExpectedCellValue(sheetName, 105, 12, 3),
					new ExpectedCellValue(sheetName, 106, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 106, 3, 0.688288744252928),
					new ExpectedCellValue(sheetName, 106, 4, 0.00591026053393374),
					new ExpectedCellValue(sheetName, 106, 5, 0.0234139973133029),
					new ExpectedCellValue(sheetName, 106, 6, 0.282386997899835),
					new ExpectedCellValue(sheetName, 106, 7, 7),
					new ExpectedCellValue(sheetName, 106, 8, 1),
					new ExpectedCellValue(sheetName, 106, 9, 1),
					new ExpectedCellValue(sheetName, 106, 10, 6),
					new ExpectedCellValue(sheetName, 106, 11, 1),
					new ExpectedCellValue(sheetName, 106, 12, 15),
				});
			}
		}

		[TestMethod]
		public void AutoGenerateExpectedResults()
		{
			string sheetName = "PivotTables";
			string range = "B100:L106";
			var sourceFilePath = @"C:\repos\EPPlus\EPPlusTest\Workbooks\PivotTables\PivotTableShowDataAs.xlsx";
			var outputFilePath = @"C:\Users\rwf\Downloads\expected.cs";

			using (var package = new ExcelPackage(new FileInfo(sourceFilePath)))
			{
				var cells = package.Workbook.Worksheets[sheetName].Cells[range];
				string text = $"TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]{Environment.NewLine}{{{Environment.NewLine}";
				foreach (var cell in cells)
				{
					string value = null;
					if (cell.Value is string)
						value = $"\"{cell.Value}\"";
					else if (cell.Value is ExcelErrorValue errorValue)
						value = $"ExcelErrorValue.Create(eErrorType.{errorValue.Type})";
					else if (cell.Value == null)
						value = "null";
					else
						value = cell.Value.ToString();

					text += $"	new ExpectedCellValue(sheetName, {cell._fromRow}, {cell._fromCol}, {value}),{Environment.NewLine}";
				}
				text += "});";
				File.WriteAllText(outputFilePath, text);
			}
		}
		#endregion
		#endregion

		#region Helper Methods
		private int GetCacheFieldSharedItemIndex(ExcelPivotTable pivotTable, int fieldIndex, string value)
		{
			// Get the index of the pivot field item that matches the specified value in the respective cache field shared item.
			int i = 0;
			var cacheField = pivotTable.CacheDefinition.CacheFields[fieldIndex];
			foreach (var item in pivotTable.Fields[fieldIndex].Items)
			{
				if (cacheField.SharedItems[item.X].Value == value)
					return i;
				i++;
			}
			return -1;
		}
		#endregion
	}
}
