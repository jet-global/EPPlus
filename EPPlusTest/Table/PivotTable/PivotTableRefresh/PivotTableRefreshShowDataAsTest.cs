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
						field.SubtotalTop = true;
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
						field.SubtotalTop = true;
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
						field.SubtotalTop = true;
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
						field.SubtotalTop = true;
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
						field.SubtotalTop = true;
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
						field.SubtotalTop = true;
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
						field.SubtotalTop = true;
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
						field.SubtotalTop = true;
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
						field.SubtotalTop = true;
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
						field.SubtotalTop = true;
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
						field.SubtotalTop = true;
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
						field.SubtotalTop = true;
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
						field.SubtotalTop = true;
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
						field.SubtotalTop = false;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B46:F73"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
					package.SaveAs(@"C:\Users\ems\Downloads\OUT.xlsx");
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
						field.SubtotalTop = true;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B18:K22"), pivotTable.Address);
					Assert.AreEqual(7, pivotTable.Fields.Count);
					package.SaveAs(newFile.File);
				}

				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
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
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableShowDataAs.xlsx")]
		public void PivotTableRefreshShowDataAsPercentOfParentRowWithColumnFieldsAndRowDataFieldsSubtotalsOn()
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
						field.SubtotalTop = false;
					}
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					package.SaveAs(@"C:\Users\ems\Downloads\OUT.xlsx");
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
		#endregion

		#region PercentOfParentColumn Tests
		/*
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
						field.SubtotalTop = true;
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
		*/
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
