using System.IO;
using System.Linq;
using EPPlusTest.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest.Table.PivotTable.PivotTableRefresh
{
	[TestClass]
	public class PivotTableRefreshCalculatedFieldsTest
	{
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
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.First();
					var formulaCacheField = cacheDefinition.CacheFields.First(c => c.Name == "CalculatedField");
					formulaCacheField.Formula = "'Wholesale Price'";
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B2:D13"), pivotTable.Address);
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
					package.Workbook.PivotCacheDefinitions.First().UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B2:D13"), pivotTable.Address);
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
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.First();
					var formulaCacheField = cacheDefinition.CacheFields.First(c => c.Name == "CalculatedField");
					formulaCacheField.Formula = "'Wholesale Price' * Item";
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B2:D13"), pivotTable.Address);
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
					package.Workbook.PivotCacheDefinitions.First().UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B18:D27"), pivotTable.Address);
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
					new ExpectedCellValue(sheetName, 22, 2, "Tent"),
					new ExpectedCellValue(sheetName, 22, 3, 1194d),
					new ExpectedCellValue(sheetName, 22, 4, 6),
					new ExpectedCellValue(sheetName, 23, 2, "Sleeping Bag"),
					new ExpectedCellValue(sheetName, 23, 3, 99d),
					new ExpectedCellValue(sheetName, 23, 4, 1),
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
					package.Workbook.PivotCacheDefinitions.First().UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("G4:L14"), pivotTable.Address);
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


		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithCalculatedFields.xlsx")]
		public void PivotTableRefreshColumnRowAndPageFieldsWithCalculatedFieldsNamedSimilarly()
		{
			var file = new FileInfo("PivotTableWithCalculatedFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "Sheet2";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable4"];
					package.Workbook.PivotCacheDefinitions.First().UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B31:C42"), pivotTable.Address);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 31, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 32, 2, "January"),
					new ExpectedCellValue(sheetName, 33, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 34, 2, "Nashville"),
					new ExpectedCellValue(sheetName, 35, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 36, 2, "February"),
					new ExpectedCellValue(sheetName, 37, 2, "Nashville"),
					new ExpectedCellValue(sheetName, 38, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 39, 2, "March"),
					new ExpectedCellValue(sheetName, 40, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 41, 2, "Nashville"),
					new ExpectedCellValue(sheetName, 42, 2, "Grand Total"),

					new ExpectedCellValue(sheetName, 31, 3, "Sum of CalculatedField3"),
					new ExpectedCellValue(sheetName, 32, 3, .93),
					new ExpectedCellValue(sheetName, 33, 3, 1.1),
					new ExpectedCellValue(sheetName, 34, 3, .85),
					new ExpectedCellValue(sheetName, 35, 3, .83),
					new ExpectedCellValue(sheetName, 36, 3, .81),
					new ExpectedCellValue(sheetName, 37, 3, .8),
					new ExpectedCellValue(sheetName, 38, 3, .99),
					new ExpectedCellValue(sheetName, 39, 3, 1.04),
					new ExpectedCellValue(sheetName, 40, 3, 1.25),
					new ExpectedCellValue(sheetName, 41, 3, 1.04),
					new ExpectedCellValue(sheetName, 42, 3, .91)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithCalculatedFields.xlsx")]
		public void PivotTableRefreshColumnRowAndPageFieldsWithCalculatedFieldWithError()
		{
			var file = new FileInfo("PivotTableWithCalculatedFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "Sheet2";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable4"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.First();
					var cacheField = cacheDefinition.CacheFields.First(c => c.Name == "CalculatedField3");
					cacheField.Formula = "=1/0";
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B31:C42"), pivotTable.Address);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 31, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 32, 2, "January"),
					new ExpectedCellValue(sheetName, 33, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 34, 2, "Nashville"),
					new ExpectedCellValue(sheetName, 35, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 36, 2, "February"),
					new ExpectedCellValue(sheetName, 37, 2, "Nashville"),
					new ExpectedCellValue(sheetName, 38, 2, "San Francisco"),
					new ExpectedCellValue(sheetName, 39, 2, "March"),
					new ExpectedCellValue(sheetName, 40, 2, "Chicago"),
					new ExpectedCellValue(sheetName, 41, 2, "Nashville"),
					new ExpectedCellValue(sheetName, 42, 2, "Grand Total"),

					new ExpectedCellValue(sheetName, 31, 3, "Sum of CalculatedField3"),
					new ExpectedCellValue(sheetName, 32, 3, ExcelErrorValue.Create(eErrorType.Div0)),
					new ExpectedCellValue(sheetName, 33, 3, ExcelErrorValue.Create(eErrorType.Div0)),
					new ExpectedCellValue(sheetName, 34, 3, ExcelErrorValue.Create(eErrorType.Div0)),
					new ExpectedCellValue(sheetName, 35, 3, ExcelErrorValue.Create(eErrorType.Div0)),
					new ExpectedCellValue(sheetName, 36, 3, ExcelErrorValue.Create(eErrorType.Div0)),
					new ExpectedCellValue(sheetName, 37, 3, ExcelErrorValue.Create(eErrorType.Div0)),
					new ExpectedCellValue(sheetName, 38, 3, ExcelErrorValue.Create(eErrorType.Div0)),
					new ExpectedCellValue(sheetName, 39, 3, ExcelErrorValue.Create(eErrorType.Div0)),
					new ExpectedCellValue(sheetName, 40, 3, ExcelErrorValue.Create(eErrorType.Div0)),
					new ExpectedCellValue(sheetName, 41, 3, ExcelErrorValue.Create(eErrorType.Div0)),
					new ExpectedCellValue(sheetName, 42, 3, ExcelErrorValue.Create(eErrorType.Div0))
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableCalculatedFieldError.xlsx")]
		public void PivotTableRefreshCalculatedFieldNamesQuoted()
		{
			var file = new FileInfo("PivotTableCalculatedFieldError.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "Sheet2";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable1"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.First();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("C3:F19"), pivotTable.Address);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 3, 3, "Row Labels"),
					new ExpectedCellValue(sheetName, 3, 4, "Sales"),
					new ExpectedCellValue(sheetName, 3, 5, "Profit"),
					new ExpectedCellValue(sheetName, 3, 6, "Sum of Profit %"),
					new ExpectedCellValue(sheetName, 4, 3, "Antarcticopy"),
					new ExpectedCellValue(sheetName, 4, 4, 3980.84),
					new ExpectedCellValue(sheetName, 4, 5, 1119.44),
					new ExpectedCellValue(sheetName, 4, 6, 0.281206981441103),
					new ExpectedCellValue(sheetName, 5, 3, "Autohaus Mielberg KG"),
					new ExpectedCellValue(sheetName, 5, 4, 433.61),
					new ExpectedCellValue(sheetName, 5, 5, 95.41),
					new ExpectedCellValue(sheetName, 5, 6, 0.220036438274025),
					new ExpectedCellValue(sheetName, 6, 3, "BYT-KOMPLET s.r.o."),
					new ExpectedCellValue(sheetName, 6, 4, 2471.3),
					new ExpectedCellValue(sheetName, 6, 5, 546),
					new ExpectedCellValue(sheetName, 6, 6, 0.220936349289847),
					new ExpectedCellValue(sheetName, 7, 3, "Deerfield Graphics Company"),
					new ExpectedCellValue(sheetName, 7, 4, 1638.1),
					new ExpectedCellValue(sheetName, 7, 5, 1638.1),
					new ExpectedCellValue(sheetName, 7, 6, 1),
					new ExpectedCellValue(sheetName, 8, 3, "Designstudio Gmunden"),
					new ExpectedCellValue(sheetName, 8, 4, 3849.7),
					new ExpectedCellValue(sheetName, 8, 5, 847.2),
					new ExpectedCellValue(sheetName, 8, 6, 0.220069096293218),
					new ExpectedCellValue(sheetName, 9, 3, "Englunds Kontorsmöbler AB"),
					new ExpectedCellValue(sheetName, 9, 4, 1038.27),
					new ExpectedCellValue(sheetName, 9, 5, 334.27),
					new ExpectedCellValue(sheetName, 9, 6, 0.321949011336165),
					new ExpectedCellValue(sheetName, 10, 3, "Gagn & Gaman"),
					new ExpectedCellValue(sheetName, 10, 4, 1352.07),
					new ExpectedCellValue(sheetName, 10, 5, 259.97),
					new ExpectedCellValue(sheetName, 10, 6, 0.192275547863646),
					new ExpectedCellValue(sheetName, 11, 3, "Guildford Water Department"),
					new ExpectedCellValue(sheetName, 11, 4, 822),
					new ExpectedCellValue(sheetName, 11, 5, 822),
					new ExpectedCellValue(sheetName, 11, 6, 1),
					new ExpectedCellValue(sheetName, 12, 3, "Heimilisprydi"),
					new ExpectedCellValue(sheetName, 12, 4, 3119.57),
					new ExpectedCellValue(sheetName, 12, 5, 521.17),
					new ExpectedCellValue(sheetName, 12, 6, 0.167064691608138),
					new ExpectedCellValue(sheetName, 13, 3, "John Haddock Insurance Co."),
					new ExpectedCellValue(sheetName, 13, 4, 9444.3),
					new ExpectedCellValue(sheetName, 13, 5, 4444.8),
					new ExpectedCellValue(sheetName, 13, 6, 0.470633080270639),
					new ExpectedCellValue(sheetName, 14, 3, "Klubben"),
					new ExpectedCellValue(sheetName, 14, 4, 18142),
					new ExpectedCellValue(sheetName, 14, 5, 6349.7),
					new ExpectedCellValue(sheetName, 14, 6, 0.35),
					new ExpectedCellValue(sheetName, 15, 3, "Progressive Home Furnishings"),
					new ExpectedCellValue(sheetName, 15, 4, 2461),
					new ExpectedCellValue(sheetName, 15, 5, 621.6),
					new ExpectedCellValue(sheetName, 15, 6, 0.25258025193011),
					new ExpectedCellValue(sheetName, 16, 3, "Selangorian Ltd."),
					new ExpectedCellValue(sheetName, 16, 4, 10007.97),
					new ExpectedCellValue(sheetName, 16, 5, 3804.07),
					new ExpectedCellValue(sheetName, 16, 6, 0.380104057066518),
					new ExpectedCellValue(sheetName, 17, 3, "The Cannon Group PLC"),
					new ExpectedCellValue(sheetName, 17, 4, 26324.08),
					new ExpectedCellValue(sheetName, 17, 5, 8148.48),
					new ExpectedCellValue(sheetName, 17, 6, 0.309544721031086),
					new ExpectedCellValue(sheetName, 18, 3, "Total"),
					new ExpectedCellValue(sheetName, 18, 4, 85084.81),
					new ExpectedCellValue(sheetName, 18, 5, 29552.21),
					new ExpectedCellValue(sheetName, 18, 6, 0.34732650869174),
					new ExpectedCellValue(sheetName, 19, 3, "Grand Total"),
					new ExpectedCellValue(sheetName, 19, 4, 170169.62),
					new ExpectedCellValue(sheetName, 19, 5, 59104.42),
					new ExpectedCellValue(sheetName, 19, 6, 0.347326508691739),
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableCalculatedFieldError.xlsx")]
		public void PivotTableRefreshCalculatedFieldWithPercentage()
		{
			var file = new FileInfo("PivotTableCalculatedFieldError.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "Sheet2";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable1"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.First();
					cacheDefinition.CacheFields.First(c => c.Name == "Profit %").Formula = "'Profit (LCY)' * 3%";
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("C3:F19"), pivotTable.Address);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 3, 3, "Row Labels"),
					new ExpectedCellValue(sheetName, 3, 4, "Sales"),
					new ExpectedCellValue(sheetName, 3, 5, "Profit"),
					new ExpectedCellValue(sheetName, 3, 6, "Sum of Profit %"),
					new ExpectedCellValue(sheetName, 4, 3, "Antarcticopy"),
					new ExpectedCellValue(sheetName, 4, 4, 3980.84),
					new ExpectedCellValue(sheetName, 4, 5, 1119.44),
					new ExpectedCellValue(sheetName, 4, 6, 33.5832),
					new ExpectedCellValue(sheetName, 5, 3, "Autohaus Mielberg KG"),
					new ExpectedCellValue(sheetName, 5, 4, 433.61),
					new ExpectedCellValue(sheetName, 5, 5, 95.41),
					new ExpectedCellValue(sheetName, 5, 6, 2.8623),
					new ExpectedCellValue(sheetName, 6, 3, "BYT-KOMPLET s.r.o."),
					new ExpectedCellValue(sheetName, 6, 4, 2471.3),
					new ExpectedCellValue(sheetName, 6, 5, 546),
					new ExpectedCellValue(sheetName, 6, 6, 16.38),
					new ExpectedCellValue(sheetName, 7, 3, "Deerfield Graphics Company"),
					new ExpectedCellValue(sheetName, 7, 4, 1638.1),
					new ExpectedCellValue(sheetName, 7, 5, 1638.1),
					new ExpectedCellValue(sheetName, 7, 6, 49.143),
					new ExpectedCellValue(sheetName, 8, 3, "Designstudio Gmunden"),
					new ExpectedCellValue(sheetName, 8, 4, 3849.7),
					new ExpectedCellValue(sheetName, 8, 5, 847.2),
					new ExpectedCellValue(sheetName, 8, 6, 25.416),
					new ExpectedCellValue(sheetName, 9, 3, "Englunds Kontorsmöbler AB"),
					new ExpectedCellValue(sheetName, 9, 4, 1038.27),
					new ExpectedCellValue(sheetName, 9, 5, 334.27),
					new ExpectedCellValue(sheetName, 9, 6, 10.0281),
					new ExpectedCellValue(sheetName, 10, 3, "Gagn & Gaman"),
					new ExpectedCellValue(sheetName, 10, 4, 1352.07),
					new ExpectedCellValue(sheetName, 10, 5, 259.97),
					new ExpectedCellValue(sheetName, 10, 6, 7.7991),
					new ExpectedCellValue(sheetName, 11, 3, "Guildford Water Department"),
					new ExpectedCellValue(sheetName, 11, 4, 822),
					new ExpectedCellValue(sheetName, 11, 5, 822),
					new ExpectedCellValue(sheetName, 11, 6, 24.66),
					new ExpectedCellValue(sheetName, 12, 3, "Heimilisprydi"),
					new ExpectedCellValue(sheetName, 12, 4, 3119.57),
					new ExpectedCellValue(sheetName, 12, 5, 521.17),
					new ExpectedCellValue(sheetName, 12, 6, 15.6351),
					new ExpectedCellValue(sheetName, 13, 3, "John Haddock Insurance Co."),
					new ExpectedCellValue(sheetName, 13, 4, 9444.3),
					new ExpectedCellValue(sheetName, 13, 5, 4444.8),
					new ExpectedCellValue(sheetName, 13, 6, 133.344),
					new ExpectedCellValue(sheetName, 14, 3, "Klubben"),
					new ExpectedCellValue(sheetName, 14, 4, 18142),
					new ExpectedCellValue(sheetName, 14, 5, 6349.7),
					new ExpectedCellValue(sheetName, 14, 6, 190.491),
					new ExpectedCellValue(sheetName, 15, 3, "Progressive Home Furnishings"),
					new ExpectedCellValue(sheetName, 15, 4, 2461),
					new ExpectedCellValue(sheetName, 15, 5, 621.6),
					new ExpectedCellValue(sheetName, 15, 6, 18.648),
					new ExpectedCellValue(sheetName, 16, 3, "Selangorian Ltd."),
					new ExpectedCellValue(sheetName, 16, 4, 10007.97),
					new ExpectedCellValue(sheetName, 16, 5, 3804.07),
					new ExpectedCellValue(sheetName, 16, 6, 114.1221),
					new ExpectedCellValue(sheetName, 17, 3, "The Cannon Group PLC"),
					new ExpectedCellValue(sheetName, 17, 4, 26324.08),
					new ExpectedCellValue(sheetName, 17, 5, 8148.48),
					new ExpectedCellValue(sheetName, 17, 6, 244.4544),
					new ExpectedCellValue(sheetName, 18, 3, "Total"),
					new ExpectedCellValue(sheetName, 18, 4, 85084.81),
					new ExpectedCellValue(sheetName, 18, 5, 29552.21),
					new ExpectedCellValue(sheetName, 18, 6, 886.5663),
					new ExpectedCellValue(sheetName, 19, 3, "Grand Total"),
					new ExpectedCellValue(sheetName, 19, 4, 170169.62),
					new ExpectedCellValue(sheetName, 19, 5, 59104.42),
					new ExpectedCellValue(sheetName, 19, 6, 1773.1326),
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\ShowDataAsComplex.xlsx")]
		public void PivotTableRefreshCalculatedFieldDateComparison()
		{
			var file = new FileInfo("ShowDataAsComplex.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "Sheet1";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable3"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.First();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("B37:C46"), pivotTable.Address);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 37, 2, "Row Labels"),
					new ExpectedCellValue(sheetName, 37, 3, "Sum of DateCalculatedField"),
					new ExpectedCellValue(sheetName, 38, 2, "Accounts Receivable"),
					new ExpectedCellValue(sheetName, 38, 3, 7),
					new ExpectedCellValue(sheetName, 39, 2, "Accounts Receivable, Total"),
					new ExpectedCellValue(sheetName, 39, 3, 7),
					new ExpectedCellValue(sheetName, 40, 2, "ASSETS"),
					new ExpectedCellValue(sheetName, 40, 3, 7),
					new ExpectedCellValue(sheetName, 41, 2, "Cash"),
					new ExpectedCellValue(sheetName, 41, 3, 7),
					new ExpectedCellValue(sheetName, 42, 2, "Customers, EU"),
					new ExpectedCellValue(sheetName, 42, 3, 7),
					new ExpectedCellValue(sheetName, 43, 2, "Customers, North America"),
					new ExpectedCellValue(sheetName, 43, 3, 7),
					new ExpectedCellValue(sheetName, 44, 2, "Liquid Assets, Total"),
					new ExpectedCellValue(sheetName, 44, 3, 7),
					new ExpectedCellValue(sheetName, 45, 2, "Securities, Total"),
					new ExpectedCellValue(sheetName, 45, 3, 7),
					new ExpectedCellValue(sheetName, 46, 2, "Grand Total"),
					new ExpectedCellValue(sheetName, 46, 3, 7),
				});
			}
		}
		#endregion
	}
}
