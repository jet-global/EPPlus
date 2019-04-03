using System.IO;
using System.Linq;
using EPPlusTest.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Table.PivotTable;

namespace EPPlusTest.Table.PivotTable.PivotTableRefresh
{
	[TestClass]
	public class PivotTableSummarizeValuesByPercentOfParentRowTest
	{
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\SummarizeValuesByTestWorkbook.xlsx")]
		public void PivotTableRefreshSummarizeValuesBySumShowDataAsPercentOfParentRow()
		{
			var file = new FileInfo("SummarizeValuesByTestWorkbook.xlsx");
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
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("A1:F12"), pivotTable.Address);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 1, 1, "Sum of Balance at End Date"),
					new ExpectedCellValue(sheetName, 2, 1, "Quarters"),
					new ExpectedCellValue(sheetName, 3, 1, "Qtr1"),
					new ExpectedCellValue(sheetName, 4, 1, null),
					new ExpectedCellValue(sheetName, 5, 1, "Qtr2"),
					new ExpectedCellValue(sheetName, 6, 1, null),
					new ExpectedCellValue(sheetName, 7, 1, "Qtr3"),
					new ExpectedCellValue(sheetName, 8, 1, null),
					new ExpectedCellValue(sheetName, 9, 1, "Qtr4"),
					new ExpectedCellValue(sheetName, 10, 1, null),
					new ExpectedCellValue(sheetName, 11, 1, null),
					new ExpectedCellValue(sheetName, 12, 1, "Grand Total"),
					new ExpectedCellValue(sheetName, 1, 2, null),
					new ExpectedCellValue(sheetName, 2, 2, "Debit/Credit"),
					new ExpectedCellValue(sheetName, 3, 2, "Both"),
					new ExpectedCellValue(sheetName, 4, 2, "Debit"),
					new ExpectedCellValue(sheetName, 5, 2, "Both"),
					new ExpectedCellValue(sheetName, 6, 2, "Debit"),
					new ExpectedCellValue(sheetName, 7, 2, "Both"),
					new ExpectedCellValue(sheetName, 8, 2, "Credit"),
					new ExpectedCellValue(sheetName, 9, 2, "Both"),
					new ExpectedCellValue(sheetName, 10, 2, "Credit"),
					new ExpectedCellValue(sheetName, 11, 2, "Debit"),
					new ExpectedCellValue(sheetName, 12, 2, null),
					new ExpectedCellValue(sheetName, 1, 3, "Account Type"),
					new ExpectedCellValue(sheetName, 2, 3, "Begin-Total"),
					new ExpectedCellValue(sheetName, 3, 3, null),
					new ExpectedCellValue(sheetName, 4, 3, null),
					new ExpectedCellValue(sheetName, 5, 3, null),
					new ExpectedCellValue(sheetName, 6, 3, null),
					new ExpectedCellValue(sheetName, 7, 3, null),
					new ExpectedCellValue(sheetName, 8, 3, null),
					new ExpectedCellValue(sheetName, 9, 3, null),
					new ExpectedCellValue(sheetName, 10, 3, null),
					new ExpectedCellValue(sheetName, 11, 3, null),
					new ExpectedCellValue(sheetName, 12, 3, null),
					new ExpectedCellValue(sheetName, 1, 4, null),
					new ExpectedCellValue(sheetName, 2, 4, "End-Total"),
					new ExpectedCellValue(sheetName, 3, 4, null),
					new ExpectedCellValue(sheetName, 4, 4, null),
					new ExpectedCellValue(sheetName, 5, 4, -6.3252),
					new ExpectedCellValue(sheetName, 6, 4, 7.3252),
					new ExpectedCellValue(sheetName, 7, 4, 1d),
					new ExpectedCellValue(sheetName, 8, 4, 0),
					new ExpectedCellValue(sheetName, 9, 4, 0.8047),
					new ExpectedCellValue(sheetName, 10, 4, .1953),
					new ExpectedCellValue(sheetName, 11, 4, 0),
					new ExpectedCellValue(sheetName, 12, 4, 1d),
					new ExpectedCellValue(sheetName, 1, 5, null),
					new ExpectedCellValue(sheetName, 2, 5, "Posting"),
					new ExpectedCellValue(sheetName, 3, 5, 0),
					new ExpectedCellValue(sheetName, 4, 5, 1d),
					new ExpectedCellValue(sheetName, 5, 5, 0),
					new ExpectedCellValue(sheetName, 6, 5, 1d),
					new ExpectedCellValue(sheetName, 7, 5, null),
					new ExpectedCellValue(sheetName, 8, 5, null),
					new ExpectedCellValue(sheetName, 9, 5, -1.4201),
					new ExpectedCellValue(sheetName, 10, 5, 0),
					new ExpectedCellValue(sheetName, 11, 5, 2.4201),
					new ExpectedCellValue(sheetName, 12, 5, 1d),
					new ExpectedCellValue(sheetName, 1, 6, null),
					new ExpectedCellValue(sheetName, 2, 6, "Grand Total"),
					new ExpectedCellValue(sheetName, 3, 6, 0),
					new ExpectedCellValue(sheetName, 4, 6, 1d),
					new ExpectedCellValue(sheetName, 5, 6, -1.3251),
					new ExpectedCellValue(sheetName, 6, 6, 2.3251),
					new ExpectedCellValue(sheetName, 7, 6, 1d),
					new ExpectedCellValue(sheetName, 8, 6, 0),
					new ExpectedCellValue(sheetName, 9, 6, 1.1587),
					new ExpectedCellValue(sheetName, 10, 6, 0.2263),
					new ExpectedCellValue(sheetName, 11, 6, -0.3851),
					new ExpectedCellValue(sheetName, 12, 6, 1d)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\SummarizeValuesByTestWorkbook.xlsx")]
		public void PivotTableRefreshSummarizeValuesByCountShowDataAsPercentOfParentRow()
		{
			var file = new FileInfo("SummarizeValuesByTestWorkbook.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable1"];

					// Show 'Wholesale Price' data as the percentage of its parent row.
					var balanceAtEndDateDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Balance at End Date");
					balanceAtEndDateDataField.Function = DataFieldFunctions.Count;

					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("A1:F12"), pivotTable.Address);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 1, 1, "Sum of Balance at End Date"),
					new ExpectedCellValue(sheetName, 2, 1, "Quarters"),
					new ExpectedCellValue(sheetName, 3, 1, "Qtr1"),
					new ExpectedCellValue(sheetName, 4, 1, null),
					new ExpectedCellValue(sheetName, 5, 1, "Qtr2"),
					new ExpectedCellValue(sheetName, 6, 1, null),
					new ExpectedCellValue(sheetName, 7, 1, "Qtr3"),
					new ExpectedCellValue(sheetName, 8, 1, null),
					new ExpectedCellValue(sheetName, 9, 1, "Qtr4"),
					new ExpectedCellValue(sheetName, 10, 1, null),
					new ExpectedCellValue(sheetName, 11, 1, null),
					new ExpectedCellValue(sheetName, 12, 1, "Grand Total"),
					new ExpectedCellValue(sheetName, 1, 2, null),
					new ExpectedCellValue(sheetName, 2, 2, "Debit/Credit"),
					new ExpectedCellValue(sheetName, 3, 2, "Both"),
					new ExpectedCellValue(sheetName, 4, 2, "Debit"),
					new ExpectedCellValue(sheetName, 5, 2, "Both"),
					new ExpectedCellValue(sheetName, 6, 2, "Debit"),
					new ExpectedCellValue(sheetName, 7, 2, "Both"),
					new ExpectedCellValue(sheetName, 8, 2, "Credit"),
					new ExpectedCellValue(sheetName, 9, 2, "Both"),
					new ExpectedCellValue(sheetName, 10, 2, "Credit"),
					new ExpectedCellValue(sheetName, 11, 2, "Debit"),
					new ExpectedCellValue(sheetName, 12, 2, null),
					new ExpectedCellValue(sheetName, 1, 3, "Account Type"),
					new ExpectedCellValue(sheetName, 2, 3, "Begin-Total"),
					new ExpectedCellValue(sheetName, 3, 3, 0.5),
					new ExpectedCellValue(sheetName, 4, 3, 0.5),
					new ExpectedCellValue(sheetName, 5, 3, 0),
					new ExpectedCellValue(sheetName, 6, 3, 1d),
					new ExpectedCellValue(sheetName, 7, 3, 0.5),
					new ExpectedCellValue(sheetName, 8, 3, 0.5),
					new ExpectedCellValue(sheetName, 9, 3, 0),
					new ExpectedCellValue(sheetName, 10, 3, 1d),
					new ExpectedCellValue(sheetName, 11, 3, 0),
					new ExpectedCellValue(sheetName, 12, 3, 1d),
					new ExpectedCellValue(sheetName, 1, 4, null),
					new ExpectedCellValue(sheetName, 2, 4, "End-Total"),
					new ExpectedCellValue(sheetName, 3, 4, 1d),
					new ExpectedCellValue(sheetName, 4, 4, 0),
					new ExpectedCellValue(sheetName, 5, 4, 0.5),
					new ExpectedCellValue(sheetName, 6, 4, 0.5),
					new ExpectedCellValue(sheetName, 7, 4, 1d),
					new ExpectedCellValue(sheetName, 8, 4, 0),
					new ExpectedCellValue(sheetName, 9, 4, 0.5),
					new ExpectedCellValue(sheetName, 10, 4, 0.5),
					new ExpectedCellValue(sheetName, 11, 4, 0),
					new ExpectedCellValue(sheetName, 12, 4, 1d),
					new ExpectedCellValue(sheetName, 1, 5, null),
					new ExpectedCellValue(sheetName, 2, 5, "Posting"),
					new ExpectedCellValue(sheetName, 3, 5, 0),
					new ExpectedCellValue(sheetName, 4, 5, 1d),
					new ExpectedCellValue(sheetName, 5, 5, 0),
					new ExpectedCellValue(sheetName, 6, 5, 1d),
					new ExpectedCellValue(sheetName, 7, 5, 1d),
					new ExpectedCellValue(sheetName, 8, 5, 0),
					new ExpectedCellValue(sheetName, 9, 5, 0.3750),
					new ExpectedCellValue(sheetName, 10, 5, 0.125),
					new ExpectedCellValue(sheetName, 11, 5, 0.5),
					new ExpectedCellValue(sheetName, 12, 5, 1d),
					new ExpectedCellValue(sheetName, 1, 6, null),
					new ExpectedCellValue(sheetName, 2, 6, "Grand Total"),
					new ExpectedCellValue(sheetName, 3, 6, 0.6),
					new ExpectedCellValue(sheetName, 4, 6, 0.4),
					new ExpectedCellValue(sheetName, 5, 6, 0.25),
					new ExpectedCellValue(sheetName, 6, 6, 0.75),
					new ExpectedCellValue(sheetName, 7, 6, 0.75),
					new ExpectedCellValue(sheetName, 8, 6, 0.25),
					new ExpectedCellValue(sheetName, 9, 6, 0.3636),
					new ExpectedCellValue(sheetName, 10, 6, 0.2727),
					new ExpectedCellValue(sheetName, 11, 6, 0.3636),
					new ExpectedCellValue(sheetName, 12, 6, 1d)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\SummarizeValuesByTestWorkbook.xlsx")]
		public void PivotTableRefreshSummarizeValuesByAverageShowDataAsPercentOfParentRow()
		{
			var file = new FileInfo("SummarizeValuesByTestWorkbook.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable1"];

					// Show 'Wholesale Price' data as the percentage of its parent row.
					var balanceAtEndDateDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Balance at End Date");
					balanceAtEndDateDataField.Function = DataFieldFunctions.Average;

					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("A1:F12"), pivotTable.Address);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 1, 1, "Sum of Balance at End Date"),
					new ExpectedCellValue(sheetName, 2, 1, "Quarters"),
					new ExpectedCellValue(sheetName, 3, 1, "Qtr1"),
					new ExpectedCellValue(sheetName, 4, 1, null),
					new ExpectedCellValue(sheetName, 5, 1, "Qtr2"),
					new ExpectedCellValue(sheetName, 6, 1, null),
					new ExpectedCellValue(sheetName, 7, 1, "Qtr3"),
					new ExpectedCellValue(sheetName, 8, 1, null),
					new ExpectedCellValue(sheetName, 9, 1, "Qtr4"),
					new ExpectedCellValue(sheetName, 10, 1, null),
					new ExpectedCellValue(sheetName, 11, 1, null),
					new ExpectedCellValue(sheetName, 12, 1, "Grand Total"),
					new ExpectedCellValue(sheetName, 1, 2, null),
					new ExpectedCellValue(sheetName, 2, 2, "Debit/Credit"),
					new ExpectedCellValue(sheetName, 3, 2, "Both"),
					new ExpectedCellValue(sheetName, 4, 2, "Debit"),
					new ExpectedCellValue(sheetName, 5, 2, "Both"),
					new ExpectedCellValue(sheetName, 6, 2, "Debit"),
					new ExpectedCellValue(sheetName, 7, 2, "Both"),
					new ExpectedCellValue(sheetName, 8, 2, "Credit"),
					new ExpectedCellValue(sheetName, 9, 2, "Both"),
					new ExpectedCellValue(sheetName, 10, 2, "Credit"),
					new ExpectedCellValue(sheetName, 11, 2, "Debit"),
					new ExpectedCellValue(sheetName, 12, 2, null),
					new ExpectedCellValue(sheetName, 1, 3, "Account Type"),
					new ExpectedCellValue(sheetName, 2, 3, "Begin-Total"),
					new ExpectedCellValue(sheetName, 3, 3, null),
					new ExpectedCellValue(sheetName, 4, 3, null),
					new ExpectedCellValue(sheetName, 5, 3, null),
					new ExpectedCellValue(sheetName, 6, 3, null),
					new ExpectedCellValue(sheetName, 7, 3, null),
					new ExpectedCellValue(sheetName, 8, 3, null),
					new ExpectedCellValue(sheetName, 9, 3, null),
					new ExpectedCellValue(sheetName, 10, 3, null),
					new ExpectedCellValue(sheetName, 11, 3, null),
					new ExpectedCellValue(sheetName, 12, 3, null),
					new ExpectedCellValue(sheetName, 1, 4, null),
					new ExpectedCellValue(sheetName, 2, 4, "End-Total"),
					new ExpectedCellValue(sheetName, 3, 4, null),
					new ExpectedCellValue(sheetName, 4, 4, null),
					new ExpectedCellValue(sheetName, 5, 4, -12.6504),
					new ExpectedCellValue(sheetName, 6, 4, 14.6504),
					new ExpectedCellValue(sheetName, 7, 4, 1d),
					new ExpectedCellValue(sheetName, 8, 4, 0),
					new ExpectedCellValue(sheetName, 9, 4, 1.6095),
					new ExpectedCellValue(sheetName, 10, 4, 0.3905),
					new ExpectedCellValue(sheetName, 11, 4, 0),
					new ExpectedCellValue(sheetName, 12, 4, 1d),
					new ExpectedCellValue(sheetName, 1, 5, null),
					new ExpectedCellValue(sheetName, 2, 5, "Posting"),
					new ExpectedCellValue(sheetName, 3, 5, 0),
					new ExpectedCellValue(sheetName, 4, 5, 1d),
					new ExpectedCellValue(sheetName, 5, 5, 0),
					new ExpectedCellValue(sheetName, 6, 5, 1d),
					new ExpectedCellValue(sheetName, 7, 5, null),
					new ExpectedCellValue(sheetName, 8, 5, null),
					new ExpectedCellValue(sheetName, 9, 5, -3.7869),
					new ExpectedCellValue(sheetName, 10, 5, 0),
					new ExpectedCellValue(sheetName, 11, 5, 4.8402),
					new ExpectedCellValue(sheetName, 12, 5, 1d),
					new ExpectedCellValue(sheetName, 1, 6, null),
					new ExpectedCellValue(sheetName, 2, 6, "Grand Total"),
					new ExpectedCellValue(sheetName, 3, 6, 0),
					new ExpectedCellValue(sheetName, 4, 6, 1.5),
					new ExpectedCellValue(sheetName, 5, 6, -5.3003),
					new ExpectedCellValue(sheetName, 6, 6, 3.1001),
					new ExpectedCellValue(sheetName, 7, 6, 1.5),
					new ExpectedCellValue(sheetName, 8, 6, 0),
					new ExpectedCellValue(sheetName, 9, 6, 3.1865),
					new ExpectedCellValue(sheetName, 10, 6, 0.8299),
					new ExpectedCellValue(sheetName, 11, 6, -1.0589),
					new ExpectedCellValue(sheetName, 12, 6, 1d)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\SummarizeValuesByTestWorkbook.xlsx")]
		public void PivotTableRefreshSummarizeValuesByMaxShowDataAsPercentOfParentRow()
		{
			var file = new FileInfo("SummarizeValuesByTestWorkbook.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable1"];

					// Show 'Wholesale Price' data as the percentage of its parent row.
					var balanceAtEndDateDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Balance at End Date");
					balanceAtEndDateDataField.Function = DataFieldFunctions.Max;

					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("A1:F12"), pivotTable.Address);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 1, 1, "Sum of Balance at End Date"),
					new ExpectedCellValue(sheetName, 2, 1, "Quarters"),
					new ExpectedCellValue(sheetName, 3, 1, "Qtr1"),
					new ExpectedCellValue(sheetName, 4, 1, null),
					new ExpectedCellValue(sheetName, 5, 1, "Qtr2"),
					new ExpectedCellValue(sheetName, 6, 1, null),
					new ExpectedCellValue(sheetName, 7, 1, "Qtr3"),
					new ExpectedCellValue(sheetName, 8, 1, null),
					new ExpectedCellValue(sheetName, 9, 1, "Qtr4"),
					new ExpectedCellValue(sheetName, 10, 1, null),
					new ExpectedCellValue(sheetName, 11, 1, null),
					new ExpectedCellValue(sheetName, 12, 1, "Grand Total"),
					new ExpectedCellValue(sheetName, 1, 2, null),
					new ExpectedCellValue(sheetName, 2, 2, "Debit/Credit"),
					new ExpectedCellValue(sheetName, 3, 2, "Both"),
					new ExpectedCellValue(sheetName, 4, 2, "Debit"),
					new ExpectedCellValue(sheetName, 5, 2, "Both"),
					new ExpectedCellValue(sheetName, 6, 2, "Debit"),
					new ExpectedCellValue(sheetName, 7, 2, "Both"),
					new ExpectedCellValue(sheetName, 8, 2, "Credit"),
					new ExpectedCellValue(sheetName, 9, 2, "Both"),
					new ExpectedCellValue(sheetName, 10, 2, "Credit"),
					new ExpectedCellValue(sheetName, 11, 2, "Debit"),
					new ExpectedCellValue(sheetName, 12, 2, null),
					new ExpectedCellValue(sheetName, 1, 3, "Account Type"),
					new ExpectedCellValue(sheetName, 2, 3, "Begin-Total"),
					new ExpectedCellValue(sheetName, 3, 3, null),
					new ExpectedCellValue(sheetName, 4, 3, null),
					new ExpectedCellValue(sheetName, 5, 3, null),
					new ExpectedCellValue(sheetName, 6, 3, null),
					new ExpectedCellValue(sheetName, 7, 3, null),
					new ExpectedCellValue(sheetName, 8, 3, null),
					new ExpectedCellValue(sheetName, 9, 3, null),
					new ExpectedCellValue(sheetName, 10, 3, null),
					new ExpectedCellValue(sheetName, 11, 3, null),
					new ExpectedCellValue(sheetName, 12, 3, null),
					new ExpectedCellValue(sheetName, 1, 4, null),
					new ExpectedCellValue(sheetName, 2, 4, "End-Total"),
					new ExpectedCellValue(sheetName, 3, 4, null),
					new ExpectedCellValue(sheetName, 4, 4, null),
					new ExpectedCellValue(sheetName, 5, 4, -.8635),
					new ExpectedCellValue(sheetName, 6, 4, 1d),
					new ExpectedCellValue(sheetName, 7, 4, 1d),
					new ExpectedCellValue(sheetName, 8, 4, 0),
					new ExpectedCellValue(sheetName, 9, 4, 1d),
					new ExpectedCellValue(sheetName, 10, 4, 0.2426),
					new ExpectedCellValue(sheetName, 11, 4, 0),
					new ExpectedCellValue(sheetName, 12, 4, 1d),
					new ExpectedCellValue(sheetName, 1, 5, null),
					new ExpectedCellValue(sheetName, 2, 5, "Posting"),
					new ExpectedCellValue(sheetName, 3, 5, 0),
					new ExpectedCellValue(sheetName, 4, 5, 1d),
					new ExpectedCellValue(sheetName, 5, 5, 0),
					new ExpectedCellValue(sheetName, 6, 5, 1d),
					new ExpectedCellValue(sheetName, 7, 5, null),
					new ExpectedCellValue(sheetName, 8, 5, null),
					new ExpectedCellValue(sheetName, 9, 5, 1d),
					new ExpectedCellValue(sheetName, 10, 5, 0),
					new ExpectedCellValue(sheetName, 11, 5, 0.95),
					new ExpectedCellValue(sheetName, 12, 5, 1d),
					new ExpectedCellValue(sheetName, 1, 6, null),
					new ExpectedCellValue(sheetName, 2, 6, "Grand Total"),
					new ExpectedCellValue(sheetName, 3, 6, null),
					new ExpectedCellValue(sheetName, 4, 6, null),
					new ExpectedCellValue(sheetName, 5, 6, -0.8635),
					new ExpectedCellValue(sheetName, 6, 6, 1d),
					new ExpectedCellValue(sheetName, 7, 6, 1d),
					new ExpectedCellValue(sheetName, 8, 6, 0),
					new ExpectedCellValue(sheetName, 9, 6, 1d),
					new ExpectedCellValue(sheetName, 10, 6, 0.2426),
					new ExpectedCellValue(sheetName, 11, 6, 0.1319),
					new ExpectedCellValue(sheetName, 12, 6, 1d)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\SummarizeValuesByTestWorkbook.xlsx")]
		public void PivotTableRefreshSummarizeValuesByMinShowDataAsPercentOfParentRow()
		{
			var file = new FileInfo("SummarizeValuesByTestWorkbook.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable1"];

					// Show 'Wholesale Price' data as the percentage of its parent row.
					var balanceAtEndDateDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Balance at End Date");
					balanceAtEndDateDataField.Function = DataFieldFunctions.Min;

					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("A1:F12"), pivotTable.Address);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 1, 1, "Sum of Balance at End Date"),
					new ExpectedCellValue(sheetName, 2, 1, "Quarters"),
					new ExpectedCellValue(sheetName, 3, 1, "Qtr1"),
					new ExpectedCellValue(sheetName, 4, 1, null),
					new ExpectedCellValue(sheetName, 5, 1, "Qtr2"),
					new ExpectedCellValue(sheetName, 6, 1, null),
					new ExpectedCellValue(sheetName, 7, 1, "Qtr3"),
					new ExpectedCellValue(sheetName, 8, 1, null),
					new ExpectedCellValue(sheetName, 9, 1, "Qtr4"),
					new ExpectedCellValue(sheetName, 10, 1, null),
					new ExpectedCellValue(sheetName, 11, 1, null),
					new ExpectedCellValue(sheetName, 12, 1, "Grand Total"),
					new ExpectedCellValue(sheetName, 1, 2, null),
					new ExpectedCellValue(sheetName, 2, 2, "Debit/Credit"),
					new ExpectedCellValue(sheetName, 3, 2, "Both"),
					new ExpectedCellValue(sheetName, 4, 2, "Debit"),
					new ExpectedCellValue(sheetName, 5, 2, "Both"),
					new ExpectedCellValue(sheetName, 6, 2, "Debit"),
					new ExpectedCellValue(sheetName, 7, 2, "Both"),
					new ExpectedCellValue(sheetName, 8, 2, "Credit"),
					new ExpectedCellValue(sheetName, 9, 2, "Both"),
					new ExpectedCellValue(sheetName, 10, 2, "Credit"),
					new ExpectedCellValue(sheetName, 11, 2, "Debit"),
					new ExpectedCellValue(sheetName, 12, 2, null),
					new ExpectedCellValue(sheetName, 1, 3, "Account Type"),
					new ExpectedCellValue(sheetName, 2, 3, "Begin-Total"),
					new ExpectedCellValue(sheetName, 3, 3, null),
					new ExpectedCellValue(sheetName, 4, 3, null),
					new ExpectedCellValue(sheetName, 5, 3, null),
					new ExpectedCellValue(sheetName, 6, 3, null),
					new ExpectedCellValue(sheetName, 7, 3, null),
					new ExpectedCellValue(sheetName, 8, 3, null),
					new ExpectedCellValue(sheetName, 9, 3, null),
					new ExpectedCellValue(sheetName, 10, 3, null),
					new ExpectedCellValue(sheetName, 11, 3, null),
					new ExpectedCellValue(sheetName, 12, 3, null),
					new ExpectedCellValue(sheetName, 1, 4, null),
					new ExpectedCellValue(sheetName, 2, 4, "End-Total"),
					new ExpectedCellValue(sheetName, 3, 4, null),
					new ExpectedCellValue(sheetName, 4, 4, null),
					new ExpectedCellValue(sheetName, 5, 4, 1d),
					new ExpectedCellValue(sheetName, 6, 4, -1.1581),
					new ExpectedCellValue(sheetName, 7, 4, 1d),
					new ExpectedCellValue(sheetName, 8, 4, 0),
					new ExpectedCellValue(sheetName, 9, 4, 4.1214),
					new ExpectedCellValue(sheetName, 10, 4, 1d),
					new ExpectedCellValue(sheetName, 11, 4, 0),
					new ExpectedCellValue(sheetName, 12, 4, 1d),
					new ExpectedCellValue(sheetName, 1, 5, null),
					new ExpectedCellValue(sheetName, 2, 5, "Posting"),
					new ExpectedCellValue(sheetName, 3, 5, 0),
					new ExpectedCellValue(sheetName, 4, 5, 1d),
					new ExpectedCellValue(sheetName, 5, 5, 0),
					new ExpectedCellValue(sheetName, 6, 5, 1d),
					new ExpectedCellValue(sheetName, 7, 5, null),
					new ExpectedCellValue(sheetName, 8, 5, null),
					new ExpectedCellValue(sheetName, 9, 5, 0),
					new ExpectedCellValue(sheetName, 10, 5, 0),
					new ExpectedCellValue(sheetName, 11, 5, 1d),
					new ExpectedCellValue(sheetName, 12, 5, 1d),
					new ExpectedCellValue(sheetName, 1, 6, null),
					new ExpectedCellValue(sheetName, 2, 6, "Grand Total"),
					new ExpectedCellValue(sheetName, 3, 6, 0),
					new ExpectedCellValue(sheetName, 4, 6, 1d),
					new ExpectedCellValue(sheetName, 5, 6, 1d),
					new ExpectedCellValue(sheetName, 6, 6, 0),
					new ExpectedCellValue(sheetName, 7, 6, null),
					new ExpectedCellValue(sheetName, 8, 6, null),
					new ExpectedCellValue(sheetName, 9, 6, 0),
					new ExpectedCellValue(sheetName, 10, 6, 0),
					new ExpectedCellValue(sheetName, 11, 6, 1d),
					new ExpectedCellValue(sheetName, 12, 6, 1d)
				});
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\SummarizeValuesByTestWorkbook.xlsx")]
		public void PivotTableRefreshSummarizeValuesByProductShowDataAsPercentOfParentRow()
		{
			var file = new FileInfo("SummarizeValuesByTestWorkbook.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "PivotTables";
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					var pivotTable = worksheet.PivotTables["PivotTable1"];

					// Show 'Wholesale Price' data as the percentage of its parent row.
					var balanceAtEndDateDataField = pivotTable.DataFields.First(f => f.Name == "Sum of Balance at End Date");
					balanceAtEndDateDataField.Function = DataFieldFunctions.Product;

					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					ExcelPivotTableTest.CheckPivotTableAddress(new ExcelAddress("A1:F12"), pivotTable.Address);
					package.SaveAs(newFile.File);
				}
				TestHelperUtility.ValidateWorksheet(newFile.File, sheetName, new[]
				{
					new ExpectedCellValue(sheetName, 1, 1, "Sum of Balance at End Date"),
					new ExpectedCellValue(sheetName, 2, 1, "Quarters"),
					new ExpectedCellValue(sheetName, 3, 1, "Qtr1"),
					new ExpectedCellValue(sheetName, 4, 1, null),
					new ExpectedCellValue(sheetName, 5, 1, "Qtr2"),
					new ExpectedCellValue(sheetName, 6, 1, null),
					new ExpectedCellValue(sheetName, 7, 1, "Qtr3"),
					new ExpectedCellValue(sheetName, 8, 1, null),
					new ExpectedCellValue(sheetName, 9, 1, "Qtr4"),
					new ExpectedCellValue(sheetName, 10, 1, null),
					new ExpectedCellValue(sheetName, 11, 1, null),
					new ExpectedCellValue(sheetName, 12, 1, "Grand Total"),
					new ExpectedCellValue(sheetName, 1, 2, null),
					new ExpectedCellValue(sheetName, 2, 2, "Debit/Credit"),
					new ExpectedCellValue(sheetName, 3, 2, "Both"),
					new ExpectedCellValue(sheetName, 4, 2, "Debit"),
					new ExpectedCellValue(sheetName, 5, 2, "Both"),
					new ExpectedCellValue(sheetName, 6, 2, "Debit"),
					new ExpectedCellValue(sheetName, 7, 2, "Both"),
					new ExpectedCellValue(sheetName, 8, 2, "Credit"),
					new ExpectedCellValue(sheetName, 9, 2, "Both"),
					new ExpectedCellValue(sheetName, 10, 2, "Credit"),
					new ExpectedCellValue(sheetName, 11, 2, "Debit"),
					new ExpectedCellValue(sheetName, 12, 2, null),
					new ExpectedCellValue(sheetName, 1, 3, "Account Type"),
					new ExpectedCellValue(sheetName, 2, 3, "Begin-Total"),
					new ExpectedCellValue(sheetName, 3, 3, null),
					new ExpectedCellValue(sheetName, 4, 3, null),
					new ExpectedCellValue(sheetName, 5, 3, null),
					new ExpectedCellValue(sheetName, 6, 3, null),
					new ExpectedCellValue(sheetName, 7, 3, null),
					new ExpectedCellValue(sheetName, 8, 3, null),
					new ExpectedCellValue(sheetName, 9, 3, null),
					new ExpectedCellValue(sheetName, 10, 3, null),
					new ExpectedCellValue(sheetName, 11, 3, null),
					new ExpectedCellValue(sheetName, 12, 3, null),
					new ExpectedCellValue(sheetName, 1, 4, null),
					new ExpectedCellValue(sheetName, 2, 4, "End-Total"),
					new ExpectedCellValue(sheetName, 3, 4, null),
					new ExpectedCellValue(sheetName, 4, 4, null),
					new ExpectedCellValue(sheetName, 5, 4, 0),
					new ExpectedCellValue(sheetName, 6, 4, 0),
					new ExpectedCellValue(sheetName, 7, 4, 1d),
					new ExpectedCellValue(sheetName, 8, 4, 0),
					new ExpectedCellValue(sheetName, 9, 4, 0),
					new ExpectedCellValue(sheetName, 10, 4, 0),
					new ExpectedCellValue(sheetName, 11, 4, 0),
					new ExpectedCellValue(sheetName, 12, 4, null),
					new ExpectedCellValue(sheetName, 1, 5, null),
					new ExpectedCellValue(sheetName, 2, 5, "Posting"),
					new ExpectedCellValue(sheetName, 3, 5, 0),
					new ExpectedCellValue(sheetName, 4, 5, 1d),
					new ExpectedCellValue(sheetName, 5, 5, 0),
					new ExpectedCellValue(sheetName, 6, 5, 1d),
					new ExpectedCellValue(sheetName, 7, 5, null),
					new ExpectedCellValue(sheetName, 8, 5, null),
					new ExpectedCellValue(sheetName, 9, 5, null),
					new ExpectedCellValue(sheetName, 10, 5, null),
					new ExpectedCellValue(sheetName, 11, 5, null),
					new ExpectedCellValue(sheetName, 12, 5, null),
					new ExpectedCellValue(sheetName, 1, 6, null),
					new ExpectedCellValue(sheetName, 2, 6, "Grand Total"),
					new ExpectedCellValue(sheetName, 3, 6, null),
					new ExpectedCellValue(sheetName, 4, 6, null),
					new ExpectedCellValue(sheetName, 5, 6, null),
					new ExpectedCellValue(sheetName, 6, 6, null),
					new ExpectedCellValue(sheetName, 7, 6, null),
					new ExpectedCellValue(sheetName, 8, 6, null),
					new ExpectedCellValue(sheetName, 9, 6, null),
					new ExpectedCellValue(sheetName, 10, 6, null),
					new ExpectedCellValue(sheetName, 11, 6, null),
					new ExpectedCellValue(sheetName, 12, 6, null)
				});
			}
		}

	}
}
