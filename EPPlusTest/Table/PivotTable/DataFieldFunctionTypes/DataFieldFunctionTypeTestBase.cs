using System.IO;
using System.Linq;
using EPPlusTest.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Table.PivotTable;

namespace EPPlusTest.Table.PivotTable.DataFieldFunctionTypes
{
	[TestClass]
	public class DataFieldFunctionTypeTestBase
	{
		#region Properties
		public ExcelPivotTable PivotTable { get; set; }
		#endregion

		#region Tests
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableDataFieldFunctionTypes.xlsx")]
		public void PivotTableRefreshDataFieldSumFunction()
		{
			var file = new FileInfo("PivotTableDataFieldFunctionTypes.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["Sheet1"];
					this.PivotTable = worksheet.PivotTables["PivotTable1"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					this.ConfigurePivotTableDataFieldFunction();
					cacheDefinition.UpdateData();
					Assert.AreEqual(7, this.PivotTable.Fields.Count);
					Assert.AreEqual(0, this.PivotTable.Fields[0].Items.Count);
					Assert.AreEqual(0, this.PivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, this.PivotTable.Fields[2].Items.Count);
					Assert.AreEqual(5, this.PivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, this.PivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, this.PivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, this.PivotTable.Fields[6].Items.Count);
					foreach (var field in this.PivotTable.Fields)
					{
						if (field.Items.Count > 0)
							this.CheckFieldItems(field);
					}
					package.SaveAs(newFile.File);
				}
				string sheetName = "Sheet1";
				this.ValidatePivotTableRefreshDataField(newFile.File, sheetName);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableDataFieldFunctionTypes.xlsx")]
		public void PivotTableRefreshDataFieldSumFunctionMultipleColumnDataFields()
		{
			var file = new FileInfo("PivotTableDataFieldFunctionTypes.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["Sheet1"];
					this.PivotTable = worksheet.PivotTables["PivotTable2"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					this.ConfigurePivotTableMultipleColumnDataFieldsFunction();
					cacheDefinition.UpdateData();
					Assert.AreEqual(7, this.PivotTable.Fields.Count);
					Assert.AreEqual(0, this.PivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, this.PivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, this.PivotTable.Fields[2].Items.Count);
					Assert.AreEqual(5, this.PivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, this.PivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, this.PivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, this.PivotTable.Fields[6].Items.Count);
					foreach (var field in this.PivotTable.Fields)
					{
						if (field.Items.Count > 0)
							this.CheckFieldItems(field);
					}
					package.SaveAs(newFile.File);
				}
				string sheetName = "Sheet1";
				this.ValidatePivotTableRefreshMultipleColumnDataFields(newFile.File, sheetName);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableDataFieldFunctionTypes.xlsx")]
		public void PivotTableRefreshDataFieldSumFunctionMultipleRowDataFields()
		{
			var file = new FileInfo("PivotTableDataFieldFunctionTypes.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				using (var package = new ExcelPackage(file))
				{
					var worksheet = package.Workbook.Worksheets["Sheet1"];
					this.PivotTable = worksheet.PivotTables["PivotTable3"];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					this.ConfigurePivotTableMultipleRowDataFieldsFunction();
					cacheDefinition.UpdateData();
					Assert.AreEqual(7, this.PivotTable.Fields.Count);
					Assert.AreEqual(0, this.PivotTable.Fields[0].Items.Count);
					Assert.AreEqual(4, this.PivotTable.Fields[1].Items.Count);
					Assert.AreEqual(4, this.PivotTable.Fields[2].Items.Count);
					Assert.AreEqual(5, this.PivotTable.Fields[3].Items.Count);
					Assert.AreEqual(0, this.PivotTable.Fields[4].Items.Count);
					Assert.AreEqual(0, this.PivotTable.Fields[5].Items.Count);
					Assert.AreEqual(0, this.PivotTable.Fields[6].Items.Count);
					foreach (var field in this.PivotTable.Fields)
					{
						if (field.Items.Count > 0)
							this.CheckFieldItems(field);
					}
					package.SaveAs(newFile.File);
				}
				string sheetName = "Sheet1";
				this.ValidatePivotTableRefreshMultipleRowDataFields(newFile.File, sheetName);
			}
		}
		#endregion

		#region SingleDataField Override Methods
		public virtual void ConfigurePivotTableDataFieldFunction()
		{
			// Does not need to be configured for sum function type.
		}

		public virtual void ValidatePivotTableRefreshDataField(FileInfo file, string sheetName)
		{
			TestHelperUtility.ValidateWorksheet(file, sheetName, new[]
			{
				new ExpectedCellValue(sheetName, 14, 1, "January"),
				new ExpectedCellValue(sheetName, 15, 1, "February"),
				new ExpectedCellValue(sheetName, 16, 1, "March"),
				new ExpectedCellValue(sheetName, 17, 1, "Grand Total"),

				new ExpectedCellValue(sheetName, 13, 2, "Car Rack"),
				new ExpectedCellValue(sheetName, 14, 2, 2078.75),
				new ExpectedCellValue(sheetName, 16, 2, 831.5),
				new ExpectedCellValue(sheetName, 17, 2, 2910.25),

				new ExpectedCellValue(sheetName, 13, 3, "Sleeping Bag"),
				new ExpectedCellValue(sheetName, 15, 3, 99d),
				new ExpectedCellValue(sheetName, 17, 3, 99d),

				new ExpectedCellValue(sheetName, 13, 4, "Headlamp"),
				new ExpectedCellValue(sheetName, 16, 4, 24.99),
				new ExpectedCellValue(sheetName, 17, 4, 24.99),

				new ExpectedCellValue(sheetName, 13, 5, "Tent"),
				new ExpectedCellValue(sheetName, 15, 5, 1194d),
				new ExpectedCellValue(sheetName, 17, 5, 1194d),

				new ExpectedCellValue(sheetName, 13, 6, "Grand Total"),
				new ExpectedCellValue(sheetName, 14, 6, 2078.75),
				new ExpectedCellValue(sheetName, 15, 6, 1293d),
				new ExpectedCellValue(sheetName, 16, 6, 856.49),
				new ExpectedCellValue(sheetName, 17, 6, 4228.24)
			});
		}
		#endregion

		#region MultipleColumnDataFields Override Methods
		public virtual void ConfigurePivotTableMultipleColumnDataFieldsFunction()
		{
			// Does not need to be configured for sum function type.
		}

		public virtual void ValidatePivotTableRefreshMultipleColumnDataFields(FileInfo file, string sheetName)
		{
			TestHelperUtility.ValidateWorksheet(file, sheetName, new[]
			{
				new ExpectedCellValue(sheetName, 24, 1, "January"),
				new ExpectedCellValue(sheetName, 25, 1, "Car Rack"),
				new ExpectedCellValue(sheetName, 26, 1, "February"),
				new ExpectedCellValue(sheetName, 27, 1, "Sleeping Bag"),
				new ExpectedCellValue(sheetName, 28, 1, "Tent"),
				new ExpectedCellValue(sheetName, 29, 1, "March"),
				new ExpectedCellValue(sheetName, 30, 1, "Car Rack"),
				new ExpectedCellValue(sheetName, 31, 1, "Headlamp"),
				new ExpectedCellValue(sheetName, 32, 1, "Grand Total"),

				new ExpectedCellValue(sheetName, 22, 2, "San Francisco"),
				new ExpectedCellValue(sheetName, 23, 2, "Sum of Total"),
				new ExpectedCellValue(sheetName, 24, 2, 415.75),
				new ExpectedCellValue(sheetName, 25, 2, 415.75),
				new ExpectedCellValue(sheetName, 26, 2, 99),
				new ExpectedCellValue(sheetName, 26, 2, 99),
				new ExpectedCellValue(sheetName, 32, 2, 514.75),

				new ExpectedCellValue(sheetName, 23, 3, "Sum of Units Sold"),
				new ExpectedCellValue(sheetName, 24, 3, 1d),
				new ExpectedCellValue(sheetName, 25, 3, 1d),
				new ExpectedCellValue(sheetName, 26, 3, 1d),
				new ExpectedCellValue(sheetName, 27, 3, 1d),
				new ExpectedCellValue(sheetName, 32, 3, 2d),

				new ExpectedCellValue(sheetName, 22, 4, "Chicago"),
				new ExpectedCellValue(sheetName, 23, 4, "Sum of Total"),
				new ExpectedCellValue(sheetName, 24, 4, 831.5),
				new ExpectedCellValue(sheetName, 25, 4, 831.5),
				new ExpectedCellValue(sheetName, 29, 4, 24.99),
				new ExpectedCellValue(sheetName, 31, 4, 24.99),
				new ExpectedCellValue(sheetName, 32, 4, 856.49),

				new ExpectedCellValue(sheetName, 23, 5, "Sum of Units Sold"),
				new ExpectedCellValue(sheetName, 24, 5, 2d),
				new ExpectedCellValue(sheetName, 25, 5, 2d),
				new ExpectedCellValue(sheetName, 29, 5, 1d),
				new ExpectedCellValue(sheetName, 31, 5, 1d),
				new ExpectedCellValue(sheetName, 32, 5, 3d),

				new ExpectedCellValue(sheetName, 22, 6, "Nashville"),
				new ExpectedCellValue(sheetName, 23, 6, "Sum of Total"),
				new ExpectedCellValue(sheetName, 24, 6, 831.5),
				new ExpectedCellValue(sheetName, 25, 6, 831.5),
				new ExpectedCellValue(sheetName, 26, 6, 1194d),
				new ExpectedCellValue(sheetName, 28, 6, 1194d),
				new ExpectedCellValue(sheetName, 29, 6, 831.5),
				new ExpectedCellValue(sheetName, 30, 6, 831.5),
				new ExpectedCellValue(sheetName, 32, 6, 2857d),

				new ExpectedCellValue(sheetName, 23, 7, "Sum of Units Sold"),
				new ExpectedCellValue(sheetName, 24, 7, 2d),
				new ExpectedCellValue(sheetName, 25, 7, 2d),
				new ExpectedCellValue(sheetName, 26, 7, 6d),
				new ExpectedCellValue(sheetName, 28, 7, 6d),
				new ExpectedCellValue(sheetName, 29, 7, 2d),
				new ExpectedCellValue(sheetName, 30, 7, 2d),
				new ExpectedCellValue(sheetName, 32, 7, 10d),

				new ExpectedCellValue(sheetName, 22, 8, "Total Sum of Total"),
				new ExpectedCellValue(sheetName, 24, 8, 2078.75),
				new ExpectedCellValue(sheetName, 25, 8, 2078.75),
				new ExpectedCellValue(sheetName, 26, 8, 1293d),
				new ExpectedCellValue(sheetName, 27, 8, 99d),
				new ExpectedCellValue(sheetName, 28, 8, 1194d),
				new ExpectedCellValue(sheetName, 29, 8, 856.49),
				new ExpectedCellValue(sheetName, 30, 8, 831.5),
				new ExpectedCellValue(sheetName, 31, 8, 24.99),
				new ExpectedCellValue(sheetName, 32, 8, 4228.24),

				new ExpectedCellValue(sheetName, 22, 9, "Total Sum of Units Sold"),
				new ExpectedCellValue(sheetName, 24, 9, 5d),
				new ExpectedCellValue(sheetName, 25, 9, 5d),
				new ExpectedCellValue(sheetName, 26, 9, 7d),
				new ExpectedCellValue(sheetName, 27, 9, 1d),
				new ExpectedCellValue(sheetName, 28, 9, 6d),
				new ExpectedCellValue(sheetName, 29, 9, 3d),
				new ExpectedCellValue(sheetName, 30, 9, 2d),
				new ExpectedCellValue(sheetName, 31, 9, 1d),
				new ExpectedCellValue(sheetName, 32, 9, 15d)
			});
		}
		#endregion

		#region MultipleRowDataFields Override Methods
		public virtual void ConfigurePivotTableMultipleRowDataFieldsFunction()
		{
			// Does not need to be configured for sum function type.
		}

		public virtual void ValidatePivotTableRefreshMultipleRowDataFields(FileInfo file, string sheetName)
		{
			TestHelperUtility.ValidateWorksheet(file, sheetName, new[]
			{
				new ExpectedCellValue(sheetName, 39, 1, "January"),
				new ExpectedCellValue(sheetName, 40, 1, "Sum of Total"),
				new ExpectedCellValue(sheetName, 41, 1, "Sum of Units Sold"),
				new ExpectedCellValue(sheetName, 42, 1, "February"),
				new ExpectedCellValue(sheetName, 43, 1, "Sum of Total"),
				new ExpectedCellValue(sheetName, 44, 1, "Sum of Units Sold"),
				new ExpectedCellValue(sheetName, 45, 1, "March"),
				new ExpectedCellValue(sheetName, 46, 1, "Sum of Total"),
				new ExpectedCellValue(sheetName, 47, 1, "Sum of Units Sold"),
				new ExpectedCellValue(sheetName, 48, 1, "Total Sum of Total"),
				new ExpectedCellValue(sheetName, 49, 1, "Total Sum of Units Sold"),

				new ExpectedCellValue(sheetName, 37, 2, "San Francisco"),
				new ExpectedCellValue(sheetName, 38, 2, "Car Rack"),
				new ExpectedCellValue(sheetName, 40, 2, 415.75),
				new ExpectedCellValue(sheetName, 41, 2, 1d),
				new ExpectedCellValue(sheetName, 48, 2, 415.75),
				new ExpectedCellValue(sheetName, 49, 2, 1d),

				new ExpectedCellValue(sheetName, 38, 3, "Sleeping Bag"),
				new ExpectedCellValue(sheetName, 43, 3, 99d),
				new ExpectedCellValue(sheetName, 44, 3, 1d),
				new ExpectedCellValue(sheetName, 48, 3, 99d),
				new ExpectedCellValue(sheetName, 49, 3, 1d),

				new ExpectedCellValue(sheetName, 37, 4, "San Francisco Total"),
				new ExpectedCellValue(sheetName, 40, 4, 415.75),
				new ExpectedCellValue(sheetName, 41, 4, 1d),
				new ExpectedCellValue(sheetName, 43, 4, 99d),
				new ExpectedCellValue(sheetName, 44, 4, 1d),
				new ExpectedCellValue(sheetName, 48, 4, 514.75),
				new ExpectedCellValue(sheetName, 49, 4, 2d),

				new ExpectedCellValue(sheetName, 37, 5, "Chicago"),
				new ExpectedCellValue(sheetName, 38, 5, "Car Rack"),
				new ExpectedCellValue(sheetName, 40, 5, 831.5),
				new ExpectedCellValue(sheetName, 41, 5, 2d),
				new ExpectedCellValue(sheetName, 48, 5, 831.5),
				new ExpectedCellValue(sheetName, 49, 5, 2d),

				new ExpectedCellValue(sheetName, 38, 6, "Headlamp"),
				new ExpectedCellValue(sheetName, 46, 6, 24.99),
				new ExpectedCellValue(sheetName, 47, 6, 1d),
				new ExpectedCellValue(sheetName, 48, 6, 24.99),
				new ExpectedCellValue(sheetName, 49, 6, 1d),

				new ExpectedCellValue(sheetName, 37, 7, "Chicago Total"),
				new ExpectedCellValue(sheetName, 40, 7, 831.5),
				new ExpectedCellValue(sheetName, 41, 7, 2d),
				new ExpectedCellValue(sheetName, 46, 7, 24.99),
				new ExpectedCellValue(sheetName, 47, 7, 1d),
				new ExpectedCellValue(sheetName, 48, 7, 856.49),
				new ExpectedCellValue(sheetName, 49, 7, 3d),

				new ExpectedCellValue(sheetName, 37, 8, "Nashville"),
				new ExpectedCellValue(sheetName, 38, 8, "Car Rack"),
				new ExpectedCellValue(sheetName, 40, 8, 831.5),
				new ExpectedCellValue(sheetName, 41, 8, 2d),
				new ExpectedCellValue(sheetName, 46, 8, 831.5),
				new ExpectedCellValue(sheetName, 47, 8, 2d),
				new ExpectedCellValue(sheetName, 48, 8, 1663d),
				new ExpectedCellValue(sheetName, 49, 8, 4d),

				new ExpectedCellValue(sheetName, 38, 9, "Tent"),
				new ExpectedCellValue(sheetName, 43, 9, 1194d),
				new ExpectedCellValue(sheetName, 44, 9, 6d),
				new ExpectedCellValue(sheetName, 48, 9, 1194d),
				new ExpectedCellValue(sheetName, 49, 9, 6d),

				new ExpectedCellValue(sheetName, 37, 10, "Nashville Total"),
				new ExpectedCellValue(sheetName, 40, 10, 831.5),
				new ExpectedCellValue(sheetName, 41, 10, 2d),
				new ExpectedCellValue(sheetName, 43, 10, 1194d),
				new ExpectedCellValue(sheetName, 44, 10, 6d),
				new ExpectedCellValue(sheetName, 46, 10, 831.5),
				new ExpectedCellValue(sheetName, 47, 10, 2d),
				new ExpectedCellValue(sheetName, 48, 10, 2857d),
				new ExpectedCellValue(sheetName, 49, 10, 10d),

				new ExpectedCellValue(sheetName, 37, 11, "Grand Total"),
				new ExpectedCellValue(sheetName, 40, 11, 2078.75),
				new ExpectedCellValue(sheetName, 41, 11, 5d),
				new ExpectedCellValue(sheetName, 43, 11, 1293d),
				new ExpectedCellValue(sheetName, 44, 11, 7d),
				new ExpectedCellValue(sheetName, 46, 11, 856.49),
				new ExpectedCellValue(sheetName, 47, 11, 3d),
				new ExpectedCellValue(sheetName, 48, 11, 4228.24),
				new ExpectedCellValue(sheetName, 49, 11, 15d)
			});
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