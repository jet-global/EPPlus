using System.IO;
using System.Linq;
using EPPlusTest.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest.Table.PivotTable.PivotTableRefresh
{
	[TestClass]
	public class PivotTableRefreshIndentationTest
	{
		#region Test Methods
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithMultipleRowDataFieldsNoColumnField.xlsx")]
		public void PivotTableRefreshAllFieldsCompactIndentation()
		{
			var file = new FileInfo("PivotTableWithMultipleRowDataFieldsNoColumnField.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "RowDataFields";
				using (var package = new ExcelPackage(file))
				{
					var sheet = package.Workbook.Worksheets[sheetName];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					package.SaveAs(newFile.File);
				}
				var pt1expected = new[]
				{
					new ExpectedCellValue(sheetName, 1, 1, 0),
					new ExpectedCellValue(sheetName, 2, 1, 0),
					new ExpectedCellValue(sheetName, 3, 1, 1),
					new ExpectedCellValue(sheetName, 4, 1, 1),
					new ExpectedCellValue(sheetName, 5, 1, 0),
					new ExpectedCellValue(sheetName, 6, 1, 1),
					new ExpectedCellValue(sheetName, 7, 1, 1),
					new ExpectedCellValue(sheetName, 8, 1, 0),
					new ExpectedCellValue(sheetName, 9, 1, 1),
					new ExpectedCellValue(sheetName, 10, 1, 1),
					new ExpectedCellValue(sheetName, 11, 1, 0),
					new ExpectedCellValue(sheetName, 12, 1, 0)
				};
				var pt2expected = new[]
				{
					new ExpectedCellValue(sheetName, 1, 6, 0),
					new ExpectedCellValue(sheetName, 2, 6, 0),
					new ExpectedCellValue(sheetName, 3, 6, 1),
					new ExpectedCellValue(sheetName, 4, 6, 2),
					new ExpectedCellValue(sheetName, 5, 6, 2),
					new ExpectedCellValue(sheetName, 6, 6, 2),
					new ExpectedCellValue(sheetName, 7, 6, 1),
					new ExpectedCellValue(sheetName, 8, 6, 2),
					new ExpectedCellValue(sheetName, 9, 6, 2),
					new ExpectedCellValue(sheetName, 10, 6, 2),
					new ExpectedCellValue(sheetName, 11, 6, 0),
					new ExpectedCellValue(sheetName, 12, 6, 0),
					new ExpectedCellValue(sheetName, 13, 6, 0),
					new ExpectedCellValue(sheetName, 14, 6, 1),
					new ExpectedCellValue(sheetName, 15, 6, 2),
					new ExpectedCellValue(sheetName, 16, 6, 2),
					new ExpectedCellValue(sheetName, 17, 6, 1),
					new ExpectedCellValue(sheetName, 18, 6, 2),
					new ExpectedCellValue(sheetName, 19, 6, 2),
					new ExpectedCellValue(sheetName, 20, 6, 0),
					new ExpectedCellValue(sheetName, 21, 6, 0),
					new ExpectedCellValue(sheetName, 22, 6, 0),
					new ExpectedCellValue(sheetName, 23, 6, 1),
					new ExpectedCellValue(sheetName, 24, 6, 2),
					new ExpectedCellValue(sheetName, 25, 6, 2),
					new ExpectedCellValue(sheetName, 26, 6, 1),
					new ExpectedCellValue(sheetName, 27, 6, 2),
					new ExpectedCellValue(sheetName, 28, 6, 2),
					new ExpectedCellValue(sheetName, 29, 6, 0),
					new ExpectedCellValue(sheetName, 30, 6, 0),
					new ExpectedCellValue(sheetName, 31, 6, 0),
					new ExpectedCellValue(sheetName, 32, 6, 0)
				};
				var pt3expected = new[]
				{
					new ExpectedCellValue(sheetName, 1, 11, 0),
					new ExpectedCellValue(sheetName, 2, 11, 0),
					new ExpectedCellValue(sheetName, 3, 11, 1),
					new ExpectedCellValue(sheetName, 4, 11, 2),
					new ExpectedCellValue(sheetName, 5, 11, 2),
					new ExpectedCellValue(sheetName, 6, 11, 1),
					new ExpectedCellValue(sheetName, 7, 11, 2),
					new ExpectedCellValue(sheetName, 8, 11, 2),
					new ExpectedCellValue(sheetName, 9, 11, 1),
					new ExpectedCellValue(sheetName, 10, 11, 2),
					new ExpectedCellValue(sheetName, 11, 11, 2),
					new ExpectedCellValue(sheetName, 12, 11, 0),
					new ExpectedCellValue(sheetName, 13, 11, 0),
					new ExpectedCellValue(sheetName, 14, 11, 0),
					new ExpectedCellValue(sheetName, 15, 11, 1),
					new ExpectedCellValue(sheetName, 16, 11, 2),
					new ExpectedCellValue(sheetName, 17, 11, 2),
					new ExpectedCellValue(sheetName, 18, 11, 1),
					new ExpectedCellValue(sheetName, 19, 11, 2),
					new ExpectedCellValue(sheetName, 20, 11, 2),
					new ExpectedCellValue(sheetName, 21, 11, 0),
					new ExpectedCellValue(sheetName, 22, 11, 0),
					new ExpectedCellValue(sheetName, 23, 11, 0),
					new ExpectedCellValue(sheetName, 24, 11, 1),
					new ExpectedCellValue(sheetName, 25, 11, 2),
					new ExpectedCellValue(sheetName, 26, 11, 2),
					new ExpectedCellValue(sheetName, 27, 11, 1),
					new ExpectedCellValue(sheetName, 28, 11, 2),
					new ExpectedCellValue(sheetName, 29, 11, 2),
					new ExpectedCellValue(sheetName, 30, 11, 0),
					new ExpectedCellValue(sheetName, 31, 11, 0),
					new ExpectedCellValue(sheetName, 32, 11, 0),
					new ExpectedCellValue(sheetName, 33, 11, 0)
				};
				this.ValidateIndentation(newFile.File, pt1expected);
				this.ValidateIndentation(newFile.File, pt2expected);
				this.ValidateIndentation(newFile.File, pt3expected);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithMultipleRowDataFieldsNoColumnField.xlsx")]
		public void PivotTableRefreshAllFieldsCompactIndentationWith8CharacterIndent()
		{
			var file = new FileInfo("PivotTableWithMultipleRowDataFieldsNoColumnField.xlsx");
			Assert.IsTrue(file.Exists);
			int indentMultiplier = 8;
			using (var newFile = new TempTestFile())
			{
				string sheetName = "RowDataFields";
				using (var package = new ExcelPackage(file))
				{
					var sheet = package.Workbook.Worksheets[sheetName];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					foreach (var pivotTable in sheet.PivotTables)
					{
						pivotTable.Indent = indentMultiplier;
					}
					cacheDefinition.UpdateData();
					package.SaveAs(newFile.File);
				}
				var pt1expected = new[]
				{
					new ExpectedCellValue(sheetName, 1, 1, 0),
					new ExpectedCellValue(sheetName, 2, 1, 0),
					new ExpectedCellValue(sheetName, 3, 1, 1),
					new ExpectedCellValue(sheetName, 4, 1, 1),
					new ExpectedCellValue(sheetName, 5, 1, 0),
					new ExpectedCellValue(sheetName, 6, 1, 1),
					new ExpectedCellValue(sheetName, 7, 1, 1),
					new ExpectedCellValue(sheetName, 8, 1, 0),
					new ExpectedCellValue(sheetName, 9, 1, 1),
					new ExpectedCellValue(sheetName, 10, 1, 1),
					new ExpectedCellValue(sheetName, 11, 1, 0),
					new ExpectedCellValue(sheetName, 12, 1, 0)
				};
				var pt2expected = new[]
				{
					new ExpectedCellValue(sheetName, 1, 6, 0),
					new ExpectedCellValue(sheetName, 2, 6, 0),
					new ExpectedCellValue(sheetName, 3, 6, 1),
					new ExpectedCellValue(sheetName, 4, 6, 2),
					new ExpectedCellValue(sheetName, 5, 6, 2),
					new ExpectedCellValue(sheetName, 6, 6, 2),
					new ExpectedCellValue(sheetName, 7, 6, 1),
					new ExpectedCellValue(sheetName, 8, 6, 2),
					new ExpectedCellValue(sheetName, 9, 6, 2),
					new ExpectedCellValue(sheetName, 10, 6, 2),
					new ExpectedCellValue(sheetName, 11, 6, 0),
					new ExpectedCellValue(sheetName, 12, 6, 0),
					new ExpectedCellValue(sheetName, 13, 6, 0),
					new ExpectedCellValue(sheetName, 14, 6, 1),
					new ExpectedCellValue(sheetName, 15, 6, 2),
					new ExpectedCellValue(sheetName, 16, 6, 2),
					new ExpectedCellValue(sheetName, 17, 6, 1),
					new ExpectedCellValue(sheetName, 18, 6, 2),
					new ExpectedCellValue(sheetName, 19, 6, 2),
					new ExpectedCellValue(sheetName, 20, 6, 0),
					new ExpectedCellValue(sheetName, 21, 6, 0),
					new ExpectedCellValue(sheetName, 22, 6, 0),
					new ExpectedCellValue(sheetName, 23, 6, 1),
					new ExpectedCellValue(sheetName, 24, 6, 2),
					new ExpectedCellValue(sheetName, 25, 6, 2),
					new ExpectedCellValue(sheetName, 26, 6, 1),
					new ExpectedCellValue(sheetName, 27, 6, 2),
					new ExpectedCellValue(sheetName, 28, 6, 2),
					new ExpectedCellValue(sheetName, 29, 6, 0),
					new ExpectedCellValue(sheetName, 30, 6, 0),
					new ExpectedCellValue(sheetName, 31, 6, 0),
					new ExpectedCellValue(sheetName, 32, 6, 0)
				};
				var pt3expected = new[]
				{
					new ExpectedCellValue(sheetName, 1, 11, 0),
					new ExpectedCellValue(sheetName, 2, 11, 0),
					new ExpectedCellValue(sheetName, 3, 11, 1),
					new ExpectedCellValue(sheetName, 4, 11, 2),
					new ExpectedCellValue(sheetName, 5, 11, 2),
					new ExpectedCellValue(sheetName, 6, 11, 1),
					new ExpectedCellValue(sheetName, 7, 11, 2),
					new ExpectedCellValue(sheetName, 8, 11, 2),
					new ExpectedCellValue(sheetName, 9, 11, 1),
					new ExpectedCellValue(sheetName, 10, 11, 2),
					new ExpectedCellValue(sheetName, 11, 11, 2),
					new ExpectedCellValue(sheetName, 12, 11, 0),
					new ExpectedCellValue(sheetName, 13, 11, 0),
					new ExpectedCellValue(sheetName, 14, 11, 0),
					new ExpectedCellValue(sheetName, 15, 11, 1),
					new ExpectedCellValue(sheetName, 16, 11, 2),
					new ExpectedCellValue(sheetName, 17, 11, 2),
					new ExpectedCellValue(sheetName, 18, 11, 1),
					new ExpectedCellValue(sheetName, 19, 11, 2),
					new ExpectedCellValue(sheetName, 20, 11, 2),
					new ExpectedCellValue(sheetName, 21, 11, 0),
					new ExpectedCellValue(sheetName, 22, 11, 0),
					new ExpectedCellValue(sheetName, 23, 11, 0),
					new ExpectedCellValue(sheetName, 24, 11, 1),
					new ExpectedCellValue(sheetName, 25, 11, 2),
					new ExpectedCellValue(sheetName, 26, 11, 2),
					new ExpectedCellValue(sheetName, 27, 11, 1),
					new ExpectedCellValue(sheetName, 28, 11, 2),
					new ExpectedCellValue(sheetName, 29, 11, 2),
					new ExpectedCellValue(sheetName, 30, 11, 0),
					new ExpectedCellValue(sheetName, 31, 11, 0),
					new ExpectedCellValue(sheetName, 32, 11, 0),
					new ExpectedCellValue(sheetName, 33, 11, 0)
				};
				this.ValidateIndentation(newFile.File, pt1expected, indentMultiplier);
				this.ValidateIndentation(newFile.File, pt2expected, indentMultiplier);
				this.ValidateIndentation(newFile.File, pt3expected, indentMultiplier);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableWithMultipleRowDataFieldsNoColumnField.xlsx")]
		public void PivotTableRefreshAllFieldsCompactIndentationWith0CharacterIndent()
		{
			var file = new FileInfo("PivotTableWithMultipleRowDataFieldsNoColumnField.xlsx");
			Assert.IsTrue(file.Exists);
			int indentMultiplier = 127; // Excel stores zero indent as 127.
			using (var newFile = new TempTestFile())
			{
				string sheetName = "RowDataFields";
				using (var package = new ExcelPackage(file))
				{
					var sheet = package.Workbook.Worksheets[sheetName];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					foreach (var pivotTable in sheet.PivotTables)
					{
						pivotTable.Indent = indentMultiplier;
					}
					cacheDefinition.UpdateData();
					package.SaveAs(newFile.File);
				}
				var pt1expected = new[]
				{
					new ExpectedCellValue(sheetName, 1, 1, 0),
					new ExpectedCellValue(sheetName, 2, 1, 0),
					new ExpectedCellValue(sheetName, 3, 1, 0),
					new ExpectedCellValue(sheetName, 4, 1, 0),
					new ExpectedCellValue(sheetName, 5, 1, 0),
					new ExpectedCellValue(sheetName, 6, 1, 0),
					new ExpectedCellValue(sheetName, 7, 1, 0),
					new ExpectedCellValue(sheetName, 8, 1, 0),
					new ExpectedCellValue(sheetName, 9, 1, 0),
					new ExpectedCellValue(sheetName, 10, 1, 0),
					new ExpectedCellValue(sheetName, 11, 1, 0),
					new ExpectedCellValue(sheetName, 12, 1, 0)
				};
				var pt2expected = new[]
				{
					new ExpectedCellValue(sheetName, 1, 6, 0),
					new ExpectedCellValue(sheetName, 2, 6, 0),
					new ExpectedCellValue(sheetName, 3, 6, 0),
					new ExpectedCellValue(sheetName, 4, 6, 0),
					new ExpectedCellValue(sheetName, 5, 6, 0),
					new ExpectedCellValue(sheetName, 6, 6, 0),
					new ExpectedCellValue(sheetName, 7, 6, 0),
					new ExpectedCellValue(sheetName, 8, 6, 0),
					new ExpectedCellValue(sheetName, 9, 6, 0),
					new ExpectedCellValue(sheetName, 10, 6, 0),
					new ExpectedCellValue(sheetName, 11, 6, 0),
					new ExpectedCellValue(sheetName, 12, 6, 0),
					new ExpectedCellValue(sheetName, 13, 6, 0),
					new ExpectedCellValue(sheetName, 14, 6, 0),
					new ExpectedCellValue(sheetName, 15, 6, 0),
					new ExpectedCellValue(sheetName, 16, 6, 0),
					new ExpectedCellValue(sheetName, 17, 6, 0),
					new ExpectedCellValue(sheetName, 18, 6, 0),
					new ExpectedCellValue(sheetName, 19, 6, 0),
					new ExpectedCellValue(sheetName, 20, 6, 0),
					new ExpectedCellValue(sheetName, 21, 6, 0),
					new ExpectedCellValue(sheetName, 22, 6, 0),
					new ExpectedCellValue(sheetName, 23, 6, 0),
					new ExpectedCellValue(sheetName, 24, 6, 0),
					new ExpectedCellValue(sheetName, 25, 6, 0),
					new ExpectedCellValue(sheetName, 26, 6, 0),
					new ExpectedCellValue(sheetName, 27, 6, 0),
					new ExpectedCellValue(sheetName, 28, 6, 0),
					new ExpectedCellValue(sheetName, 29, 6, 0),
					new ExpectedCellValue(sheetName, 30, 6, 0),
					new ExpectedCellValue(sheetName, 31, 6, 0),
					new ExpectedCellValue(sheetName, 32, 6, 0)
				};
				var pt3expected = new[]
				{
					new ExpectedCellValue(sheetName, 1, 11, 0),
					new ExpectedCellValue(sheetName, 2, 11, 0),
					new ExpectedCellValue(sheetName, 3, 11, 0),
					new ExpectedCellValue(sheetName, 4, 11, 0),
					new ExpectedCellValue(sheetName, 5, 11, 0),
					new ExpectedCellValue(sheetName, 6, 11, 0),
					new ExpectedCellValue(sheetName, 7, 11, 0),
					new ExpectedCellValue(sheetName, 8, 11, 0),
					new ExpectedCellValue(sheetName, 9, 11, 0),
					new ExpectedCellValue(sheetName, 10, 11, 0),
					new ExpectedCellValue(sheetName, 11, 11, 0),
					new ExpectedCellValue(sheetName, 12, 11, 0),
					new ExpectedCellValue(sheetName, 13, 11, 0),
					new ExpectedCellValue(sheetName, 14, 11, 0),
					new ExpectedCellValue(sheetName, 15, 11, 0),
					new ExpectedCellValue(sheetName, 16, 11, 0),
					new ExpectedCellValue(sheetName, 17, 11, 0),
					new ExpectedCellValue(sheetName, 18, 11, 0),
					new ExpectedCellValue(sheetName, 19, 11, 0),
					new ExpectedCellValue(sheetName, 20, 11, 0),
					new ExpectedCellValue(sheetName, 21, 11, 0),
					new ExpectedCellValue(sheetName, 22, 11, 0),
					new ExpectedCellValue(sheetName, 23, 11, 0),
					new ExpectedCellValue(sheetName, 24, 11, 0),
					new ExpectedCellValue(sheetName, 25, 11, 0),
					new ExpectedCellValue(sheetName, 26, 11, 0),
					new ExpectedCellValue(sheetName, 27, 11, 0),
					new ExpectedCellValue(sheetName, 28, 11, 0),
					new ExpectedCellValue(sheetName, 29, 11, 0),
					new ExpectedCellValue(sheetName, 30, 11, 0),
					new ExpectedCellValue(sheetName, 31, 11, 0),
					new ExpectedCellValue(sheetName, 32, 11, 0),
					new ExpectedCellValue(sheetName, 33, 11, 0)
				};
				this.ValidateIndentation(newFile.File, pt1expected, indentMultiplier);
				this.ValidateIndentation(newFile.File, pt2expected, indentMultiplier);
				this.ValidateIndentation(newFile.File, pt3expected, indentMultiplier);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableTabularSettingsEnabled.xlsx")]
		public void PivotTableRefreshIndentationTabularEnabled()
		{
			var file = new FileInfo("PivotTableTabularSettingsEnabled.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "FourTabularFields";
				using (var package = new ExcelPackage(file))
				{
					var sheet = package.Workbook.Worksheets[sheetName];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					package.SaveAs(newFile.File);
				}
				var pt1expected = new[]
				{
					new ExpectedCellValue(sheetName, 3, 1, 0),
					new ExpectedCellValue(sheetName, 6, 1, 0),
					new ExpectedCellValue(sheetName, 10, 1, 0),
					new ExpectedCellValue(sheetName, 11, 1, 0),
					new ExpectedCellValue(sheetName, 14, 1, 0),
					new ExpectedCellValue(sheetName, 15, 1, 0),
					new ExpectedCellValue(sheetName, 18, 1, 0),
					new ExpectedCellValue(sheetName, 19, 1, 0),
					new ExpectedCellValue(sheetName, 22, 1, 0),
					new ExpectedCellValue(sheetName, 23, 1, 0),
					new ExpectedCellValue(sheetName, 26, 1, 0),
					new ExpectedCellValue(sheetName, 27, 1, 0),
					new ExpectedCellValue(sheetName, 30, 1, 0),
					new ExpectedCellValue(sheetName, 31, 1, 0),

					new ExpectedCellValue(sheetName, 3, 2, 0),
					new ExpectedCellValue(sheetName, 5, 2, 0),
					new ExpectedCellValue(sheetName, 7, 2, 0),
					new ExpectedCellValue(sheetName, 9, 2, 0),
					new ExpectedCellValue(sheetName, 11, 2, 0),
					new ExpectedCellValue(sheetName, 13, 2, 0),
					new ExpectedCellValue(sheetName, 15, 2, 0),
					new ExpectedCellValue(sheetName, 17, 2, 0),
					new ExpectedCellValue(sheetName, 19, 2, 0),
					new ExpectedCellValue(sheetName, 21, 2, 0),
					new ExpectedCellValue(sheetName, 23, 2, 0),
					new ExpectedCellValue(sheetName, 25, 2, 0),
					new ExpectedCellValue(sheetName, 27, 2, 0),
					new ExpectedCellValue(sheetName, 29, 2, 0),

					new ExpectedCellValue(sheetName, 3, 3, 0),
					new ExpectedCellValue(sheetName, 4, 3, 1),
					new ExpectedCellValue(sheetName, 7, 3, 0),
					new ExpectedCellValue(sheetName, 8, 3, 1),
					new ExpectedCellValue(sheetName, 11, 3, 0),
					new ExpectedCellValue(sheetName, 12, 3, 1),
					new ExpectedCellValue(sheetName, 15, 3, 0),
					new ExpectedCellValue(sheetName, 16, 3, 1),
					new ExpectedCellValue(sheetName, 19, 3, 0),
					new ExpectedCellValue(sheetName, 20, 3, 1),
					new ExpectedCellValue(sheetName, 23, 3, 0),
					new ExpectedCellValue(sheetName, 24, 3, 1),
					new ExpectedCellValue(sheetName, 27, 3, 0),
					new ExpectedCellValue(sheetName, 28, 3, 1),
				};
				this.ValidateIndentation(newFile.File, pt1expected);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableCompactForm.xlsx")]
		public void PivotTableRefreshIndentationCompactFormDisabled()
		{
			var file = new FileInfo("PivotTableCompactForm.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "FourRowFields";
				using (var package = new ExcelPackage(file))
				{
					var sheet = package.Workbook.Worksheets[sheetName];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					package.SaveAs(newFile.File);
				}
				var pt1expected = new[]
				{
					// Pivot Table 5
					new ExpectedCellValue(sheetName, 60, 1, 0),
					new ExpectedCellValue(sheetName, 61, 1, 0),
					new ExpectedCellValue(sheetName, 65, 1, 0),
					new ExpectedCellValue(sheetName, 69, 1, 0),
					new ExpectedCellValue(sheetName, 73, 1, 0),
					new ExpectedCellValue(sheetName, 77, 1, 0),
					new ExpectedCellValue(sheetName, 81, 1, 0),
					new ExpectedCellValue(sheetName, 85, 1, 0),
					new ExpectedCellValue(sheetName, 89, 1, 0),

					new ExpectedCellValue(sheetName, 60, 2, 0),
					new ExpectedCellValue(sheetName, 62, 2, 0),
					new ExpectedCellValue(sheetName, 63, 2, 1),
					new ExpectedCellValue(sheetName, 66, 2, 0),
					new ExpectedCellValue(sheetName, 67, 2, 1),
					new ExpectedCellValue(sheetName, 70, 2, 0),
					new ExpectedCellValue(sheetName, 71, 2, 1),
					new ExpectedCellValue(sheetName, 74, 2, 0),
					new ExpectedCellValue(sheetName, 75, 2, 1),
					new ExpectedCellValue(sheetName, 78, 2, 0),
					new ExpectedCellValue(sheetName, 79, 2, 1),
					new ExpectedCellValue(sheetName, 82, 2, 0),
					new ExpectedCellValue(sheetName, 83, 2, 1),
					new ExpectedCellValue(sheetName, 86, 2, 0),
					new ExpectedCellValue(sheetName, 87, 2, 1),

					new ExpectedCellValue(sheetName, 60, 3, 0),
					new ExpectedCellValue(sheetName, 64, 3, 0),
					new ExpectedCellValue(sheetName, 72, 3, 0),
					new ExpectedCellValue(sheetName, 76, 3, 0),
					new ExpectedCellValue(sheetName, 80, 3, 0),
					new ExpectedCellValue(sheetName, 84, 3, 0),
					new ExpectedCellValue(sheetName, 88, 3, 0),
				};
				var pt2expected = new[]
				{
					// Pivot Table 7
					new ExpectedCellValue(sheetName, 60, 11, 0),
					new ExpectedCellValue(sheetName, 61, 11, 0),
					new ExpectedCellValue(sheetName, 65, 11, 0),
					new ExpectedCellValue(sheetName, 69, 11, 0),
					new ExpectedCellValue(sheetName, 73, 11, 0),
					new ExpectedCellValue(sheetName, 77, 11, 0),
					new ExpectedCellValue(sheetName, 81, 11, 0),
					new ExpectedCellValue(sheetName, 85, 11, 0),
					new ExpectedCellValue(sheetName, 89, 11, 0),

					new ExpectedCellValue(sheetName, 60, 12, 0),
					new ExpectedCellValue(sheetName, 62, 12, 0),
					new ExpectedCellValue(sheetName, 63, 12, 1),
					new ExpectedCellValue(sheetName, 64, 12, 2),
					new ExpectedCellValue(sheetName, 66, 12, 0),
					new ExpectedCellValue(sheetName, 67, 12, 1),
					new ExpectedCellValue(sheetName, 68, 12, 2),
					new ExpectedCellValue(sheetName, 70, 12, 0),
					new ExpectedCellValue(sheetName, 71, 12, 1),
					new ExpectedCellValue(sheetName, 72, 12, 2),
					new ExpectedCellValue(sheetName, 74, 12, 0),
					new ExpectedCellValue(sheetName, 75, 12, 1),
					new ExpectedCellValue(sheetName, 76, 12, 2),
					new ExpectedCellValue(sheetName, 78, 12, 0),
					new ExpectedCellValue(sheetName, 79, 12, 1),
					new ExpectedCellValue(sheetName, 80, 12, 2),
					new ExpectedCellValue(sheetName, 82, 12, 0),
					new ExpectedCellValue(sheetName, 83, 12, 1),
					new ExpectedCellValue(sheetName, 84, 12, 2),
					new ExpectedCellValue(sheetName, 86, 12, 0),
					new ExpectedCellValue(sheetName, 87, 12, 1),
					new ExpectedCellValue(sheetName, 88, 12, 2),
				};
				this.ValidateIndentation(newFile.File, pt1expected);
				this.ValidateIndentation(newFile.File, pt2expected);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTables\PivotTableCompactAndTabularFormMultipleDataFields.xlsx")]
		public void PivotTableRefreshIndentationCompactFormDisabledAndTabularFormEnabled()
		{
			var file = new FileInfo("PivotTableCompactAndTabularFormMultipleDataFields.xlsx");
			Assert.IsTrue(file.Exists);
			using (var newFile = new TempTestFile())
			{
				string sheetName = "ThreeRowFields";
				using (var package = new ExcelPackage(file))
				{
					var sheet = package.Workbook.Worksheets[sheetName];
					var cacheDefinition = package.Workbook.PivotCacheDefinitions.Single();
					cacheDefinition.UpdateData();
					package.SaveAs(newFile.File);
				}
				var pt1expected = new[]
				{
					// Pivot Table 7
					new ExpectedCellValue(sheetName, 116, 3, 0),
					new ExpectedCellValue(sheetName, 125, 3, 0),
					new ExpectedCellValue(sheetName, 137, 3, 0),
					new ExpectedCellValue(sheetName, 145, 3, 0),
					new ExpectedCellValue(sheetName, 146, 3, 0),

					new ExpectedCellValue(sheetName, 116, 4, 0),
					new ExpectedCellValue(sheetName, 117, 4, 0),
					new ExpectedCellValue(sheetName, 121, 4, 0),
					new ExpectedCellValue(sheetName, 125, 4, 0),
					new ExpectedCellValue(sheetName, 129, 4, 0),
					new ExpectedCellValue(sheetName, 133, 4, 0),
					new ExpectedCellValue(sheetName, 137, 4, 0),
					new ExpectedCellValue(sheetName, 141, 4, 0),

					new ExpectedCellValue(sheetName, 116, 5, 0),
					new ExpectedCellValue(sheetName, 117, 5, 0),
					new ExpectedCellValue(sheetName, 118, 5, 1),
					new ExpectedCellValue(sheetName, 119, 5, 0),
					new ExpectedCellValue(sheetName, 120, 5, 1),
					new ExpectedCellValue(sheetName, 121, 5, 0),
					new ExpectedCellValue(sheetName, 122, 5, 1),
					new ExpectedCellValue(sheetName, 123, 5, 0),
					new ExpectedCellValue(sheetName, 124, 5, 1),
					new ExpectedCellValue(sheetName, 125, 5, 0),
					new ExpectedCellValue(sheetName, 126, 5, 1),
					new ExpectedCellValue(sheetName, 127, 5, 0),
					new ExpectedCellValue(sheetName, 128, 5, 1),
					new ExpectedCellValue(sheetName, 129, 5, 0),
					new ExpectedCellValue(sheetName, 130, 5, 1),
					new ExpectedCellValue(sheetName, 131, 5, 0),
					new ExpectedCellValue(sheetName, 132, 5, 1),
					new ExpectedCellValue(sheetName, 133, 5, 0),
					new ExpectedCellValue(sheetName, 134, 5, 1),
					new ExpectedCellValue(sheetName, 135, 5, 0),
					new ExpectedCellValue(sheetName, 136, 5, 1),
					new ExpectedCellValue(sheetName, 137, 5, 0),
					new ExpectedCellValue(sheetName, 138, 5, 1),
					new ExpectedCellValue(sheetName, 139, 5, 0),
					new ExpectedCellValue(sheetName, 140, 5, 1),
					new ExpectedCellValue(sheetName, 141, 5, 0),
					new ExpectedCellValue(sheetName, 142, 5, 1),
					new ExpectedCellValue(sheetName, 143, 5, 0),
					new ExpectedCellValue(sheetName, 144, 5, 1),
				};
				this.ValidateIndentation(newFile.File, pt1expected);
			}
		}
		#endregion

		#region Helper Methods
		private void ValidateIndentation(FileInfo file, ExpectedCellValue[] expectedValues, int indentMultiplier = 0)
		{
			using (var package = new ExcelPackage(file))
			{
				var workbook = package.Workbook;
				foreach (var expected in expectedValues)
				{
					var cell = workbook.Worksheets[expected.Sheet].Cells[expected.Row, expected.Column];
					Assert.AreEqual((int)expected.Value * (indentMultiplier + 1), cell.Style.Indent, 
						$"Indent at cell ({expected.Row}, {expected.Column}) did not match.");
				}
			}
		}
		#endregion
	}
}
