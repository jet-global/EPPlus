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
