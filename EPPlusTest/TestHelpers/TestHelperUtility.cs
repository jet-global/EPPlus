using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Utils;

namespace EPPlusTest.TestHelpers
{
	internal class TestHelperUtility
	{
		#region Public Methods
		public static void ValidateWorkbook(FileInfo newFile, Dictionary<string, IEnumerable<ExpectedCellValue>> expectedSheetInfo)
		{
			using (var excelPackage = new ExcelPackage(newFile))
			{
				Assert.AreEqual(expectedSheetInfo.Count, excelPackage.Workbook.Worksheets.Count);
				foreach (var sheetName in expectedSheetInfo.Keys)
				{
					TestHelperUtility.ValidateWorksheet(excelPackage, sheetName, expectedSheetInfo[sheetName]);
				}
			}
		}

		public static void ValidateWorksheet(FileInfo newFile, string sheetName, IEnumerable<ExpectedCellValue> expectedCellValues)
		{
			using (var excelPackage = new ExcelPackage(newFile))
			{
				TestHelperUtility.ValidateWorksheet(excelPackage, sheetName, expectedCellValues);
			}
		}
		#endregion

		#region Private Methods
		private static void ValidateWorksheet(ExcelPackage package, string sheetName, IEnumerable<ExpectedCellValue> expectedCellValues)
		{
			var worksheet = package.Workbook.Worksheets[sheetName];
			Assert.IsNotNull(worksheet, $"The sheet {sheetName} was not found in the workbook.");
			foreach (ExpectedCellValue cell in expectedCellValues)
			{
				var range = worksheet.Cells[cell.Row, cell.Column];
				TestHelperUtility.ValidateRange(cell, range.Value);
				if (cell.Formula != null)
					Assert.AreEqual(cell.Formula, range.Formula);
			}
		}

		private static void ValidateRange(ExpectedCellValue expectedCell, object value)
		{
			Assert.IsNotNull(expectedCell);
			if (TestHelperUtility.ValueIsNumeric(value))
			{
				Assert.IsTrue(TestHelperUtility.ValueIsNumeric(expectedCell.Value), $"Expected non-numeric cell value at address [{expectedCell.Sheet},{expectedCell.Row},{expectedCell.Column}]. Actual value: {value}.");
				TestHelperUtility.AssertNumericValuesAreEqual(expectedCell.Value, value, expectedCell.Sheet, expectedCell.Row, expectedCell.Column);
			}
			else if (value is ExcelErrorValue errorValue)
			{
				Assert.IsTrue(ExcelErrorValue.Values.TryGetErrorType(expectedCell.Value.ToString(), out eErrorType errorType), 
					$"Expected {expectedCell.Value}, actual {value}");
				Assert.AreEqual(errorType, errorValue.Type);
			}
			else
				Assert.AreEqual(expectedCell.Value, value, $"Cells at address [{expectedCell.Sheet},{expectedCell.Row},{expectedCell.Column}] do not match.");
		}

		private static bool ValueIsNumeric(object value)
		{
			if (ConvertUtil.IsNumeric(value))
				return true;
			if (value is string && double.TryParse((string)value, out var parsedValue))
				return true;
			return false;
		}

		private static void AssertNumericValuesAreEqual(object value1, object value2, string expectedSheet, int expectedRow, int expectedColumn)
		{
			var doubleValue1 = Convert.ToDouble(value1);
			var doubleValue2 = Convert.ToDouble(value2);
			Assert.AreEqual(Math.Round(doubleValue1, 2), Math.Round(doubleValue2, 2), $"Values at [{expectedSheet},{expectedRow},{expectedColumn}] do not match.");
		}
		#endregion
	}
}
