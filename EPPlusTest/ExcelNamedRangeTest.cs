using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest
{
	[TestClass]
	public class ExcelNamedRangeTest
	{
		#region Constructor Tests
		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void ConstructExcelNamedRangeNullWorkbookThrowsException()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
				new ExcelNamedRange("somename", null, worksheet, "Sheet1!B2", 0);
			}
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentException))]
		public void ConstructExcelNamedRangeNullNameThrowsException()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
				new ExcelNamedRange(null, excelPackage.Workbook, worksheet, "Sheet1!B2", 0);
			}
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentException))]
		public void ConstructExcelNamedRangeEmptyNameThrowsException()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
				new ExcelNamedRange(string.Empty, excelPackage.Workbook, worksheet, "Sheet1!B2", 0);
			}
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentException))]
		public void ConstructExcelNamedRangeNullFormulaThrowsException()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
				new ExcelNamedRange("somename", excelPackage.Workbook, worksheet, null, 0);
			}
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentException))]
		public void ConstructExcelNamedRangeEmptyFormulaThrowsException()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
				new ExcelNamedRange(string.Empty, excelPackage.Workbook, worksheet, string.Empty, 0);
			}
		}
		#endregion

		#region ActualSheetID and LocalSheetID Tests
		[TestMethod]
		public void SheetIDOnWorkbookScopedNamedRangeIsConstant()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				var name = package.Workbook.Names.Add("MyNamedRange", new ExcelRangeBase(sheet, "$A$1"));
				Assert.AreEqual(-1, name.LocalSheetID);
				Assert.IsNull(name.LocalSheet);
			}
		}

		[TestMethod]
		public void SheetIDOnWorksheetScopedNamedRange()
		{
			using (var package = new ExcelPackage())
			{
				for (int i = 0; i < 10; i++)
					package.Workbook.Worksheets.Add($"Sheet {i}");
				var sheet = package.Workbook.Worksheets.Add("Test Sheet");
				var name = sheet.Names.Add("MyNamedRange", new ExcelRangeBase(sheet, "$A$1"));
				Assert.AreEqual(sheet.PositionID - 1, name.LocalSheetID);
				Assert.AreSame(sheet, name.LocalSheet);
			}
		}
		#endregion

		#region GetRelativeAddress Tests
		[TestMethod]
		public virtual void TryGetNamedRangeAddress()
		{
			string sheetName1 = "Sheet1", sheetName2 = "Name with spaces";
			using (var package = new ExcelPackage())
			using (var workbook = package.Workbook)
			using (var worksheet1 = workbook.Worksheets.Add(sheetName1))
			using (var worksheet2 = workbook.Worksheets.Add(sheetName2))
			using (var cell1 = worksheet1.Cells["$D3"])
			using (var cell2 = worksheet1.Cells["G$5"])
			using (var cell3 = worksheet2.Cells["E5"])
			using (var cell4 = worksheet2.Cells["$C$8"])
			{
				var names = workbook.Names;
				names.Add("Name1", cell1);
				names.Add("Name2", cell2);
				names.Add("Name3", cell3);
				names.Add("Name4", cell4);

				// Relative references to a named range are relative to cell A1. Offsets that cause the 
				// relative address to exceed the maximum row or column will wrap around.
				// Examples:
				//	$B2 means on column B and down one row from the relative address.
				//	D$5 means on row 5 and right three columns from the relative address.
				//	C3 means right two and down three from the relative address.
				var expected = $"'{sheetName1}'!$D4";
				string actual = string.Join(string.Empty, package.Workbook.Names["Name1"].GetRelativeNameFormula(2, 2).Select(t => t.Value).ToList());
				Assert.AreEqual(expected, actual);
				expected = $"'{sheetName1}'!H$5";
				actual = string.Join(string.Empty, package.Workbook.Names["Name2"].GetRelativeNameFormula(2, 2).Select(t => t.Value).ToList());
				Assert.AreEqual(expected, actual);
				expected = $"'{sheetName2}'!F6";
				actual = string.Join(string.Empty, package.Workbook.Names["Name3"].GetRelativeNameFormula(2, 2).Select(t => t.Value).ToList());
				Assert.AreEqual(expected, actual);
				expected = $"'{sheetName2}'!$C$8";
				actual = string.Join(string.Empty, package.Workbook.Names["Name4"].GetRelativeNameFormula(2, 2).Select(t => t.Value).ToList());
				Assert.AreEqual(expected, actual);
			}
		}
		#endregion

		#region UpdateFormula Tests
		[TestMethod]
		public void UpdateFormulaRelativeRowUpdateDoesNotChange()
		{
			const string formula = "Sheet1!E5";
			using (var excelPackage = new ExcelPackage())
			{
				var worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
				var namedRange = new ExcelNamedRange("someName", excelPackage.Workbook, worksheet, formula, 0);
				namedRange.UpdateFormula(2, 0, 4, 0, worksheet);
				Assert.AreEqual("'SHEET1'!E5", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void UpdateFormulaAbsoluteRowUpdatesReference()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
				var namedRange = new ExcelNamedRange("someName", excelPackage.Workbook, worksheet, "Sheet1!E$5", 0);
				namedRange.UpdateFormula(2, 0, 4, 0, worksheet);
				Assert.AreEqual("'SHEET1'!E$9", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void UpdateFormulaAbsoluteRowDeleteUpdatesReference()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
				var namedRange = new ExcelNamedRange("someName", excelPackage.Workbook, worksheet, "Sheet1!E$5", 0);
				namedRange.UpdateFormula(2, 0, -2, 0, worksheet);
				Assert.AreEqual("'SHEET1'!E$3", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void UpdateFormulaRelativeColumnUpdateDoesNotChange()
		{
			const string formula = "Sheet1!E5";
			using (var excelPackage = new ExcelPackage())
			{
				var worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
				var namedRange = new ExcelNamedRange("someName", excelPackage.Workbook, worksheet, formula, 0);
				namedRange.UpdateFormula(0, 2, 0, 5, worksheet);
				Assert.AreEqual("'SHEET1'!E5", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void UpdateFormulaAbsoluteColumnUpdatesReference()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
				var namedRange = new ExcelNamedRange("someName", excelPackage.Workbook, worksheet, "Sheet1!$E5", 0);
				namedRange.UpdateFormula(0, 2, 0, 4, worksheet);
				Assert.AreEqual("'SHEET1'!$I5", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void UpdateFormulaAbsoluteColumnDeleteUpdatesReference()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
				var namedRange = new ExcelNamedRange("someName", excelPackage.Workbook, worksheet, "Sheet1!$E5", 0);
				namedRange.UpdateFormula(0, 2, 0, -2, worksheet);
				Assert.AreEqual("'SHEET1'!$C5", namedRange.NameFormula);
			}
		}


		[TestMethod]
		public void UpdateFormulaAbsolutePartialFixedRangeUpdatesReference()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
				var namedRange = new ExcelNamedRange("someName", excelPackage.Workbook, worksheet, "sheet1!$E5:U$7", 0);
				namedRange.UpdateFormula(3, 3, 2, 4, worksheet);
				Assert.AreEqual("'SHEET1'!$I5:U$9", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void UpdateFormulaSpaceSeparatedColumnDeleteUpdatesReferences()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
				var namedRange = new ExcelNamedRange("someName", excelPackage.Workbook, worksheet, "Sheet1!$L$10 Sheet1!$E5 Sheet1!G$8 Sheet1!F12", 0);
				namedRange.UpdateFormula(0, 2, 0, -2, worksheet);
				Assert.AreEqual("'SHEET1'!$K$10 'SHEET1'!$C5 'SHEET1'!E$8 SHEET1!C12", namedRange.NameFormula);
			}
		}

		[TestMethod]
		public void UpdateFormulaCommaSeparatedColumnDeleteUpdatesReferences()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
				var namedRange = new ExcelNamedRange("someName", excelPackage.Workbook, worksheet, "Sheet1!$L$10,Sheet1!G$8,Sheet1!F12,Sheet1!$E5", 0);
				namedRange.UpdateFormula(0, 2, 0, -2, worksheet);
				Assert.AreEqual("'SHEET1'!$J$10,'SHEET1'!G$8,'SHEET1'!F12,'SHEET1'!$C5", namedRange.NameFormula);
			}
		}
		#endregion

		#region TryGetAsAddress Tests
		[TestMethod]
		public void TryGetAddressWithSingleAddress()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
				var expectedAddress = new ExcelRange(worksheet, "SHEET1!$F$8");
				var namedRange = new ExcelNamedRange("someName", excelPackage.Workbook, worksheet, "Sheet1!$F$8", 0);
				bool success = namedRange.TryGetAsAddress(out ExcelRange address);
				Assert.IsTrue(success);
				Assert.AreEqual(expectedAddress.Address, address.Address);
				Assert.AreEqual(expectedAddress.FullAddress, address.FullAddress);
				Assert.AreEqual(expectedAddress.WorkSheet, address.WorkSheet);
				Assert.AreEqual(expectedAddress.Worksheet, address.Worksheet);
				Assert.AreEqual(expectedAddress.Addresses, address.Addresses);
			}
		}

		[TestMethod]
		public void TryGetAddressWithRangeAddress()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
				var expectedAddress = new ExcelRange(worksheet, "SHEET1!$F$8:S$15");
				var namedRange = new ExcelNamedRange("someName", excelPackage.Workbook, worksheet, "Sheet1!$F$8:S$15", 0);
				bool success = namedRange.TryGetAsAddress(out ExcelRange address);
				Assert.IsTrue(success);
				Assert.AreEqual(expectedAddress.Address, address.Address);
				Assert.AreEqual(expectedAddress.FullAddress, address.FullAddress);
				Assert.AreEqual(expectedAddress.WorkSheet, address.WorkSheet);
				Assert.AreEqual(expectedAddress.Worksheet, address.Worksheet);
				Assert.AreEqual(expectedAddress.Addresses, address.Addresses);
			}
		}

		[TestMethod]
		public void TryGetAddressWithCommaSeparatedAddress()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
				var expectedAddress = new ExcelRange(worksheet, "'SHEET1'!$F$8,'SHEET1'!D$5,'SHEET1'!$G7");
				var namedRange = new ExcelNamedRange("someName", excelPackage.Workbook, worksheet, "Sheet1!$F$8,Sheet1!D$5,Sheet1!$G7", 0);
				bool success = namedRange.TryGetAsAddress(out ExcelRange address);
				Assert.IsTrue(success);
				Assert.AreEqual(expectedAddress.Address, address.Address);
				Assert.AreEqual(expectedAddress.FullAddress, address.FullAddress);
				Assert.AreEqual(expectedAddress.WorkSheet, address.WorkSheet);
				Assert.AreEqual(expectedAddress.Worksheet, address.Worksheet);
				Assert.AreEqual(expectedAddress.Addresses, address.Addresses);
			}
		}

		// TODO: Test tryGetAddress with #REF addresses

		[TestMethod]
		public void TryGetAddressWithSpaceSeparatedAddress()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
				var expectedAddress = new ExcelRange(worksheet, "'SHEET1'!$F$8 'SHEET1'!D$5 'SHEET1'!$G7");
				var namedRange = new ExcelNamedRange("someName", excelPackage.Workbook, worksheet, "Sheet1!$F$8 Sheet1!D$5 Sheet1!$G7", 0);
				bool success = namedRange.TryGetAsAddress(out ExcelRange address);
				Assert.IsTrue(success);
				Assert.AreEqual(expectedAddress.Address, address.Address);
				Assert.AreEqual(expectedAddress.FullAddress, address.FullAddress);
				Assert.AreEqual(expectedAddress.WorkSheet, address.WorkSheet);
				Assert.AreEqual(expectedAddress.Worksheet, address.Worksheet);
				Assert.AreEqual(expectedAddress.Addresses, address.Addresses);
			}
		}

		[TestMethod]
		public void TryGetAddressWithAddressInFormula()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
				var namedRange = new ExcelNamedRange("someName", excelPackage.Workbook, worksheet, "CONCATENATE(Sheet1!C3)", 0);
				bool success = namedRange.TryGetAsAddress(out ExcelRange address);
				Assert.IsFalse(success);
				Assert.IsNull(address);
			}
		}

		[TestMethod]
		public void TryGetAddressWithNoAddress()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
				var namedRange = new ExcelNamedRange("someName", excelPackage.Workbook, worksheet, "ROW()", 0);
				bool success = namedRange.TryGetAsAddress(out ExcelRange address);
				Assert.IsFalse(success);
				Assert.IsNull(address);
			}
		}
		#endregion
	}
}