using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

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
		[ExpectedException(typeof(ArgumentNullException))]
		public void ConstructExcelNamedRangeNullNameThrowsException()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
				new ExcelNamedRange(null, excelPackage.Workbook, worksheet, "Sheet1!B2", 0);
			}
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void ConstructExcelNamedRangeEmptyNameThrowsException()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
				new ExcelNamedRange(string.Empty, excelPackage.Workbook, worksheet, "Sheet1!B2", 0);
			}
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void ConstructExcelNamedRangeNullFormulaThrowsException()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
				new ExcelNamedRange("somename", excelPackage.Workbook, worksheet, null, 0);
			}
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
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

		#region GetRelativeNameFormula Tests
		[TestMethod]
		public virtual void GetRelativeNameFormulaAddress()
		{
			string sheetName1 = "Sheet1", sheetName2 = "Name with spaces";
			using (var package = new ExcelPackage())
			{
				var worksheet1 = package.Workbook.Worksheets.Add(sheetName1);
				var worksheet2 = package.Workbook.Worksheets.Add(sheetName2);
				var names = package.Workbook.Names;
				names.Add("Name1", worksheet1.Cells["Sheet1!$D3"]);
				names.Add("Name2", worksheet1.Cells["Sheet1!G$5"]);
				names.Add("Name3", worksheet2.Cells["Sheet1!E5"]);
				names.Add("Name4", worksheet2.Cells["Sheet1!$C$8"]);
				names.Add("Name5", "'[Demo Waterfall Chart.xlsx]Sheet1'!E5");

				// Relative references to a named range are relative to cell A1. Offsets that cause the 
				// relative address to exceed the maximum row or column will wrap around.
				// Examples:
				//	$B2 means on column B and down one row from the relative address.
				//	D$5 means on row 5 and right three columns from the relative address.
				//	C3 means right two and down three from the relative address.
				var expected = $"'Sheet1'!$D4";
				string actual = string.Join(string.Empty, package.Workbook.Names["Name1"].GetRelativeNameFormula(2, 2).Select(t => t.Value).ToList());
				Assert.AreEqual(expected, actual);
				expected = $"'Sheet1'!H$5";
				actual = string.Join(string.Empty, package.Workbook.Names["Name2"].GetRelativeNameFormula(2, 2).Select(t => t.Value).ToList());
				Assert.AreEqual(expected, actual);
				expected = $"'Sheet1'!F6";
				actual = string.Join(string.Empty, package.Workbook.Names["Name3"].GetRelativeNameFormula(2, 2).Select(t => t.Value).ToList());
				Assert.AreEqual(expected, actual);
				expected = $"'Sheet1'!$C$8";
				actual = string.Join(string.Empty, package.Workbook.Names["Name4"].GetRelativeNameFormula(2, 2).Select(t => t.Value).ToList());
				Assert.AreEqual(expected, actual);

				// External references will not be updated.
				actual = string.Join(string.Empty, package.Workbook.Names["Name5"].GetRelativeNameFormula(2, 2).Select(t => t.Value).ToList());
				Assert.AreEqual("'[Demo Waterfall Chart.xlsx]Sheet1'!E5", actual);
			}
		}

		[TestMethod]
		public void GetRelativeNameFormulaFullColumnTest()
		{
			using (var package = new ExcelPackage())
			{
				var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
				sheet1.Names.Add("name", "'Sheet1'!$C:$C");
				IEnumerable<Token> tokens = sheet1.Names.First().GetRelativeNameFormula(4, 2);
				Assert.AreEqual(1, tokens.Count());
				var address = new ExcelAddress(tokens.First().Value);
				var x = address._isFullColumn;
				Assert.AreEqual("'Sheet1'!$C:$C", address.Address);
			}
		}

		[TestMethod]
		public void GetRelativeNameFormulaUnionWithFullRowsAndColumns()
		{
			using (var package = new ExcelPackage())
			{
				var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
				sheet1.Names.Add("name", "B5:D9,C:C,A1:A2,4:5,'Sheet1'!E:'Sheet1'!E,'DiffSheet'!4:4,$G:H");
				IEnumerable<Token> tokens = sheet1.Names.First().GetRelativeNameFormula(4, 2);
				Assert.AreEqual(13, tokens.Count());
				Assert.AreEqual("C8:E12", new ExcelAddress(tokens.ElementAt(0).Value).Address);
				Assert.AreEqual("D:D", new ExcelAddress(tokens.ElementAt(2).Value).Address);
				Assert.AreEqual("B4:B5", new ExcelAddress(tokens.ElementAt(4).Value).Address);
				Assert.AreEqual("7:8", new ExcelAddress(tokens.ElementAt(6).Value).Address);
				Assert.AreEqual("'Sheet1'!F:F", new ExcelAddress(tokens.ElementAt(8).Value).Address);
				Assert.AreEqual("'DiffSheet'!7:7", new ExcelAddress(tokens.ElementAt(10).Value).Address);
				Assert.AreEqual("$G:I", new ExcelAddress(tokens.ElementAt(12).Value).Address);
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
				Assert.AreEqual("'Sheet1'!E5", namedRange.NameFormula);
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
				Assert.AreEqual("'Sheet1'!E$9", namedRange.NameFormula);
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
				Assert.AreEqual("'Sheet1'!E$3", namedRange.NameFormula);
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
				Assert.AreEqual("'Sheet1'!E5", namedRange.NameFormula);
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
				Assert.AreEqual("'Sheet1'!$I5", namedRange.NameFormula);
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
				Assert.AreEqual("'Sheet1'!$C5", namedRange.NameFormula);
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
				Assert.AreEqual("'sheet1'!$I5:U$9", namedRange.NameFormula);
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
				Assert.AreEqual("'Sheet1'!$J$10,'Sheet1'!G$8,'Sheet1'!F12,'Sheet1'!$C5", namedRange.NameFormula);
			}
		}
		#endregion

		#region GetFormulaAsCellRange Tests
		[TestMethod]
		public void GetFormulaAsAddressTest()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var worksheet1 = excelPackage.Workbook.Worksheets.Add("Sheet1");
				var worksheet2 = excelPackage.Workbook.Worksheets.Add("Sheet2");
				worksheet1.Names.Add("name1", "Sheet1!G7");
				var address = worksheet1.Names["name1"].GetFormulaAsCellRange();
				Assert.AreEqual("Sheet1!G7", address.Address);
				Assert.AreEqual(worksheet1, address.Worksheet);

				worksheet1.Names["name1"].NameFormula = "(Sheet1!G7)";
				address = worksheet1.Names["name1"].GetFormulaAsCellRange();
				Assert.AreEqual("Sheet1!G7", address.Address);
				Assert.AreEqual(worksheet1, address.Worksheet);

				worksheet1.Names["name1"].NameFormula = "(SHeet1!B2):SHEET1!$C5";
				address = worksheet1.Names["name1"].GetFormulaAsCellRange();
				Assert.AreEqual("SHeet1!B2:SHEET1!$C5", address.Address);
				Assert.AreEqual(worksheet1, address.Worksheet);

				worksheet1.Names["name1"].NameFormula = "SHeet1!B2:SHEET1!$C5";
				address = worksheet1.Names["name1"].GetFormulaAsCellRange();
				Assert.AreEqual("SHeet1!B2:SHEET1!$C5", address.Address);
				Assert.AreEqual(worksheet1, address.Worksheet);

				worksheet1.Names["name1"].NameFormula = "SHeet2!B2:SHEET2!$C5";
				address = worksheet1.Names["name1"].GetFormulaAsCellRange();
				Assert.AreEqual("SHeet2!B2:SHEET2!$C5", address.Address);
				Assert.AreEqual(worksheet2, address.Worksheet);

				worksheet1.Names["name1"].NameFormula = "SHEET1!B2,SHEET1!$C5,SHEET1!$D$6";
				address = worksheet1.Names["name1"].GetFormulaAsCellRange();
				Assert.AreEqual("SHEET1!B2,SHEET1!$C5,SHEET1!$D$6", address.Address);
				Assert.AreEqual(worksheet1, address.Worksheet);
			}
		}

		[TestMethod]
		public void GetFormulaAsAddressNestedNamedRangeTest()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
				worksheet.Names.Add("name1", "Sheet1!G7");
				worksheet.Names.Add("name2", "Sheet1!G6,name1");
				var name2Address = worksheet.Names["name2"].GetFormulaAsCellRange();
				Assert.AreEqual("Sheet1!G6,Sheet1!G7", name2Address?.Address);

				worksheet.Names["name1"].NameFormula = "notanaddress";
				name2Address = worksheet.Names["name2"].GetFormulaAsCellRange();
				Assert.IsNull(name2Address);
			}
		}

		[TestMethod]
		public void GetFormulaAsAddressOffsetFormulaTest()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
				worksheet.Names.Add("name1", "Offset(G7,1,1)");
				var address = worksheet.Names["name1"].GetFormulaAsCellRange();
				Assert.AreEqual(new ExcelRange(worksheet, "'Sheet1'!H8").Address, address.Address);
				Assert.AreEqual(worksheet, address.Worksheet);

				worksheet.Names["name1"].NameFormula = "Sheet1!G7,Offset(G7,1,1)";
				address = worksheet.Names["name1"].GetFormulaAsCellRange();
				Assert.AreEqual("Sheet1!G7,'Sheet1'!H8", address.Address);
				Assert.AreEqual(worksheet, address.Worksheet);

				worksheet.Names["name1"].NameFormula = "Sheet1!G7,Offset(Sheet1!C2,2,1)";
				address = worksheet.Names["name1"].GetFormulaAsCellRange();
				Assert.AreEqual("Sheet1!G7,'Sheet1'!D4", address.Address);

				worksheet.Names["name1"].NameFormula = "Sheet1!G7,Offset(notanaddress,2,1)";
				address = worksheet.Names["name1"].GetFormulaAsCellRange();
				Assert.IsNull(address);
			}
		}

		[TestMethod]
		public void GetFormulaAsAddressIndirectFormulaTest()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
				worksheet.Names.Add("name1", @"Sheet1!G7,Indirect(""Sheet1!C5"")");
				var address = worksheet.Names["name1"].GetFormulaAsCellRange();
				Assert.AreEqual("Sheet1!G7,Sheet1!C5", address.Address);

				worksheet.Names["name1"].NameFormula = @"Sheet1!G7,Indirect(""notanadddress"")";
				address = worksheet.Names["name1"].GetFormulaAsCellRange();
				Assert.IsNull(address);
			}
		}

		[TestMethod]
		public void GetFormulaAsAddressNestedNamedRangeWithFunctionTest()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
				worksheet.Names.Add("name1", "OFFset(Sheet1!C3:Sheet1!D4, 2, 3),Sheet1!F6");
				worksheet.Names.Add("name2", "Sheet1!G6,name1");
				var name2Address = worksheet.Names["name2"].GetFormulaAsCellRange();
				Assert.AreEqual("Sheet1!G6,'Sheet1'!F5:G6,Sheet1!F6", name2Address.Address);

				worksheet.Names["name1"].NameFormula = "OFFset(Sheet1!C3:Sheet1!D4, 2, 3),#REF!F6";
				name2Address = worksheet.Names["name2"].GetFormulaAsCellRange();
				Assert.AreEqual("Sheet1!G6,'Sheet1'!F5:G6,#REF!F6", name2Address.Address);
			}
		}

		[TestMethod]
		[ExpectedException(typeof(InvalidOperationException))]
		public void GetFormulaAsAddressWithoutWorksheetThrowsException()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
				worksheet.Names.Add("name1", "C3:D4");
				worksheet.Names["name1"].GetFormulaAsCellRange();
			}
		}

		[TestMethod]
		[ExpectedException(typeof(InvalidOperationException))]
		public void GetFormulaAsAddressWithInvalidWorksheetThrowsException()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
				worksheet.Names.Add("name1", "Sheet2!C3:Sheet2!D4");
				worksheet.Names["name1"].GetFormulaAsCellRange();
			}
		}
		#endregion

		#region Named Range Formula Calculation Tests
		[TestMethod]
		public void NamedRangeFormulaCalculationTest()
		{
			using (var package = new ExcelPackage())
			{
				var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
				sheet1.Names.Add("name1", "\"7/23/2018\"");
				sheet1.Names.Add("name2", "\"4/15/2017\"");
				sheet1.Cells[2, 2].Formula = "DAYS360(name1, name2)";
				sheet1.Calculate();
				Assert.AreEqual(-458, sheet1.Cells[2, 2].Value);

				sheet1.Names["name1"].NameFormula = "23";
				sheet1.Names["name2"].NameFormula = "102";
				sheet1.Cells[2, 2].Formula = "=name1 + name2";
				sheet1.Calculate();
				Assert.AreEqual(125d, sheet1.Cells[2, 2].Value);

				sheet1.Names["name1"].NameFormula = "{1,2,3,4,5}";
				sheet1.Names["name2"].NameFormula = "{6,7,8}";
				sheet1.Cells[2, 2].Formula = "=SUM(name1, name2)";
				sheet1.Calculate();
				Assert.AreEqual(36d, sheet1.Cells[2, 2].Value);

				sheet1.Names["name1"].NameFormula = "{1,2,3,4,5}";
				sheet1.Names.Add("name3", "14");
				sheet1.Names["name2"].NameFormula = "{6,7,8,name3}";
				sheet1.Cells[2, 2].Formula = "=SUM(name1, name2)";
				sheet1.Calculate();
				Assert.AreEqual(50d, sheet1.Cells[2, 2].Value);

				sheet1.Names["name1"].NameFormula = "{1,2,3,4,5}";
				sheet1.Names["name2"].NameFormula = "OFFSET(Sheet1!$C$3,2,3)";
				sheet1.Cells["F5"].Value = 10;
				sheet1.Cells[2, 2].Formula = "=SUM(name1, name2)";
				sheet1.Calculate();
				Assert.AreEqual(25d, sheet1.Cells[2, 2].Value);

				sheet1.Names["name1"].NameFormula = "{1,2,3,4,5}";
				sheet1.Names["name3"].NameFormula = "Sheet1!$C$3";
				sheet1.Names["name2"].NameFormula = "OFFSET(name3,2,3)";
				sheet1.Cells["F5"].Value = 10;
				sheet1.Cells["C3"].Value = 5;
				sheet1.Cells[2, 2].Formula = "=SUM(name1, name2, name3)";
				sheet1.Calculate();
				Assert.AreEqual(30d, sheet1.Cells[2, 2].Value);
			}
		}
		#endregion
	}
}