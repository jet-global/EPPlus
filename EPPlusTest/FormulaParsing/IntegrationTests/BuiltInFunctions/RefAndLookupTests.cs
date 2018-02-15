using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Exceptions;
using Rhino.Mocks;

namespace EPPlusTest.FormulaParsing.IntegrationTests.BuiltInFunctions
{
	[TestClass]
	public class RefAndLookupTests : FormulaParserTestBase
	{
		#region Class Variables
		private ExcelDataProvider _excelDataProvider;
		const string WorksheetName = null;
		private ExcelPackage _package;
		private ExcelWorksheet _worksheet;
		#endregion

		#region Test Setup
		[TestInitialize]
		public void Initialize()
		{
			_excelDataProvider = MockRepository.GenerateStub<ExcelDataProvider>();
			_excelDataProvider.Stub(x => x.GetDimensionEnd(Arg<string>.Is.Anything)).Return(new ExcelCellAddress(10, 1));
			_parser = new FormulaParser(_excelDataProvider);
			_package = new ExcelPackage();
			_worksheet = _package.Workbook.Worksheets.Add("Test");
		}

		[TestCleanup]
		public void Cleanup()
		{
			_package.Dispose();
		}
		#endregion

		#region VLookup Tests
		[TestMethod]
		public void VLookupShouldReturnCorrespondingValue()
		{
			using (var pck = new ExcelPackage())
			{
				var ws = pck.Workbook.Worksheets.Add("test");
				ws.Cells["A1"].Value = 1;
				ws.Cells["B1"].Value = 1;
				ws.Cells["A2"].Value = 2;
				ws.Cells["B2"].Value = 5;
				ws.Cells["A3"].Formula = "VLOOKUP(2, A1:B2, 2)";
				ws.Calculate();
				var result = ws.Cells["A3"].Value;
				Assert.AreEqual(5, result);
			}
		}

		[TestMethod]
		public void VLookupWithTextAndFalseRangeLookupShouldMatchWildcard()
		{
			using (var pck = new ExcelPackage())
			{
				var ws = pck.Workbook.Worksheets.Add("test");
				ws.Cells["A1"].Value = "apples";
				ws.Cells["B1"].Value = "bananas";
				ws.Cells["A2"].Value = "other";
				ws.Cells["B2"].Value = "stuff";
				ws.Cells["A3"].Formula = @"VLOOKUP(""oth*?"", A1:B2, 2, FALSE)";
				ws.Calculate();
				var result = ws.Cells["A3"].Value;
				Assert.AreEqual("stuff", result);
			}
		}

		[TestMethod]
		public void VLookupWithTextAndTrueRangeLookupShouldReturnFirstClosestMatch()
		{
			using (var pck = new ExcelPackage())
			{
				var ws = pck.Workbook.Worksheets.Add("test");
				ws.Cells["A1"].Value = "apples";
				ws.Cells["B1"].Value = "bananas";
				ws.Cells["A2"].Value = "other";
				ws.Cells["B2"].Value = "stuff";
				ws.Cells["A3"].Formula = @"VLOOKUP(""oth*?"", A1:B2, 2, TRUE)";
				ws.Calculate();
				var result = ws.Cells["A3"].Value;
				Assert.AreEqual("bananas", result);
			}
		}

		[TestMethod]
		public void VLookupWithAllReferenceArguments()
		{
			using (var pck = new ExcelPackage())
			{
				var ws = pck.Workbook.Worksheets.Add("test");
				ws.Cells["A1"].Value = 1;
				ws.Cells["B1"].Value = 1;
				ws.Cells["A2"].Value = 2;
				ws.Cells["B2"].Value = 5;

				ws.Cells["Z5"].Value = 2;
				ws.Cells["Z6"].Value = 2;
				ws.Cells["A3"].Formula = "VLOOKUP(Z5, A1:B2, Z6)";
				ws.Calculate();
				var result = ws.Cells["A3"].Value;
				Assert.AreEqual(5, result);
			}
		}

		[TestMethod, Ignore]
		/// <remarks>This currently does not work because if does not return a reference.</remarks>
		public void VLookupOfIf()
		{
			using (var pck = new ExcelPackage())
			{
				var ws = pck.Workbook.Worksheets.Add("test");
				ws.Cells["B1"].Value = 1;
				ws.Cells["C1"].Value = 1;
				ws.Cells["B2"].Value = 2;
				ws.Cells["C2"].Value = 5;
				ws.Cells["A3"].Formula = "VLOOKUP(2, IF(1,A1:B2,Z10), 2)";
				ws.Calculate();
				var result = ws.Cells["A3"].Value;
				Assert.AreEqual(5, result);
			}
		}

		[TestMethod]
		public void VLookupShouldReturnClosestValueBelowIfLastArgIsTrue()
		{
			using (var pck = new ExcelPackage())
			{
				var ws = pck.Workbook.Worksheets.Add("test");
				var lookupAddress = "A1:B2";
				ws.Cells["A1"].Value = 3;
				ws.Cells["B1"].Value = 1;
				ws.Cells["A2"].Value = 5;
				ws.Cells["B2"].Value = 5;
				ws.Cells["A3"].Formula = "VLOOKUP(4, " + lookupAddress + ", 2, true)";
				ws.Calculate();
				var result = ws.Cells["A3"].Value;
				Assert.AreEqual(1, result);
			}
		}
		#endregion

		#region HLookup Tests
		[TestMethod]
		public void HLookupShouldReturnCorrespondingValue()
		{
			var lookupAddress = "A1:B2";
			_worksheet.Cells["A1"].Value = 1;
			_worksheet.Cells["B1"].Value = 2;
			_worksheet.Cells["A2"].Value = 2;
			_worksheet.Cells["B2"].Value = 5;
			_worksheet.Cells["A3"].Formula = "HLOOKUP(2, " + lookupAddress + ", 2)";
			_worksheet.Calculate();
			var result = _worksheet.Cells["A3"].Value;
			Assert.AreEqual(5, result);
		}

		[TestMethod]
		public void HLookupWithTextAndFalseRangeLookupShouldMatchWildcard()
		{
			using (var pck = new ExcelPackage())
			{
				var ws = pck.Workbook.Worksheets.Add("test");
				ws.Cells["A1"].Value = "apples";
				ws.Cells["B1"].Value = "other";
				ws.Cells["A2"].Value = "bananas";
				ws.Cells["B2"].Value = "stuff";
				ws.Cells["A3"].Formula = @"HLOOKUP(""oth*?"", A1:B2, 2, FALSE)";
				ws.Calculate();
				var result = ws.Cells["A3"].Value;
				Assert.AreEqual("stuff", result);
			}
		}

		[TestMethod]
		public void HLookupWithTextAndTrueRangeLookupShouldReturnFirstClosestMatch()
		{
			using (var pck = new ExcelPackage())
			{
				var ws = pck.Workbook.Worksheets.Add("test");
				ws.Cells["A1"].Value = "apples";
				ws.Cells["B1"].Value = "other";
				ws.Cells["A2"].Value = "bananas";
				ws.Cells["B2"].Value = "stuff";
				ws.Cells["A3"].Formula = @"HLOOKUP(""oth*?"", A1:B2, 2, TRUE)";
				ws.Calculate();
				var result = ws.Cells["A3"].Value;
				Assert.AreEqual("bananas", result);
			}
		}

		[TestMethod]
		public void HLookupShouldReturnClosestValueBelowIfLastArgIsTrue()
		{
			_worksheet.Cells["A1"].Value = 3;
			_worksheet.Cells["B1"].Value = 5;
			_worksheet.Cells["A2"].Value = 1;
			_worksheet.Cells["B2"].Value = 2;
			_worksheet.Cells["A3"].Formula = "HLOOKUP(4, A1:B2, 2, TRUE)";
			_worksheet.Calculate();
			var result = _worksheet.Cells["A3"].Value;
			Assert.AreEqual(1, result);
		}

		[TestMethod]
		public void HLookupWithAllReferenceArguments()
		{
			using (var pck = new ExcelPackage())
			{
				var ws = pck.Workbook.Worksheets.Add("test");
				ws.Cells["A1"].Value = 1;
				ws.Cells["B1"].Value = 3;
				ws.Cells["A2"].Value = 2;
				ws.Cells["B2"].Value = 5;

				ws.Cells["Z5"].Value = 3;
				ws.Cells["Z6"].Value = 2;
				ws.Cells["A3"].Formula = "HLOOKUP(Z5, A1:B2, Z6)";
				ws.Calculate();
				var result = ws.Cells["A3"].Value;
				Assert.AreEqual(5, result);
			}
		}
		#endregion

		#region Lookup Tests
		[TestMethod]
		public void LookupShouldReturnMatchingValue()
		{
			_worksheet.Cells["A1"].Value = 3;
			_worksheet.Cells["B1"].Value = 5;
			_worksheet.Cells["A2"].Value = 4;
			_worksheet.Cells["B2"].Value = 1;
			_worksheet.Cells["A3"].Formula = "LOOKUP(4, A1:B2)";
			_worksheet.Calculate();
			var result = _worksheet.Cells["A3"].Value;
			Assert.AreEqual(1, result);
		}
		#endregion

		#region Row Tests
		[TestMethod]
		public void RowShouldReturnRowNumber()
		{
			_excelDataProvider.Stub(x => x.GetRangeFormula("", 4, 1)).Return("Row()");
			var result = _parser.ParseAt("A4");
			Assert.AreEqual(4, result);
		}

		[TestMethod]
		public void RowSholdHandleReference()
		{
			using (var package = new ExcelPackage())
			{
				var s1 = package.Workbook.Worksheets.Add("test");
				s1.Cells["A1"].Formula = "ROW(A4)";
				s1.Calculate();
				Assert.AreEqual(4, s1.Cells["A1"].Value);
			}
		}

		[TestMethod]
		public void RowOfRangeReturnsFirstRowNumber()
		{
			using (var package = new ExcelPackage())
			{
				var s1 = package.Workbook.Worksheets.Add("test");
				s1.Cells["A1"].Formula = "ROW(C5:D9)";
				s1.Calculate();
				Assert.AreEqual(5, s1.Cells["A1"].Value);
			}
		}

		[TestMethod]
		public void RowOfOffsetOfNestedFunction()
		{
			using (var package = new ExcelPackage())
			{
				var s1 = package.Workbook.Worksheets.Add("test");
				s1.Cells["A1"].Formula = "ROW(OFFSET(A1,IF(1,1,0),IF(1,2,0)))";
				s1.Calculate();
				Assert.AreEqual(2, s1.Cells["A1"].Value);
			}
		}
		#endregion

		#region Column Tests
		[TestMethod]
		public void ColumnShouldReturnRowNumber()
		{
			_excelDataProvider.Stub(x => x.GetRangeFormula("", 4, 2)).Return("Column()");
			var result = _parser.ParseAt("B4");
			Assert.AreEqual(2, result);
		}

		[TestMethod]
		public void ColumnSholdHandleReference()
		{
			using (var package = new ExcelPackage())
			{
				var s1 = package.Workbook.Worksheets.Add("test");
				s1.Cells["A1"].Formula = "COLUMN(B4)";
				s1.Calculate();
				Assert.AreEqual(2, s1.Cells["A1"].Value);
			}
		}

		[TestMethod]
		public void ColumnOfRangeReturnsFirstColumnNumber()
		{
			using (var package = new ExcelPackage())
			{
				var s1 = package.Workbook.Worksheets.Add("test");
				s1.Cells["A1"].Formula = "COLUMN(C5:D9)";
				s1.Calculate();
				Assert.AreEqual(3, s1.Cells["A1"].Value);
			}
		}

		[TestMethod]
		public void ColumnOfOffsetOfNestedFunction()
		{
			using (var package = new ExcelPackage())
			{
				var s1 = package.Workbook.Worksheets.Add("test");
				s1.Cells["A1"].Formula = "COLUMN(OFFSET(A1,IF(1,1,0),IF(1,2,0)))";
				s1.Calculate();
				Assert.AreEqual(3, s1.Cells["A1"].Value);
			}
		}
		#endregion

		#region Rows Test
		[TestMethod]
		public void RowsShouldReturnNbrOfRows()
		{
			using (var package = new ExcelPackage())
			{
				var s1 = package.Workbook.Worksheets.Add("test");
				s1.Cells["A4"].Formula = "Rows(A5:B7)";
				s1.Calculate();
				Assert.AreEqual(3, s1.Cells["A4"].Value);
			}
		}
		#endregion

		#region Columns Test
		[TestMethod]
		public void ColumnsShouldReturnNbrOfCols()
		{
			using (var package = new ExcelPackage())
			{
				var s1 = package.Workbook.Worksheets.Add("test");
				s1.Cells["A4"].Formula = "Columns(A5:B7)";
				s1.Calculate();
				Assert.AreEqual(2, s1.Cells["A4"].Value);
			}
		}
		#endregion

		#region Choose Tests
		[TestMethod]
		public void ChooseShouldReturnCorrectResult()
		{
			var result = _parser.Parse("Choose(1, \"A\", \"B\")");
			Assert.AreEqual("A", result);
		}
		#endregion

		#region Offset Tests
		[TestMethod]
		public void OffsetShouldReturnASingleValue()
		{
			using (var package = new ExcelPackage())
			{
				var s1 = package.Workbook.Worksheets.Add("Test");
				s1.Cells["B3"].Value = 1d;
				s1.Cells["A5"].Formula = "OFFSET(A1, 2, 1)";
				s1.Calculate();
				Assert.AreEqual(1d, s1.Cells["A5"].Value);
			}
		}

		[TestMethod]
		public void OffsetShouldReturnARange()
		{
			using (var package = new ExcelPackage())
			{
				var s1 = package.Workbook.Worksheets.Add("Test");
				s1.Cells["B1"].Value = 1d;
				s1.Cells["B2"].Value = 1d;
				s1.Cells["B3"].Value = 1d;
				s1.Cells["A5"].Formula = "SUM(OFFSET(A1:A3, 0, 1))";
				s1.Calculate();
				Assert.AreEqual(3d, s1.Cells["A5"].Value);
			}
		}

		[TestMethod]
		public void OffsetHandlesReferencesForAllArguments()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells[1, 1].Formula = "=OFFSET(A1,A2,A3)";
				worksheet.Cells[2, 1].Value = 2;
				worksheet.Cells[3, 1].Value = 2;
				worksheet.Cells[3, 3].Value = 5;
				worksheet.Calculate();
				Assert.AreEqual(5, worksheet.Cells[1, 1].Value);
			}
		}

		[TestMethod]
		public void OffsetDirectReferenceToMultiRangeShouldSetValueError()
		{
			using (var package = new ExcelPackage())
			{
				var s1 = package.Workbook.Worksheets.Add("Test");
				s1.Cells["B1"].Value = 1d;
				s1.Cells["B2"].Value = 1d;
				s1.Cells["B3"].Value = 1d;
				s1.Cells["A5"].Formula = "OFFSET(A1:A3, 0, 1)";
				s1.Calculate();
				var result = s1.Cells["A5"].Value;
				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), result);
			}
		}

		[TestMethod]
		public void OffsetShouldReturnARangeAccordingToWidth()
		{
			using (var package = new ExcelPackage())
			{
				var s1 = package.Workbook.Worksheets.Add("Test");
				s1.Cells["B1"].Value = 1d;
				s1.Cells["B2"].Value = 1d;
				s1.Cells["B3"].Value = 1d;
				s1.Cells["A5"].Formula = "SUM(OFFSET(A1:A3, 0, 1, 2))";
				s1.Calculate();
				Assert.AreEqual(2d, s1.Cells["A5"].Value);
			}
		}

		[TestMethod]
		public void OffsetShouldReturnARangeAccordingToHeight()
		{
			using (var package = new ExcelPackage())
			{
				var s1 = package.Workbook.Worksheets.Add("Test");
				s1.Cells["B1"].Value = 1d;
				s1.Cells["B2"].Value = 1d;
				s1.Cells["B3"].Value = 1d;
				s1.Cells["C1"].Value = 2d;
				s1.Cells["C2"].Value = 2d;
				s1.Cells["C3"].Value = 2d;
				s1.Cells["A5"].Formula = "SUM(OFFSET(A1:A3, 0, 1, 2, 2))";
				s1.Calculate();
				Assert.AreEqual(6d, s1.Cells["A5"].Value);
			}
		}

		[TestMethod]
		public void OffsetShouldCoverMultipleColumns()
		{
			using (var package = new ExcelPackage())
			{
				var s1 = package.Workbook.Worksheets.Add("Test");
				s1.Cells["C1"].Value = 1d;
				s1.Cells["C2"].Value = 1d;
				s1.Cells["C3"].Value = 1d;
				s1.Cells["D1"].Value = 2d;
				s1.Cells["D2"].Value = 2d;
				s1.Cells["D3"].Value = 2d;
				s1.Cells["A5"].Formula = "SUM(OFFSET(A1:B3, 0, 2))";
				s1.Calculate();
				Assert.AreEqual(9d, s1.Cells["A5"].Value);
			}
		}

		[TestMethod]
		public void OffsetShouldGetValueOnCorrectSheet()
		{
			using (var package = new ExcelPackage())
			{
				var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
				var sheet2 = package.Workbook.Worksheets.Add("Sheet2");
				sheet1.Cells["E5"].Value = "Bad";
				sheet2.Cells["E5"].Value = "Good";
				sheet2.Cells["C3"].Formula = "OFFSET(D4, 1, 1)";
				sheet2.Calculate();
				Assert.AreEqual("Good", sheet2.Cells["C3"].Value);
			}
		}

		[TestMethod]
		public void OffsetShouldGetCalculatedValue()
		{
			using (var package = new ExcelPackage())
			{
				var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
				sheet1.Cells["E5"].Formula = "CONCATENATE(\"Y\",\"o\")";
				sheet1.Cells["C3"].Formula = "OFFSET(D4, 1, 1)";
				sheet1.Cells["C3"].Calculate();
				Assert.AreEqual("Yo", sheet1.Cells["C3"].Value);
			}
		}

		[TestMethod]
		public void OffsetWithSemiCircularReferenceGetsValue()
		{
			using (var package = new ExcelPackage())
			{
				var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
				sheet1.Cells["D3"].Formula = "IF(TRUE, 0, E3)";
				sheet1.Cells["D4"].Value = "Good";
				sheet1.Cells["E3"].Formula = "OFFSET(D3, 1,0)";
				sheet1.Cells["E3"].Calculate();
				Assert.AreEqual("Good", sheet1.Cells["E3"].Value);
			}
		}

		[TestMethod]
		public void OffsetWithSimpleSemiCircularReferenceThrowsException()
		{
			using (var package = new ExcelPackage())
			{
				var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
				sheet1.Cells["E3"].Formula = "offset(E3, 1,0)";
				sheet1.Cells["E4"].Value = "Good";
				sheet1.Cells["E3"].Calculate();
				Assert.AreEqual("Good", sheet1.Cells["E3"].Value);
			}
		}

		[TestMethod]
		public void OffsetWithSemiCircularChainedReferenceGetsValue()
		{
			using (var package = new ExcelPackage())
			{
				var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
				sheet1.Cells["D3"].Formula = "IF(TRUE, 0, E3)";
				sheet1.Cells["E3"].Formula = "OFFSET(D3, 1,0)";
				sheet1.Cells["D4"].Formula = @"INDIRECT(""E4"")";
				sheet1.Cells["E3"].Value = "Good";
				sheet1.Cells["E3"].Calculate();
				Assert.AreEqual("Good", sheet1.Cells["E3"].Value);
			}
		}

		[TestMethod]
		public void OffsetsWithSemiCircularChainedReferenceGetsValue()
		{
			using (var package = new ExcelPackage())
			{
				var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
				sheet1.Cells["D3"].Formula = "SUM(OFFSET(C3, 1, 0), OFFSET(E3, 1, 0))";
				sheet1.Cells["C3"].Formula = "=D3";
				sheet1.Cells["C4"].Value = 10d;
				sheet1.Cells["E3"].Formula = "=D3";
				sheet1.Cells["E4"].Value = 5d;
				sheet1.Cells["D3"].Calculate();
				Assert.AreEqual(15d, sheet1.Cells["D3"].Value);
			}
		}

		[TestMethod]
		public void OffsetWithNestedOffsetRowValueTest()
		{
			using (var package = new ExcelPackage())
			{
				var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
				sheet1.Cells["G8"].Formula = "OFFSET(A1, OFFSET(C3, 1, 1), 3)";
				sheet1.Cells["D4"].Value = 1;
				sheet1.Cells["D2"].Value = "Success!";
				sheet1.Cells["G8"].Calculate();
				Assert.AreEqual("Success!", sheet1.Cells["G8"].Value);
			}
		}

		[TestMethod]
		[ExpectedException(typeof(CircularReferenceException))]
		public void OffsetWithSimpleCircularReferenceThrowsException()
		{
			using (var package = new ExcelPackage())
			{
				var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
				sheet1.Cells["E3"].Formula = "offset(F3, 1,0)";
				sheet1.Cells["F4"].Formula = "E3";
				sheet1.Cells["E3"].Calculate();
			}
		}

		[TestMethod]
		[ExpectedException(typeof(CircularReferenceException))]
		public void OffsetWithCircularReferenceThrowsException()
		{
			using (var package = new ExcelPackage())
			{
				var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
				sheet1.Cells["E3"].Formula = "offset(F3, 1,0)";
				sheet1.Cells["F4"].Formula = "G4";
				sheet1.Cells["G4"].Formula = "IF(TRUE, D3, C3)";
				sheet1.Cells["D3"].Formula = "E3";
				sheet1.Cells["E3"].Calculate();
			}
		}

		[TestMethod]
		[ExpectedException(typeof(CircularReferenceException))]
		public void OffsetWithCrossSheetCircularReferenceThrowsException()
		{
			using (var package = new ExcelPackage())
			{
				var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
				var sheet2 = package.Workbook.Worksheets.Add("Sheet2");
				sheet1.Cells["E3"].Formula = "offset(F3, 1,0)";
				sheet1.Cells["F4"].Formula = "G4";
				sheet1.Cells["G4"].Formula = "IF(TRUE, D3, C3)";
				sheet1.Cells["D3"].Formula = "Sheet2!E3";
				sheet2.Cells["E3"].Formula = "Sheet1!E3";
				sheet1.Cells["E3"].Calculate();
			}
		}
		#endregion

		#region Miscellaneous Tests
		[TestMethod]
		public void AddressShouldReturnCorrectResult()
		{
			_excelDataProvider.Stub(x => x.ExcelMaxRows).Return(12345);
			var result = _parser.Parse("Address(1, 1)");
			Assert.AreEqual("$A$1", result);
		}

		[TestMethod]
		public void IndirectShouldReturnARange()
		{
			using (var package = new ExcelPackage(new MemoryStream()))
			{
				var s1 = package.Workbook.Worksheets.Add("Test");
				s1.Cells["A1:A2"].Value = 2;
				s1.Cells["A3"].Formula = "SUM(Indirect(\"A1:A2\"))";
				s1.Calculate();
				Assert.AreEqual(4d, s1.Cells["A3"].Value);

				s1.Cells["A4"].Formula = "SUM(Indirect(\"A1:A\" & \"2\"))";
				s1.Calculate();
				Assert.AreEqual(4d, s1.Cells["A4"].Value);
			}
		}

		[TestMethod]
		public void MatchShouldReturnIndexOfMatchingValue()
		{
			_worksheet.Cells["A1"].Value = 3;
			_worksheet.Cells["A2"].Value = 3;
			_worksheet.Cells["A3"].Formula = "MATCH(3, A1:A2)";
			_worksheet.Calculate();
			var result = _worksheet.Cells["A3"].Value;
			Assert.AreEqual(1, result);
		}

		[TestMethod, Ignore]
		public void VLookupShouldHandleNames()
		{
			using (var package = new ExcelPackage(new FileInfo(@"c:\temp\Book3.xlsx")))
			{
				var s1 = package.Workbook.Worksheets.First();
				var v = s1.Cells["X10"].Formula;
				//s1.Calculate();
				v = s1.Cells["X10"].Formula;
			}
		}
		#endregion
	}
}
