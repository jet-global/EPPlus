using System;
using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math
{
	[TestClass]
	public class CountTests : MathFunctionsTestBase
	{
		#region Count Tests
		[TestMethod]
		public void CountWithSingleNumber()
		{
			var function = new Count();
			var arguments = FunctionsHelper.CreateArgs(2);
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(1d, result.Result);
		}

		[TestMethod]
		public void CountWithNumericString()
		{
			var function = new Count();
			var arguments = FunctionsHelper.CreateArgs("2");
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(1d, result.Result);
		}

		[TestMethod]
		public void CountWithNonNumericString()
		{
			var function = new Count();
			var arguments = FunctionsHelper.CreateArgs("word");
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(0d, result.Result);
		}

		[TestMethod]
		public void CountWithDateInString()
		{
			var function = new Count();
			var arguments = FunctionsHelper.CreateArgs("7/5/2017");
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(1d, result.Result);
		}

		[TestMethod]
		public void CountWithBoolean()
		{
			var function = new Count();
			var arguments = FunctionsHelper.CreateArgs(true);
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(1d, result.Result);
		}

		[TestMethod]
		public void CountWithBooleanInString()
		{
			var function = new Count();
			var arguments = FunctionsHelper.CreateArgs("TRUE");
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(0d, result.Result);
		}

		[TestMethod]
		public void CountWithEmptyString()
		{
			var function = new Count();
			var arguments = FunctionsHelper.CreateArgs(string.Empty);
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(0d, result.Result);
		}

		[TestMethod]
		public void CountWithValuesInSingleCellReferences()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B2"].Formula = "COUNT(C2)";
				worksheet.Cells["C2"].Value = 2;
				worksheet.Cells["B3"].Formula = "COUNT(C3)";
				worksheet.Cells["C3"].Value = "2";
				worksheet.Cells["B4"].Formula = "COUNT(C4)";
				worksheet.Cells["C4"].Value = "word";
				worksheet.Cells["B5"].Formula = "COUNT(C5)";
				worksheet.Cells["C5"].Value = (new DateTime(2017, 7, 5)).ToOADate();
				worksheet.Cells["B6"].Formula = "COUNT(C6)";
				worksheet.Cells["C6"].Value = "7/5/2017";
				worksheet.Cells["B7"].Formula = "COUNT(C7)";
				worksheet.Cells["C7"].Value = true;
				worksheet.Cells["B8"].Formula = "COUNT(C8)";
				worksheet.Cells["C8"].Value = "TRUE";
				worksheet.Cells["B9"].Formula = "COUNT(C9)";
				worksheet.Cells["C9"].Value = string.Empty;
				worksheet.Cells["B10"].Formula = "COUNT(C10)";
				worksheet.Cells["C10"].Value = null;
				worksheet.Calculate();
				Assert.AreEqual(1d, worksheet.Cells["B2"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B3"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B4"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B5"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B6"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B7"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B8"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B9"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B10"].Value);
			}
		}

		[TestMethod]
		public void CountWithValuesInMultiCellRanges()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B2"].Formula = "COUNT(C2:D2)";
				worksheet.Cells["C2"].Value = 3;
				worksheet.Cells["D2"].Value = 2.5;
				worksheet.Cells["B3"].Formula = "COUNT(C3:D3)";
				worksheet.Cells["C3"].Value = 3;
				worksheet.Cells["D3"].Value = "2.5";
				worksheet.Cells["B4"].Formula = "COUNT(C4:D4)";
				worksheet.Cells["C4"].Value = 3;
				worksheet.Cells["D4"].Value = "word";
				worksheet.Cells["B5"].Formula = "COUNT(C5:D5)";
				worksheet.Cells["C5"].Value = 3;
				worksheet.Cells["D5"].Value = true;
				worksheet.Cells["B6"].Formula = "COUNT(C6:D6)";
				worksheet.Cells["C6"].Value = 3;
				worksheet.Cells["D6"].Value = "TRUE";
				worksheet.Cells["B7"].Formula = "COUNT(C7:D7)";
				worksheet.Cells["C7"].Value = 3;
				worksheet.Cells["D7"].Value = (new DateTime(2017, 7, 5)).ToOADate();
				worksheet.Cells["B8"].Formula = "COUNT(C8:D8)";
				worksheet.Cells["C8"].Value = 3;
				worksheet.Cells["D8"].Value = "7/5/2017";
				worksheet.Cells["B9"].Formula = "COUNT(C9:D9)";
				worksheet.Cells["C9"].Value = 3;
				worksheet.Cells["D9"].Value = string.Empty;
				worksheet.Cells["B10"].Formula = "COUNT(C10:D10)";
				worksheet.Cells["C10"].Value = 3;
				worksheet.Cells["D10"].Value = null;
				worksheet.Cells["B11"].Formula = "COUNT(C11:D11)";
				worksheet.Cells["C11"].Formula = "notAValidFormula"; // Evaluates to #NAME.
				worksheet.Cells["D11"].Formula = "1/0"; // Evalueates to #DIV/0
				worksheet.Calculate();
				Assert.AreEqual(2d, worksheet.Cells["B2"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B3"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B4"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B5"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B6"].Value);
				Assert.AreEqual(2d, worksheet.Cells["B7"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B8"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B9"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B10"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B11"].Value);
			}
		}

		[TestMethod]
		public void CountWithValuesInArrays()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B2"].Formula = "COUNT({1,2,3})";
				worksheet.Cells["B3"].Formula = "COUNT({\"1\",\"2\",\"3\"})";
				worksheet.Cells["B4"].Formula = "COUNT(C4)";
				worksheet.Cells["C4"].Formula = "{1,2,3}";
				worksheet.Cells["B5"].Formula = "COUNT(C5)";
				worksheet.Cells["C5"].Formula = "{\"1\",\"2\",\"3\"}";
				worksheet.Calculate();
				Assert.AreEqual(3d, worksheet.Cells["B2"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B3"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B4"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void CountWithValuesAsFormulas()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B2"].Formula = "COUNT(C2)";
				worksheet.Cells["C2"].Value = 3;
				worksheet.Cells["B3"].Formula = "COUNT(C3)";
				worksheet.Cells["C3"].Formula = "3";
				worksheet.Cells["B4"].Formula = "COUNT(C4)";
				worksheet.Cells["C4"].Value = "3";
				worksheet.Cells["B5"].Formula = "COUNT(C5)";
				worksheet.Cells["C5"].Formula = "\"3\"";
				worksheet.Cells["B6"].Formula = "COUNT(C6)";
				worksheet.Cells["C6"].Value = true;
				worksheet.Cells["B7"].Formula = "COUNT(C7)";
				worksheet.Cells["C7"].Formula = "TRUE";
				worksheet.Cells["B8"].Formula = "COUNT(C8)";
				worksheet.Cells["C8"].Value = "7/5/2017";
				worksheet.Cells["B9"].Formula = "COUNT(C9)";
				worksheet.Cells["C9"].Formula = "\"7/5/2017\"";
				worksheet.Calculate();
				Assert.AreEqual(1d, worksheet.Cells["B2"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B3"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B4"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B5"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B6"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B7"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B8"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B9"].Value);
			}
		}
		#endregion
	}
}
