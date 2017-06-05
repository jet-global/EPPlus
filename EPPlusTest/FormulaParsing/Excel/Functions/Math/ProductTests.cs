using System;
using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math
{
	[TestClass]
	public class ProductTests : MathFunctionsTestBase
	{
		#region Product Function (Execute) Tests
		[TestMethod]
		public void ProductWithIntegerInputReturnsCorrectValue()
		{
			var function = new Product();
			var result = function.Execute(FunctionsHelper.CreateArgs(5, 8), this.ParsingContext);
			Assert.AreEqual(40d, result.Result);
		}

		[TestMethod]
		public void ProductWithZeroReturnsCorrectValue()
		{
			var function = new Product();
			var result = function.Execute(FunctionsHelper.CreateArgs(88, 0), this.ParsingContext);
			Assert.AreEqual(0d, result.Result);
		}

		[TestMethod]
		public void ProductWithDoublesReturnsCorrectValue()
		{
			var function = new Product();
			var result = function.Execute(FunctionsHelper.CreateArgs(5.5, 3.8), this.ParsingContext);
			Assert.AreEqual(20.9, result.Result);
		}

		[TestMethod]
		public void ProductWithTwoNegativeIntegersReturnsCorrectValue()
		{
			var function = new Product();
			var result = function.Execute(FunctionsHelper.CreateArgs(-9, -52), this.ParsingContext);
			Assert.AreEqual(468d, result.Result);
		}

		[TestMethod]
		public void ProductWithOneNegativeIntegerOnePositiveIntegerReturnsCorrectValue()
		{
			var function = new Product();
			var result = function.Execute(FunctionsHelper.CreateArgs(-15, 2), this.ParsingContext);
			Assert.AreEqual(-30, result.Result);
		}

		[TestMethod]
		public void ProductWithFractionInputsReturnsCorrectValue()
		{
			var function = new Product();
			var result = function.Execute(FunctionsHelper.CreateArgs((2 / 3), (1 / 5)), this.ParsingContext);
			Assert.AreEqual(0.13333333, result.Result);
		}

		[TestMethod]
		public void ProductWithDatesAsResultOfDateFunctionReturnsCorrectValue()
		{
			var function = new Product();
			var dateInput = new DateTime(2017, 5, 1);
			var result = function.Execute(FunctionsHelper.CreateArgs(dateInput, 1), this.ParsingContext);
			Assert.AreEqual(42856d, result.Result);
		}

		[TestMethod]
		public void ProductWithDatesAsStringsReturnsCorrectValue()
		{
			var function = new Product();
			var result = function.Execute(FunctionsHelper.CreateArgs("5/5/2017", 2), this.ParsingContext);
			Assert.AreEqual(85720d, result.Result);
		}

		[TestMethod]
		public void ProductWithDateNotAsStringReturnsCorrectValue()
		{
			var function = new Product();
			var result = function.Execute(FunctionsHelper.CreateArgs(5 / 5 / 2017, 2), this.ParsingContext);
			Assert.AreEqual(0.000991572, result.Result);
		}

		[TestMethod]
		public void ProductWithGeneralAndEmptyStringReturnsPoundValue()
		{
			var function = new Product();
			var result = function.Execute(FunctionsHelper.CreateArgs("string", ""), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void ProductWithOneArgumentReturnsCorrectValue()
		{
			var function = new Product();
			var result = function.Execute(FunctionsHelper.CreateArgs(5), this.ParsingContext);
			Assert.AreEqual(5d, result.Result);
		}

		[TestMethod]
		public void ProductWithNullSecondArgumentReturnsCorrectValue()
		{
			var function = new Product();
			var result = function.Execute(FunctionsHelper.CreateArgs(5, null), this.ParsingContext);
			Assert.AreEqual(5d, result.Result);
		}

		[TestMethod]
		public void ProductWithNullFirstArgumentReturnsCorrectValue()
		{
			var function = new Product();
			var result = function.Execute(FunctionsHelper.CreateArgs(null, 5), this.ParsingContext);
			Assert.AreEqual(5d, result.Result);
		}

		[TestMethod]
		public void ProductWithNoArgumentsReturnsPoundValue()
		{
			var function = new Product();
			var result = function.Execute(FunctionsHelper.CreateArgs(), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void ProductWithNumbersAsStringsReturnsCorrectValue()
		{
			var function = new Product();
			var result = function.Execute(FunctionsHelper.CreateArgs("5", "8"), this.ParsingContext);
			Assert.AreEqual(40d, result.Result);
		}

		[TestMethod]
		public void ProductWithOneInputAsExcelRangeReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["B1"].Value = 5;
				ws.Cells["B2"].Value = 8;
				ws.Cells["B3"].Formula = "PRODUCT(B1:B2)";
				ws.Calculate();
				Assert.AreEqual(40d, ws.Cells["B3"].Value);
			}
		}

		[TestMethod]
		public void ProductWithTwoExcelRangesAsInputsReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["B1"].Value = 5;
				ws.Cells["B2"].Value = 8;
				ws.Cells["C1"].Value = 10;
				ws.Cells["C2"].Value = 5;
				ws.Cells["B3"].Formula = "PRODUCT(B1:B2, C1:C2)";
				ws.Calculate();
				Assert.AreEqual(2000d, ws.Cells["B3"].Value);
			}
		}
		#endregion
	}
}
