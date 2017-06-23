using System;
using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math
{
	/// <summary>
	/// Tests for the FLOOR function, as defined at 
	/// https://support.office.com/en-us/article/FLOOR-function-14bb497c-24f2-4e04-b327-b0b4de5a8886
	/// and https://exceljet.net/excel-functions/excel-floor-function .
	/// </summary>
	[TestClass]
	public class FloorTests : MathFunctionsTestBase
	{
		private ParsingContext _parsingContext = ParsingContext.Create();

		[TestMethod]
		public void FloorOfNineIntoThreeIsNine()
		{
			var func = new Floor();
			var args = FunctionsHelper.CreateArgs(9, 3);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(9.0, result.Result);
		}

		[TestMethod]
		public void FloorWorksForDecimalsToo()
		{
			var func = new Floor();
			var args = FunctionsHelper.CreateArgs(1.57, 0.1);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(1.5, result.Result);
			args = FunctionsHelper.CreateArgs(0.234, 0.01);
			result = func.Execute(args, _parsingContext);
			Assert.AreEqual(0.23, result.Result);
		}

		[TestMethod]
		public void FloorWorksForWeirdDecimalsToo()
		{
			var func = new Floor();
			var args = FunctionsHelper.CreateArgs(1.57, 0.397);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(1.191, result.Result);
		}

		[TestMethod]
		public void FloorOfNegativeOneIntoThreeIsNegativeThree()
		{
			var func = new Floor();
			var args = FunctionsHelper.CreateArgs(-1, 3);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(-3.0, result.Result);
		}

		[TestMethod]
		public void FloorOfTenAndElevenIntoThreeIsNine()
		{
			var func = new Floor();
			var args = FunctionsHelper.CreateArgs(10, 3);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(9.0, result.Result);
			args = FunctionsHelper.CreateArgs(11, 3);
			result = func.Execute(args, _parsingContext);
			Assert.AreEqual(9.0, result.Result);
		}

		[TestMethod]
		public void FloorOfNegativeTwoPointFiveIntoNegativeTwoIsNegativeTwo()
		{
			var func = new Floor();
			var args = FunctionsHelper.CreateArgs(-2.5, -2);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(-2.0, result.Result);
		}

		[TestMethod]
		public void FloorOfNegativeFourIntoNegativeTwoIsNegativeFour()
		{
			var func = new Floor();
			var args = FunctionsHelper.CreateArgs(-4, -2);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(-4.0, result.Result);
		}

		[TestMethod]
		public void FloorOfNegativeFourIntoTwoIsNegativeFour()
		{
			var func = new Floor();
			var args = FunctionsHelper.CreateArgs(-4, 2);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(-4.0, result.Result);
		}

		[TestMethod]
		public void FloorOfNegativeIntoPositiveNumberRoundsDownTowardsNegativeInfinity()
		{
			var func = new Floor();
			var args = FunctionsHelper.CreateArgs(-2.5, 2);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(-4.0, result.Result);
		}

		[TestMethod]
		public void FloorOfPositiveIntoNegativeNumberIsPoundNum()
		{
			var func = new Floor();
			var args = FunctionsHelper.CreateArgs(2.5, -2);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result.Result);
		}

		[TestMethod]
		public void FloorWithZeroSignificancePoundDivZeroes()
		{
			var func = new Floor();
			var args = FunctionsHelper.CreateArgs(9, 0);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Div0), result.Result);
		}

		[TestMethod]
		public void FloorWithOneArgumentPoundValues()
		{
			var func = new Floor();
			var args = FunctionsHelper.CreateArgs(9);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), result.Result);
		}

		[TestMethod]
		public void FloorOfZeroIntoAnythingIsZero()
		{
			var rng = new Random();
			var func = new Floor();
			for (int i = 0; i < 1000; i++)
			{
				var args = FunctionsHelper.CreateArgs(0, rng.Next());
				var result = func.Execute(args, _parsingContext);
				Assert.AreEqual(0.0, result.Result);
			}
			for (int i = 0; i < 1000; i++)
			{
				var args = FunctionsHelper.CreateArgs(0, -rng.Next());
				var result = func.Execute(args, _parsingContext);
				Assert.AreEqual(0.0, result.Result);
			}
			for (int i = 0; i < 1000; i++)
			{
				var args = FunctionsHelper.CreateArgs(0, rng.Next() + rng.NextDouble());
				var result = func.Execute(args, _parsingContext);
				Assert.AreEqual(0.0, result.Result);
			}
		}

		[TestMethod]
		public void FloorShouldReturnCorrectResultWhenSignificanceIsBetween0And1()
		{
			var func = new Floor();
			var args = FunctionsHelper.CreateArgs(26.75d, 0.1);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(26.7d, (double)result.Result, 0.000000001);
		}

		[TestMethod]
		public void FloorShouldReturnCorrectResultWhenSignificanceIs1()
		{
			var func = new Floor();
			var args = FunctionsHelper.CreateArgs(26.75d, 1);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(26d, result.Result);
		}

		[TestMethod]
		public void FloorShouldReturnCorrectResultWhenSignificanceIsMinus1()
		{
			var func = new Floor();
			var args = FunctionsHelper.CreateArgs(-26.75d, -1);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(-26d, result.Result);
		}

		[TestMethod]
		public void FloorFunctionWithErrorValuesAsInputReturnsTheInputErrorValue()
		{
			var func = new Floor();
			var argNA = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.NA),5);
			var argNAME = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Name),5);
			var argVALUE = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Value),5);
			var argNUM = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Num),5);
			var argDIV0 = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Div0),5);
			var argREF = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Ref),5);
			var resultNA = func.Execute(argNA, this.ParsingContext);
			var resultNAME = func.Execute(argNAME, this.ParsingContext);
			var resultVALUE = func.Execute(argVALUE, this.ParsingContext);
			var resultNUM = func.Execute(argNUM, this.ParsingContext);
			var resultDIV0 = func.Execute(argDIV0, this.ParsingContext);
			var resultREF = func.Execute(argREF, this.ParsingContext);
			Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)resultNA.Result).Type);
			Assert.AreEqual(eErrorType.Name, ((ExcelErrorValue)resultNAME.Result).Type);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)resultVALUE.Result).Type);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)resultNUM.Result).Type);
			Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)resultDIV0.Result).Type);
			Assert.AreEqual(eErrorType.Ref, ((ExcelErrorValue)resultREF.Result).Type);
		}

		[TestMethod]
		public void FloorWithNoInputsReturnsPoundValue()
		{
			var function = new Floor();
			var result = function.Execute(FunctionsHelper.CreateArgs(), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void FloorWithGeneralSringFirstInputReturnsPoundValue()
		{
			var function = new Floor();
			var result = function.Execute(FunctionsHelper.CreateArgs("string", 2), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void FloorWithNumericStringFirstInputReturnsCorrectValue()
		{
			var function = new Floor();
			var result = function.Execute(FunctionsHelper.CreateArgs("15", 2), this.ParsingContext);
			Assert.AreEqual(14d, result.Result);
		}

		[TestMethod]
		public void FloorWithDateFunctionFirstInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "FLOOR(DATE(2017, 6, 14), 6)";
				worksheet.Calculate();
				Assert.AreEqual(42900d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void FloorWithDateAsStringFirstInputReturnsCorrectValue()
		{
			var function = new Floor();
			var result = function.Execute(FunctionsHelper.CreateArgs("5/5/2017", 4), this.ParsingContext);
			Assert.AreEqual(42860d, result.Result);
		}

		[TestMethod]
		public void FloorWithBooleanFirstInputReturnsCorrectValue()
		{
			var function = new Floor();
			var booleanTrue = function.Execute(FunctionsHelper.CreateArgs(true, 5), this.ParsingContext);
			var booleanFalse = function.Execute(FunctionsHelper.CreateArgs(false, 3), this.ParsingContext);
			Assert.AreEqual(0d, booleanTrue.Result);
			Assert.AreEqual(0d, booleanFalse.Result);
		}

		[TestMethod]
		public void FloorWithCellReferenceFirstInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 16;
				worksheet.Cells["B2"].Formula = "FLOOR(B1, 7)";
				worksheet.Calculate();
				Assert.AreEqual(14d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void FloorWithGeneralStringSecondInputReturnsPoundValue()
		{
			var function = new Floor();
			var result = function.Execute(FunctionsHelper.CreateArgs(15, "string"), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void FloorWithNumericStringSecondInputReturnsCorrectValue()
		{
			var function = new Floor();
			var result = function.Execute(FunctionsHelper.CreateArgs(15, "7"), this.ParsingContext);
			Assert.AreEqual(14d, result.Result);
		}

		[TestMethod]
		public void FloorWithDateFunctionSecondInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "FLOOR(DATE(2017, 6, 14), 6)";
				worksheet.Calculate();
				Assert.AreEqual(42900d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void FloorWithDateAsStringSecondInputReturnsCorrectValue()
		{
			var function = new Floor();
			var result = function.Execute(FunctionsHelper.CreateArgs(15, "5/5/2017"), this.ParsingContext);
			Assert.AreEqual(0d, result.Result);
		}

		[TestMethod]
		public void FloorWithBooleanSecondInputsReturnsCorrectValue()
		{
			var function = new Floor();
			var booleanTrue = function.Execute(FunctionsHelper.CreateArgs(5, true), this.ParsingContext);
			var booleanFalse = function.Execute(FunctionsHelper.CreateArgs(15, false), this.ParsingContext);
			Assert.AreEqual(5d, booleanTrue.Result);
			Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)booleanFalse.Result).Type);
		}

		[TestMethod]
		public void FloorWithCellReferenceSecondInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 4;
				worksheet.Cells["B2"].Formula = "FLOOR(67, B1)";
				worksheet.Calculate();
				Assert.AreEqual(64d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void FloorWithNoSecondInputReturnsDivZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "FLOOR(15, )";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B1"].Value).Type);
			}
		}

		[TestMethod]
		public void FloorWithNoFirstInputReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "FLOOR(, 6)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B1"].Value);
			}
		}
	}
}
