using System;
using System.Collections.Generic;
using System.Linq;
using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;

namespace EPPlusTest.Excel.Functions
{
	[TestClass]
	public class MathFunctionsTests
	{
		private ParsingContext _parsingContext;

		[TestInitialize]
		public void Initialize()
		{
			_parsingContext = ParsingContext.Create();
			_parsingContext.Scopes.NewScope(RangeAddress.Empty);
		}

		[TestMethod]
		public void PiShouldReturnPIConstant()
		{
			var expectedValue = (double)Math.Round(Math.PI, 14);
			var func = new Pi();
			var args = FunctionsHelper.CreateArgs(0);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(expectedValue, result.Result);
		}

		[TestMethod]
		public void AsinShouldReturnCorrectResult()
		{
			const double expectedValue = 1.5708;
			var func = new Asin();
			var args = FunctionsHelper.CreateArgs(1d);
			var result = func.Execute(args, _parsingContext);
			var rounded = Math.Round((double)result.Result, 4);
			Assert.AreEqual(expectedValue, rounded);
		}

		[TestMethod]
		public void AsinhShouldReturnCorrectResult()
		{
			const double expectedValue = 0.0998;
			var func = new Asinh();
			var args = FunctionsHelper.CreateArgs(0.1d);
			var result = func.Execute(args, _parsingContext);
			var rounded = Math.Round((double)result.Result, 4);
			Assert.AreEqual(expectedValue, rounded);
		}

		[TestMethod]
		public void SumSqShouldCalculateArray()
		{
			var func = new Sumsq();
			var args = FunctionsHelper.CreateArgs(2, 4);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(20d, result.Result);
		}

		[TestMethod]
		public void SumSqShouldIncludeTrueAsOne()
		{
			var func = new Sumsq();
			var args = FunctionsHelper.CreateArgs(2, 4, true);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(21d, result.Result);
		}

		[TestMethod]
		public void SumSqShouldNoCountTrueTrueInArray()
		{
			var func = new Sumsq();
			var args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(2, 4, true));
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(20d, result.Result);
		}

		[TestMethod]
		public void ExpShouldCalculateCorrectResult()
		{
			var func = new Exp();
			var args = FunctionsHelper.CreateArgs(4);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(54.59815003d, System.Math.Round((double)result.Result, 8));
		}

		[TestMethod]
		public void AverageShouldCalculateCorrectResult()
		{
			var expectedResult = (4d + 2d + 5d + 2d) / 4d;
			var func = new Average();
			var args = FunctionsHelper.CreateArgs(4d, 2d, 5d, 2d);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(expectedResult, result.Result);
		}

		[TestMethod]
		public void AverageShouldCalculateCorrectResultWithEnumerableAndBoolMembers()
		{
			var expectedResult = (4d + 2d + 5d + 2d + 1d) / 5d;
			var func = new Average();
			var args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(4d, 2d), 5d, 2d, true);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(expectedResult, result.Result);
		}

		[TestMethod]
		public void AverageShouldIgnoreHiddenFieldsIfIgnoreHiddenValuesIsTrue()
		{
			var expectedResult = (4d + 2d + 2d + 1d) / 4d;
			var func = new Average();
			func.IgnoreHiddenValues = true;
			var args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(4d, 2d), 5d, 2d, true);
			args.ElementAt(1).SetExcelStateFlag(ExcelCellState.HiddenCell);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(expectedResult, result.Result);
		}

		[TestMethod]
		public void AverageShouldThrowDivByZeroExcelErrorValueIfEmptyArgs()
		{
			var func = new Average();
			var args = new FunctionArgument[0];
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void AverageAShouldCalculateCorrectResult()
		{
			var expectedResult = (4d + 2d + 5d + 2d) / 4d;
			var func = new AverageA();
			var args = FunctionsHelper.CreateArgs(4d, 2d, 5d, 2d);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(expectedResult, result.Result);
		}

		[TestMethod]
		public void AverageAShouldIncludeTrueAs1()
		{
			var expectedResult = (4d + 2d + 5d + 2d + 1d) / 5d;
			var func = new AverageA();
			var args = FunctionsHelper.CreateArgs(4d, 2d, 5d, 2d, true);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(expectedResult, result.Result);
		}

		[TestMethod]
		public void AverageAShouldThrowValueExceptionIfNonNumericTextIsSupplied()
		{
			var func = new AverageA();
			var args = FunctionsHelper.CreateArgs(4d, 2d, 5d, 2d, "ABC");
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(OfficeOpenXml.FormulaParsing.ExpressionGraph.DataType.ExcelError, result.DataType);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)(result.Result)).Type);
		}

		[TestMethod]
		public void AverageAShouldCountValueAs0IfNonNumericTextIsSuppliedInArray()
		{
			var func = new AverageA();
			var args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(1d, 2d, 3d, "ABC"));
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(1.5d, result.Result);
		}

		[TestMethod]
		public void AverageAShouldCountNumericStringWithValue()
		{
			var func = new AverageA();
			var args = FunctionsHelper.CreateArgs(4d, 2d, "9");
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(5d, result.Result);
		}

		[TestMethod]
		public void RandShouldReturnAValueBetween0and1()
		{
			var func = new Rand();
			var args = new FunctionArgument[0];
			var result1 = func.Execute(args, _parsingContext);
			Assert.IsTrue(((double)result1.Result) > 0 && ((double)result1.Result) < 1);
			var result2 = func.Execute(args, _parsingContext);
			Assert.AreNotEqual(result1.Result, result2.Result, "The two numbers were the same");
			Assert.IsTrue(((double)result2.Result) > 0 && ((double)result2.Result) < 1);
		}

		[TestMethod]
		public void RandBetweenShouldReturnAnIntegerValueBetweenSuppliedValues()
		{
			var func = new RandBetween();
			var args = FunctionsHelper.CreateArgs(1, 5);
			var result = func.Execute(args, _parsingContext);
			CollectionAssert.Contains(new List<double> { 1d, 2d, 3d, 4d, 5d }, result.Result);
		}

		[TestMethod]
		public void RandBetweenShouldReturnAnIntegerValueBetweenSuppliedValuesWhenLowIsNegative()
		{
			var func = new RandBetween();
			var args = FunctionsHelper.CreateArgs(-5, 0);
			var result = func.Execute(args, _parsingContext);
			CollectionAssert.Contains(new List<double> { 0d, -1d, -2d, -3d, -4d, -5d }, result.Result);
		}

		[TestMethod]
		public void CountShouldReturnNumberOfNumericItems()
		{
			var func = new Count();
			var args = FunctionsHelper.CreateArgs(1d, 2m, 3, new DateTime(2012, 4, 1), "4");
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(5d, result.Result);
		}

		[TestMethod]
		public void CountShouldIncludeEnumerableMembers()
		{
			var func = new Count();
			var args = FunctionsHelper.CreateArgs(1d, FunctionsHelper.CreateArgs(12, 13));
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(3d, result.Result);
		}

		[TestMethod, Ignore]
		public void CountShouldIgnoreHiddenValuesIfIgnoreHiddenValuesIsTrue()
		{
			var func = new Count();
			func.IgnoreHiddenValues = true;
			var args = FunctionsHelper.CreateArgs(1d, FunctionsHelper.CreateArgs(12, 13));
			args.ElementAt(0).SetExcelStateFlag(ExcelCellState.HiddenCell);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(2d, result.Result);
		}

		[TestMethod]
		public void CountAShouldIncludeEnumerableMembers()
		{
			var func = new CountA();
			var args = FunctionsHelper.CreateArgs(1d, FunctionsHelper.CreateArgs(12, 13));
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(3d, result.Result);
		}

		[TestMethod]
		public void CountAShouldIgnoreHiddenValuesIfIgnoreHiddenValuesIsTrue()
		{
			var func = new CountA();
			func.IgnoreHiddenValues = true;
			var args = FunctionsHelper.CreateArgs(1d, FunctionsHelper.CreateArgs(12, 13));
			args.ElementAt(0).SetExcelStateFlag(ExcelCellState.HiddenCell);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(2d, result.Result);
		}

		[TestMethod]
		public void CosShouldReturnCorrectResult()
		{
			var func = new Cos();
			var args = FunctionsHelper.CreateArgs(2);
			var result = func.Execute(args, _parsingContext);
			var roundedResult = Math.Round((double)result.Result, 9);
			Assert.AreEqual(-0.416146837d, roundedResult);
		}

		[TestMethod]
		public void CosHShouldReturnCorrectResult()
		{
			var func = new Cosh();
			var args = FunctionsHelper.CreateArgs(2);
			var result = func.Execute(args, _parsingContext);
			var roundedResult = Math.Round((double)result.Result, 9);
			Assert.AreEqual(3.762195691, roundedResult);
		}

		[TestMethod]
		public void AcosShouldReturnCorrectResult()
		{
			var func = new Acos();
			var args = FunctionsHelper.CreateArgs(0.1);
			var result = func.Execute(args, _parsingContext);
			var roundedResult = Math.Round((double)result.Result, 4);
			Assert.AreEqual(1.4706, roundedResult);
		}

		[TestMethod]
		public void ACosHShouldReturnCorrectResult()
		{
			var func = new Acosh();
			var args = FunctionsHelper.CreateArgs(2);
			var result = func.Execute(args, _parsingContext);
			var roundedResult = Math.Round((double)result.Result, 3);
			Assert.AreEqual(1.317, roundedResult);
		}

		[TestMethod]
		public void SinShouldReturnCorrectResult()
		{
			var func = new Sin();
			var args = FunctionsHelper.CreateArgs(2);
			var result = func.Execute(args, _parsingContext);
			var roundedResult = Math.Round((double)result.Result, 9);
			Assert.AreEqual(0.909297427, roundedResult);
		}

		[TestMethod]
		public void SinhShouldReturnCorrectResult()
		{
			var func = new Sinh();
			var args = FunctionsHelper.CreateArgs(2);
			var result = func.Execute(args, _parsingContext);
			var roundedResult = Math.Round((double)result.Result, 9);
			Assert.AreEqual(3.626860408d, roundedResult);
		}

		[TestMethod]
		public void TanShouldReturnCorrectResult()
		{
			var func = new Tan();
			var args = FunctionsHelper.CreateArgs(2);
			var result = func.Execute(args, _parsingContext);
			var roundedResult = Math.Round((double)result.Result, 9);
			Assert.AreEqual(-2.185039863d, roundedResult);
		}

		[TestMethod]
		public void TanhShouldReturnCorrectResult()
		{
			var func = new Tanh();
			var args = FunctionsHelper.CreateArgs(2);
			var result = func.Execute(args, _parsingContext);
			var roundedResult = Math.Round((double)result.Result, 9);
			Assert.AreEqual(0.96402758d, roundedResult);
		}

		[TestMethod]
		public void AtanShouldReturnCorrectResult()
		{
			var func = new Atan();
			var args = FunctionsHelper.CreateArgs(10);
			var result = func.Execute(args, _parsingContext);
			var roundedResult = Math.Round((double)result.Result, 9);
			Assert.AreEqual(1.471127674d, roundedResult);
		}

		[TestMethod]
		public void Atan2ShouldReturnCorrectResult()
		{
			var func = new Atan2();
			var args = FunctionsHelper.CreateArgs(1, 2);
			var result = func.Execute(args, _parsingContext);
			var roundedResult = Math.Round((double)result.Result, 9);
			Assert.AreEqual(1.107148718d, roundedResult);
		}

		[TestMethod]
		public void AtanhShouldReturnCorrectResult()
		{
			var func = new Atanh();
			var args = FunctionsHelper.CreateArgs(0.1);
			var result = func.Execute(args, _parsingContext);
			var roundedResult = Math.Round((double)result.Result, 4);
			Assert.AreEqual(0.1003d, roundedResult);
		}

		[TestMethod]
		public void SqrtPiShouldReturnCorrectResult()
		{
			var func = new SqrtPi();
			var args = FunctionsHelper.CreateArgs(2);
			var result = func.Execute(args, _parsingContext);
			var roundedResult = Math.Round((double)result.Result, 9);
			Assert.AreEqual(2.506628275d, roundedResult);
		}

		[TestMethod]
		public void TruncShouldReturnCorrectResult()
		{
			var func = new Trunc();
			var args = FunctionsHelper.CreateArgs(99.99);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(99d, result.Result);
		}

		[TestMethod]
		public void FactShouldRoundDownAndReturnCorrectResult()
		{
			var func = new Fact();
			var args = FunctionsHelper.CreateArgs(5.99);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(120d, result.Result);
		}

		[TestMethod]
		public void FactShouldReturnPoundNumWhenNegativeNumber()
		{
			var func = new Fact();
			var args = FunctionsHelper.CreateArgs(-1);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(OfficeOpenXml.FormulaParsing.ExpressionGraph.DataType.ExcelError, result.DataType);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)(result.Result)).Type);
		}

		[TestMethod]
		public void CountIfShouldHandleNegativeCriteria()
		{
			using (var package = new ExcelPackage())
			{
				var sheet1 = package.Workbook.Worksheets.Add("test");
				sheet1.Cells["A1"].Value = -1;
				sheet1.Cells["A2"].Value = -2;
				sheet1.Cells["A3"].Formula = "CountIf(A1:A2,\"-1\")";
				sheet1.Calculate();
				Assert.AreEqual(1d, sheet1.Cells["A3"].Value);
			}
		}
		[TestMethod]
		public void Rank()
		{
			using (var p = new ExcelPackage())
			{
				var w = p.Workbook.Worksheets.Add("testsheet");
				w.SetValue(1, 1, 1);
				w.SetValue(2, 1, 1);
				w.SetValue(3, 1, 2);
				w.SetValue(4, 1, 2);
				w.SetValue(5, 1, 4);
				w.SetValue(6, 1, 4);

				w.SetFormula(1, 2, "RANK(1,A1:A5)");
				w.SetFormula(1, 3, "RANK(1,A1:A5,1)");
				w.SetFormula(1, 4, "RANK.AVG(1,A1:A5)");
				w.SetFormula(1, 5, "RANK.AVG(1,A1:A5,1)");

				w.SetFormula(2, 2, "RANK.EQ(2,A1:A5)");
				w.SetFormula(2, 3, "RANK.EQ(2,A1:A5,1)");
				w.SetFormula(2, 4, "RANK.AVG(2,A1:A5,1)");
				w.SetFormula(2, 5, "RANK.AVG(2,A1:A5,0)");

				w.SetFormula(3, 2, "RANK(3,A1:A5)");
				w.SetFormula(3, 3, "RANK(3,A1:A5,1)");
				w.SetFormula(3, 4, "RANK.AVG(3,A1:A5,1)");
				w.SetFormula(3, 5, "RANK.AVG(3,A1:A5,0)");

				w.SetFormula(4, 2, "RANK.EQ(4,A1:A5)");
				w.SetFormula(4, 3, "RANK.EQ(4,A1:A5,1)");
				w.SetFormula(4, 4, "RANK.AVG(4,A1:A5,1)");
				w.SetFormula(4, 5, "RANK.AVG(4,A1:A5)");


				w.SetFormula(5, 4, "RANK.AVG(4,A1:A6,1)");
				w.SetFormula(5, 5, "RANK.AVG(4,A1:A6)");

				w.Calculate();

				Assert.AreEqual(w.GetValue(1, 2), 4D);
				Assert.AreEqual(w.GetValue(1, 3), 1D);
				Assert.AreEqual(w.GetValue(1, 4), 4.5D);
				Assert.AreEqual(w.GetValue(1, 5), 1.5D);

				Assert.AreEqual(w.GetValue(2, 2), 2D);
				Assert.AreEqual(w.GetValue(2, 3), 3D);
				Assert.AreEqual(w.GetValue(2, 4), 3.5D);
				Assert.AreEqual(w.GetValue(2, 5), 2.5D);

				Assert.IsInstanceOfType(w.GetValue(3, 2), typeof(ExcelErrorValue));
				Assert.IsInstanceOfType(w.GetValue(3, 3), typeof(ExcelErrorValue));
				Assert.IsInstanceOfType(w.GetValue(3, 4), typeof(ExcelErrorValue));
				Assert.IsInstanceOfType(w.GetValue(3, 5), typeof(ExcelErrorValue));

				Assert.AreEqual(w.GetValue(4, 2), 1D);
				Assert.AreEqual(w.GetValue(4, 3), 5D);
				Assert.AreEqual(w.GetValue(4, 4), 5D);
				Assert.AreEqual(w.GetValue(4, 5), 1D);

				Assert.AreEqual(w.GetValue(5, 4), 5.5D);
				Assert.AreEqual(w.GetValue(5, 5), 1.5D);
			}
		}

		[TestMethod]
		public void AcosWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Acos();
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void AcoshWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Acosh();
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void AsinWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Asin();
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void AsinhWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Asinh();
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void AtanWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Atan();
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void Atan2WithInvalidArgumentReturnsPoundValue()
		{
			var func = new Atan2();
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void AtanhWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Atanh();
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void AverageIfWithInvalidArgumentReturnsPoundValue()
		{
			var func = new AverageIf();
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void AverageIfsWithInvalidArgumentReturnsPoundValue()
		{
			var func = new AverageIfs();
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void CosWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Cos();
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void CoshWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Cosh();
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void CountWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Count();
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void CountAWithInvalidArgumentReturnsPoundValue()
		{
			var func = new CountA();
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void CountBlankWithInvalidArgumentReturnsPoundValue()
		{
			var func = new CountBlank();
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void CountIfWithInvalidArgumentReturnsPoundValue()
		{
			var func = new CountIf();
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void CountIfsWithInvalidArgumentReturnsPoundValue()
		{
			var func = new CountIfs();
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void DegreesWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Degrees();
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void ExpWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Exp();
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void FactWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Fact();
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void RandBetweenWithInvalidArgumentReturnsPoundValue()
		{
			var func = new RandBetween();
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SinWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Sin();
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SinhWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Sinh();
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SqrtWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Sqrt();
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SqrtPiWithInvalidArgumentReturnsPoundValue()
		{
			var func = new SqrtPi();
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SubtotalWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Subtotal();
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SumIfWithInvalidArgumentReturnsPoundValue()
		{
			var func = new SumIf();
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SumIfsWithInvalidArgumentReturnsPoundValue()
		{
			var func = new SumIfs();
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void TanWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Tan();
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void TanhWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Tanh();
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void TruncWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Trunc();
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}
	}
}
