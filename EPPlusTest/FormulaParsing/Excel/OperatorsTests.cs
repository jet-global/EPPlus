using System;
using System.Globalization;
using System.IO;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace EPPlusTest.Excel
{
	[TestClass]
	public class OperatorsTests
	{
		#region Logical Comparison Operator Tests
		[TestMethod]
		public void OperatorLogicalOperatorsShouldCorrectlyCompareIntegersAndText()
		{
			// Numbers are always strictly less than text.
			Assert.AreEqual(false, Operator.GreaterThan.Apply(new CompileResult(1000000, DataType.Integer), new CompileResult("Numbers are always less than text.", DataType.String)).Result);
			Assert.AreEqual(false, Operator.GreaterThan.Apply(new CompileResult(1000000, DataType.Integer), new CompileResult("1", DataType.String)).Result);
			Assert.AreEqual(false, Operator.GreaterThanOrEqual.Apply(new CompileResult(1000000, DataType.Integer), new CompileResult("1", DataType.String)).Result);
			Assert.AreEqual(true, Operator.LessThan.Apply(new CompileResult(1000000, DataType.Integer), new CompileResult("1", DataType.String)).Result);
			Assert.AreEqual(true, Operator.LessThanOrEqual.Apply(new CompileResult(1000000, DataType.Integer), new CompileResult("1", DataType.String)).Result);
			Assert.AreEqual(false, Operator.EqualsTo.Apply(new CompileResult(1, DataType.Integer), new CompileResult("1", DataType.String)).Result);
			Assert.AreEqual(true, Operator.NotEqualsTo.Apply(new CompileResult(1, DataType.Integer), new CompileResult("1", DataType.String)).Result);

			Assert.AreEqual(true, Operator.GreaterThan.Apply(new CompileResult("Text is greater than numbers.", DataType.String), new CompileResult(1000000, DataType.Integer)).Result);
			Assert.AreEqual(true, Operator.GreaterThan.Apply(new CompileResult("1", DataType.String), new CompileResult(1000000, DataType.Integer)).Result);
			Assert.AreEqual(true, Operator.GreaterThanOrEqual.Apply(new CompileResult("1", DataType.String), new CompileResult(1000000, DataType.Integer)).Result);
			Assert.AreEqual(false, Operator.LessThan.Apply(new CompileResult("1", DataType.String), new CompileResult(1000000, DataType.Integer)).Result);
			Assert.AreEqual(false, Operator.LessThanOrEqual.Apply(new CompileResult("1", DataType.String), new CompileResult(1000000, DataType.Integer)).Result);
		}

		[TestMethod]
		public void OperatorLogicalOperatorsShouldCorrectlyCompareDecimalsAndText()
		{
			// Numbers are always strictly less than text.
			Assert.AreEqual(false, Operator.GreaterThan.Apply(new CompileResult(1000000.0d, DataType.Decimal), new CompileResult("Numbers are always less than text.", DataType.String)).Result);
			Assert.AreEqual(false, Operator.GreaterThan.Apply(new CompileResult(1000000.0d, DataType.Decimal), new CompileResult("1", DataType.String)).Result);
			Assert.AreEqual(false, Operator.GreaterThanOrEqual.Apply(new CompileResult(1000000.0d, DataType.Decimal), new CompileResult("1", DataType.String)).Result);
			Assert.AreEqual(true, Operator.LessThan.Apply(new CompileResult(1000000.0d, DataType.Decimal), new CompileResult("1", DataType.String)).Result);
			Assert.AreEqual(true, Operator.LessThanOrEqual.Apply(new CompileResult(1000000.0d, DataType.Decimal), new CompileResult("1", DataType.String)).Result);
			Assert.AreEqual(false, Operator.EqualsTo.Apply(new CompileResult(1, DataType.Decimal), new CompileResult("1", DataType.String)).Result);
			Assert.AreEqual(true, Operator.NotEqualsTo.Apply(new CompileResult(1, DataType.Decimal), new CompileResult("1", DataType.String)).Result);

			Assert.AreEqual(true, Operator.GreaterThan.Apply(new CompileResult("Text is greater than numbers.", DataType.String), new CompileResult(1000000.0d, DataType.Decimal)).Result);
			Assert.AreEqual(true, Operator.GreaterThan.Apply(new CompileResult("1", DataType.String), new CompileResult(1000000.0d, DataType.Decimal)).Result);
			Assert.AreEqual(true, Operator.GreaterThanOrEqual.Apply(new CompileResult("1", DataType.String), new CompileResult(1000000.0d, DataType.Decimal)).Result);
			Assert.AreEqual(false, Operator.LessThan.Apply(new CompileResult("1", DataType.String), new CompileResult(1000000.0d, DataType.Decimal)).Result);
			Assert.AreEqual(false, Operator.LessThanOrEqual.Apply(new CompileResult("1", DataType.String), new CompileResult(1000000.0d, DataType.Decimal)).Result);
		}

		[TestMethod]
		public void OperatorLogicalOperatorsShouldCorrectlyCompareNumbersAndLogicalValues()
		{
			// Logical values are strictly larger than all numeric values.
			Assert.AreEqual(true, Operator.GreaterThan.Apply(new CompileResult(true, DataType.Boolean), new CompileResult(int.MaxValue, DataType.Integer)).Result);
			Assert.AreEqual(true, Operator.GreaterThan.Apply(new CompileResult(false, DataType.Boolean), new CompileResult(int.MaxValue, DataType.Integer)).Result);
			Assert.AreEqual(true, Operator.GreaterThan.Apply(new CompileResult(true, DataType.Boolean), new CompileResult(int.MinValue, DataType.Integer)).Result);
			Assert.AreEqual(true, Operator.GreaterThan.Apply(new CompileResult(false, DataType.Boolean), new CompileResult(int.MinValue, DataType.Integer)).Result);
			Assert.AreEqual(true, Operator.GreaterThanOrEqual.Apply(new CompileResult(false, DataType.Boolean), new CompileResult(int.MaxValue, DataType.Integer)).Result);
			Assert.AreEqual(false, Operator.LessThan.Apply(new CompileResult(true, DataType.Boolean), new CompileResult(int.MinValue, DataType.Integer)).Result);
			Assert.AreEqual(false, Operator.LessThan.Apply(new CompileResult(false, DataType.Boolean), new CompileResult(int.MinValue, DataType.Integer)).Result);
			Assert.AreEqual(false, Operator.LessThanOrEqual.Apply(new CompileResult(true, DataType.Boolean), new CompileResult(int.MinValue, DataType.Integer)).Result);
			Assert.AreEqual(false, Operator.EqualsTo.Apply(new CompileResult(false, DataType.Boolean), new CompileResult(0.0, DataType.Decimal)).Result);
			Assert.AreEqual(false, Operator.EqualsTo.Apply(new CompileResult(true, DataType.Boolean), new CompileResult(1.0, DataType.Decimal)).Result);
			Assert.AreEqual(true, Operator.NotEqualsTo.Apply(new CompileResult(false, DataType.Boolean), new CompileResult(0.0, DataType.Decimal)).Result);
			Assert.AreEqual(true, Operator.NotEqualsTo.Apply(new CompileResult(true, DataType.Boolean), new CompileResult(1.0, DataType.Decimal)).Result);

			Assert.AreEqual(false, Operator.GreaterThan.Apply(new CompileResult(int.MaxValue, DataType.Integer), new CompileResult(true, DataType.Boolean)).Result);
			Assert.AreEqual(false, Operator.GreaterThan.Apply(new CompileResult(int.MaxValue, DataType.Integer), new CompileResult(false, DataType.Boolean)).Result);
			Assert.AreEqual(false, Operator.GreaterThan.Apply(new CompileResult(int.MinValue, DataType.Integer), new CompileResult(true, DataType.Boolean)).Result);
			Assert.AreEqual(false, Operator.GreaterThan.Apply(new CompileResult(int.MinValue, DataType.Integer), new CompileResult(false, DataType.Boolean)).Result);
			Assert.AreEqual(false, Operator.GreaterThanOrEqual.Apply(new CompileResult(int.MaxValue, DataType.Integer), new CompileResult(false, DataType.Boolean)).Result);
			Assert.AreEqual(true, Operator.LessThan.Apply(new CompileResult(int.MinValue, DataType.Integer), new CompileResult(true, DataType.Boolean)).Result);
			Assert.AreEqual(true, Operator.LessThan.Apply(new CompileResult(int.MinValue, DataType.Integer), new CompileResult(false, DataType.Boolean)).Result);
			Assert.AreEqual(true, Operator.LessThanOrEqual.Apply(new CompileResult(int.MinValue, DataType.Integer), new CompileResult(true, DataType.Boolean)).Result);
		}

		[TestMethod]
		public void OperatorLogicalOperatorsShouldCorrectlyCompareTextAndLogicalValues()
		{
			// Logical values are always strictly greater than text.
			Assert.AreEqual(true, Operator.GreaterThan.Apply(new CompileResult(true, DataType.Boolean), new CompileResult(string.Empty, DataType.String)).Result);
			Assert.AreEqual(true, Operator.GreaterThan.Apply(new CompileResult(false, DataType.Boolean), new CompileResult(string.Empty, DataType.String)).Result);
			Assert.AreEqual(true, Operator.GreaterThan.Apply(new CompileResult(true, DataType.Boolean), new CompileResult("We are confident in this test because W > T", DataType.String)).Result);
			Assert.AreEqual(true, Operator.GreaterThan.Apply(new CompileResult(false, DataType.Boolean), new CompileResult("Gotta start with a letter bigger than F to be confident in this test", DataType.String)).Result);
			Assert.AreEqual(true, Operator.GreaterThanOrEqual.Apply(new CompileResult(false, DataType.Boolean), new CompileResult("ZZZZ huge string", DataType.String)).Result);
			Assert.AreEqual(false, Operator.LessThan.Apply(new CompileResult(true, DataType.Boolean), new CompileResult("arbitrary text", DataType.String)).Result);
			Assert.AreEqual(false, Operator.LessThan.Apply(new CompileResult(false, DataType.Boolean), new CompileResult("arbitrary text", DataType.String)).Result);
			Assert.AreEqual(false, Operator.LessThanOrEqual.Apply(new CompileResult(true, DataType.Boolean), new CompileResult(int.MinValue, DataType.String)).Result);
			Assert.AreEqual(false, Operator.EqualsTo.Apply(new CompileResult(false, DataType.Boolean), new CompileResult("FALSE", DataType.String)).Result);
			Assert.AreEqual(false, Operator.EqualsTo.Apply(new CompileResult(true, DataType.Boolean), new CompileResult("TRUE", DataType.String)).Result);
			Assert.AreEqual(true, Operator.NotEqualsTo.Apply(new CompileResult(false, DataType.Boolean), new CompileResult("FALSE", DataType.String)).Result);
			Assert.AreEqual(true, Operator.NotEqualsTo.Apply(new CompileResult(true, DataType.Boolean), new CompileResult("TRUE", DataType.String)).Result);

			Assert.AreEqual(false, Operator.GreaterThan.Apply(new CompileResult(string.Empty, DataType.String), new CompileResult(true, DataType.Boolean)).Result);
			Assert.AreEqual(false, Operator.GreaterThan.Apply(new CompileResult(string.Empty, DataType.String), new CompileResult(false, DataType.Boolean)).Result);
			Assert.AreEqual(false, Operator.GreaterThan.Apply(new CompileResult("We start with W", DataType.String), new CompileResult(true, DataType.Boolean)).Result);
			Assert.AreEqual(false, Operator.GreaterThan.Apply(new CompileResult("Might start with M", DataType.String), new CompileResult(false, DataType.Boolean)).Result);
			Assert.AreEqual(false, Operator.GreaterThanOrEqual.Apply(new CompileResult("Z is a pretty big string", DataType.String), new CompileResult(false, DataType.Boolean)).Result);
			Assert.AreEqual(true, Operator.LessThan.Apply(new CompileResult("Zzzz... Text is always less than logical values", DataType.String), new CompileResult(true, DataType.Boolean)).Result);
			Assert.AreEqual(true, Operator.LessThan.Apply(new CompileResult("Text", DataType.String), new CompileResult(false, DataType.Boolean)).Result);
			Assert.AreEqual(true, Operator.LessThanOrEqual.Apply(new CompileResult("Zzzz", DataType.String), new CompileResult(true, DataType.Boolean)).Result);
		}

		[TestMethod]
		public void OperatorLogicalOperatorsShouldCorrectlyCompareDecimalValues()
		{
			Assert.AreEqual(false, Operator.GreaterThan.Apply(new CompileResult(1000000.0, DataType.Decimal), new CompileResult(1000000.01, DataType.Decimal)).Result);
			Assert.AreEqual(false, Operator.GreaterThan.Apply(new CompileResult(-1000000, DataType.Decimal), new CompileResult(0, DataType.Decimal)).Result);
			Assert.AreEqual(true, Operator.LessThan.Apply(new CompileResult(100.0, DataType.Decimal), new CompileResult(101.0, DataType.Decimal)).Result);
			Assert.AreEqual(true, Operator.LessThanOrEqual.Apply(new CompileResult(100.0, DataType.Decimal), new CompileResult(101.0, DataType.Decimal)).Result);
			Assert.AreEqual(true, Operator.LessThanOrEqual.Apply(new CompileResult(100.0, DataType.Decimal), new CompileResult(100.0, DataType.Decimal)).Result);
			Assert.AreEqual(false, Operator.EqualsTo.Apply(new CompileResult(1.1, DataType.Decimal), new CompileResult(1.0, DataType.Decimal)).Result);
			Assert.AreEqual(true, Operator.EqualsTo.Apply(new CompileResult(1.0, DataType.Decimal), new CompileResult(1.0, DataType.Decimal)).Result);
			Assert.AreEqual(true, Operator.NotEqualsTo.Apply(new CompileResult(1.0, DataType.Decimal), new CompileResult(1.1, DataType.Decimal)).Result);

			Assert.AreEqual(true, Operator.GreaterThan.Apply(new CompileResult(1000000.01, DataType.Decimal), new CompileResult(1000000.0, DataType.Decimal)).Result);
			Assert.AreEqual(true, Operator.GreaterThan.Apply(new CompileResult(-1.0, DataType.Decimal), new CompileResult(-100000.0, DataType.Decimal)).Result);
			Assert.AreEqual(true, Operator.GreaterThanOrEqual.Apply(new CompileResult(1000000.1, DataType.Decimal), new CompileResult(1000000.0, DataType.Decimal)).Result);
			Assert.AreEqual(true, Operator.GreaterThanOrEqual.Apply(new CompileResult(1000000.0, DataType.Decimal), new CompileResult(1000000.0, DataType.Decimal)).Result);
			Assert.AreEqual(false, Operator.LessThan.Apply(new CompileResult(100000, DataType.Decimal), new CompileResult(1000, DataType.Decimal)).Result);
			Assert.AreEqual(false, Operator.LessThanOrEqual.Apply(new CompileResult(10.0, DataType.Decimal), new CompileResult(9.9, DataType.Decimal)).Result);
		}

		[TestMethod]
		public void OperatorLogicalOperatorsShouldCorrectlyCompareStrings()
		{
			// Text comparison operators are case-insensitive.
			Assert.AreEqual(false, Operator.GreaterThan.Apply(new CompileResult("A", DataType.String), new CompileResult("B", DataType.String)).Result);
			Assert.AreEqual(false, Operator.GreaterThan.Apply(new CompileResult("a", DataType.String), new CompileResult("B", DataType.String)).Result);
			Assert.AreEqual(false, Operator.GreaterThan.Apply(new CompileResult("\"", DataType.String), new CompileResult("a", DataType.String)).Result);
			Assert.AreEqual(true, Operator.LessThan.Apply(new CompileResult("A", DataType.String), new CompileResult("b", DataType.String)).Result);
			Assert.AreEqual(true, Operator.LessThan.Apply(new CompileResult("a", DataType.String), new CompileResult("B", DataType.String)).Result);
			Assert.AreEqual(true, Operator.LessThanOrEqual.Apply(new CompileResult("A", DataType.String), new CompileResult("B", DataType.String)).Result);
			Assert.AreEqual(true, Operator.LessThanOrEqual.Apply(new CompileResult("aaa", DataType.String), new CompileResult("B", DataType.String)).Result);
			Assert.AreEqual(true, Operator.EqualsTo.Apply(new CompileResult("abcde", DataType.String), new CompileResult("AbCdE", DataType.String)).Result);
			Assert.AreEqual(false, Operator.NotEqualsTo.Apply(new CompileResult("abcde", DataType.String), new CompileResult("abcde", DataType.String)).Result);

			Assert.AreEqual(true, Operator.GreaterThan.Apply(new CompileResult("B", DataType.String), new CompileResult("A", DataType.String)).Result);
			Assert.AreEqual(true, Operator.GreaterThan.Apply(new CompileResult("The first character that doesn't match is controlling", DataType.String), new CompileResult("The first character is smaller here", DataType.String)).Result);
			Assert.AreEqual(true, Operator.GreaterThanOrEqual.Apply(new CompileResult("Dogs", DataType.String), new CompileResult("Cats", DataType.String)).Result);
			Assert.AreEqual(true, Operator.GreaterThanOrEqual.Apply(new CompileResult("dogs", DataType.String), new CompileResult("Dogs", DataType.String)).Result);
			Assert.AreEqual(false, Operator.LessThan.Apply(new CompileResult("Dogs", DataType.String), new CompileResult("Cats", DataType.String)).Result);
			Assert.AreEqual(false, Operator.LessThanOrEqual.Apply(new CompileResult("Dogs", DataType.String), new CompileResult("Cheetahs", DataType.String)).Result);
		}

		[TestMethod]
		public void OperatorLogicalOperatorsShouldCorrectlyCompareLogicalValues()
		{
			// TRUE is strictly greater than FALSE.
			Assert.AreEqual(true, Operator.GreaterThan.Apply(new CompileResult(true, DataType.Boolean), new CompileResult(false, DataType.Boolean)).Result);
			Assert.AreEqual(true, Operator.GreaterThanOrEqual.Apply(new CompileResult(true, DataType.Boolean), new CompileResult(false, DataType.Boolean)).Result);
			Assert.AreEqual(false, Operator.EqualsTo.Apply(new CompileResult(true, DataType.Boolean), new CompileResult(false, DataType.Boolean)).Result);
			Assert.AreEqual(true, Operator.NotEqualsTo.Apply(new CompileResult(true, DataType.Boolean), new CompileResult(false, DataType.Boolean)).Result);
			Assert.AreEqual(true, Operator.EqualsTo.Apply(new CompileResult(true, DataType.Boolean), new CompileResult(true, DataType.Boolean)).Result);
			Assert.AreEqual(true, Operator.GreaterThanOrEqual.Apply(new CompileResult(true, DataType.Boolean), new CompileResult(true, DataType.Boolean)).Result);
			Assert.AreEqual(true, Operator.LessThanOrEqual.Apply(new CompileResult(true, DataType.Boolean), new CompileResult(true, DataType.Boolean)).Result);
			Assert.AreEqual(false, Operator.GreaterThan.Apply(new CompileResult(true, DataType.Boolean), new CompileResult(true, DataType.Boolean)).Result);
			Assert.AreEqual(false, Operator.LessThan.Apply(new CompileResult(true, DataType.Boolean), new CompileResult(true, DataType.Boolean)).Result);
			Assert.AreEqual(true, Operator.EqualsTo.Apply(new CompileResult(false, DataType.Boolean), new CompileResult(false, DataType.Boolean)).Result);
			Assert.AreEqual(true, Operator.GreaterThanOrEqual.Apply(new CompileResult(false, DataType.Boolean), new CompileResult(false, DataType.Boolean)).Result);
			Assert.AreEqual(true, Operator.LessThanOrEqual.Apply(new CompileResult(false, DataType.Boolean), new CompileResult(false, DataType.Boolean)).Result);
			Assert.AreEqual(false, Operator.LessThan.Apply(new CompileResult(false, DataType.Boolean), new CompileResult(false, DataType.Boolean)).Result);
			Assert.AreEqual(false, Operator.GreaterThan.Apply(new CompileResult(false, DataType.Boolean), new CompileResult(false, DataType.Boolean)).Result);
			Assert.AreEqual(false, Operator.LessThan.Apply(new CompileResult(true, DataType.Boolean), new CompileResult(false, DataType.Boolean)).Result);
			Assert.AreEqual(false, Operator.LessThanOrEqual.Apply(new CompileResult(true, DataType.Boolean), new CompileResult(false, DataType.Boolean)).Result);
			Assert.AreEqual(false, Operator.GreaterThan.Apply(new CompileResult(false, DataType.Boolean), new CompileResult(true, DataType.Boolean)).Result);
			Assert.AreEqual(false, Operator.GreaterThanOrEqual.Apply(new CompileResult(false, DataType.Boolean), new CompileResult(true, DataType.Boolean)).Result);
			Assert.AreEqual(true, Operator.LessThan.Apply(new CompileResult(false, DataType.Boolean), new CompileResult(true, DataType.Boolean)).Result);
			Assert.AreEqual(true, Operator.LessThanOrEqual.Apply(new CompileResult(false, DataType.Boolean), new CompileResult(true, DataType.Boolean)).Result);
		}
		#endregion

		#region Operator Plus Tests
		[TestMethod]
		public void OperatorPlusShouldThrowExceptionIfNonNumericOperand()
		{
			var result = Operator.Plus.Apply(new CompileResult(1, DataType.Integer), new CompileResult("a", DataType.String));
			Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), result.Result);
		}

		[TestMethod]
		public void OperatorPlusShouldAddNumericStringAndNumber()
		{
			var result = Operator.Plus.Apply(new CompileResult(1, DataType.Integer), new CompileResult("2", DataType.String));
			Assert.AreEqual(3d, result.Result);
		}
		#endregion

		#region Operator Minus Tests
		[TestMethod]
		public void OperatorMinusShouldThrowExceptionIfNonNumericOperand()
		{
			var result = Operator.Minus.Apply(new CompileResult(1, DataType.Integer), new CompileResult("a", DataType.String));
			Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), result.Result);
		}

		[TestMethod]
		public void OperatorMinusShouldSubtractNumericStringAndNumber()
		{
			var result = Operator.Minus.Apply(new CompileResult(5, DataType.Integer), new CompileResult("2", DataType.String));
			Assert.AreEqual(3d, result.Result);
		}
		#endregion

		#region Operator Divide Tests
		[TestMethod]
		public void OperatorDivideShouldReturnDivideByZeroIfRightOperandIsZero()
		{
			var result = Operator.Divide.Apply(new CompileResult(1d, DataType.Decimal), new CompileResult(0d, DataType.Decimal));
			Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Div0), result.Result);
		}

		[TestMethod]
		public void OperatorDivideShouldDivideCorrectly()
		{
			var result = Operator.Divide.Apply(new CompileResult(9d, DataType.Decimal), new CompileResult(3d, DataType.Decimal));
			Assert.AreEqual(3d, result.Result);
		}

		[TestMethod]
		public void OperatorDivideShouldReturnValueErrorIfNonNumericOperand()
		{
			var result = Operator.Divide.Apply(new CompileResult(1, DataType.Integer), new CompileResult("a", DataType.String));
			Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), result.Result);
		}

		[TestMethod]
		public void OperatorDivideShouldDivideNumericStringAndNumber()
		{
			var result = Operator.Divide.Apply(new CompileResult(9, DataType.Integer), new CompileResult("3", DataType.String));
			Assert.AreEqual(3d, result.Result);
		}
		#endregion

		#region Operator Multiply Tests
		[TestMethod]
		public void OperatorMultiplyShouldThrowExceptionIfNonNumericOperand()
		{
			Operator.Multiply.Apply(new CompileResult(1, DataType.Integer), new CompileResult("a", DataType.String));
		}

		[TestMethod]
		public void OperatorMultiplyShouldMultiplyNumericStringAndNumber()
		{
			var result = Operator.Multiply.Apply(new CompileResult(1, DataType.Integer), new CompileResult("3", DataType.String));
			Assert.AreEqual(3d, result.Result);
		}

		[TestMethod]
		public void OperatorMultiplyWithNonZeroIntegersReturnsCorrectResult()
		{
			var result = Operator.Multiply.Apply(new CompileResult(5, DataType.Integer), new CompileResult(8, DataType.Integer));
			Assert.AreEqual(40d, result.Result);
		}

		[TestMethod]
		public void OperatorMultiplyWithZeroIntegerReturnsZero()
		{
			var result = Operator.Multiply.Apply(new CompileResult(5, DataType.Integer), new CompileResult(0, DataType.Integer));
			Assert.AreEqual(0d, result.Result);
		}

		[TestMethod]
		public void OperatorMultiplyWithTwoNegativeIntegersReturnsCorrectValue()
		{
			var result = Operator.Multiply.Apply(new CompileResult(-5, DataType.Integer), new CompileResult(-8, DataType.Integer));
			Assert.AreEqual(40d, result.Result);
		}

		[TestMethod]
		public void OperatorMultiplyWithOneNegativeIntegerReturnsCorrectValue()
		{
			var result = Operator.Multiply.Apply(new CompileResult(-5, DataType.Integer), new CompileResult(8, DataType.Integer));
			Assert.AreEqual(-40d, result.Result);
		}

		[TestMethod]
		public void OperatorMultiplyWithDoublesReturnsCorrectValue()
		{
			var result = Operator.Multiply.Apply(new CompileResult(3.3, DataType.Decimal), new CompileResult(-5.6, DataType.Decimal));
			Assert.AreEqual(-18.48d, (double)result.Result, 0.000001);
		}

		[TestMethod]
		public void OperatorMultiplyWithFractionsReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["B1"].Formula = "(2/3) * (5/4)";
				ws.Calculate();
				Assert.AreEqual(0.83333333, (double)ws.Cells["B1"].Value, 0.000001);
			}
		}

		[TestMethod]
		public void OperatorMultiplyWithDateFunctionResultReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["B2"].Formula = "2 * DATE(2017,5,1)";
				ws.Calculate();
				Assert.AreEqual(85712d, ws.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void OperatorMultiplyWithDateAsStringReturnsCorrectValue()
		{
			var result = Operator.Multiply.Apply(new CompileResult(2, DataType.Integer), new CompileResult("5/1/2017", DataType.String));
			Assert.AreEqual(85712d, result.Result);
		}

		[TestMethod]
		public void OperatorMultiplyWithTwoRangesAsInputReturnsPoundValue()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["B1"].Value = 1;
				ws.Cells["B2"].Value = 2;
				ws.Cells["B3"].Value = 3;
				ws.Cells["B4"].Value = 4;
				ws.Cells["B5"].Value = 5;
				ws.Cells["B6"].Formula = "B1*B2:B4";
				ws.Calculate();
				Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)ws.Cells["B6"].Value).Type);
			}
		}

		[TestMethod]
		public void OperatorMultiplyWithNoArgumentsReturnsPoundValue()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["B1"].Value = 1;
				ws.Cells["B2"].Value = 1;
				ws.Cells["B3"].Value = 1;
				ws.Cells["B4"].Value = 1;
				ws.Cells["B5"].Value = 1;
				ws.Cells["B6"].Formula = "*";
				ws.Calculate();
				Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)ws.Cells["B6"].Value).Type);
			}
		}

		[TestMethod]
		public void OperatorMultiplyWithMaxInputsReturnsCorrectValue()
		{
			// The maximum number of inputs the function takes is 264.
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["B1"].Value = 2;
				ws.Cells["B2"].Value = 2;
				ws.Cells["B3"].Value = 2;
				ws.Cells["B4"].Value = 2;
				ws.Cells["B5"].Value = 2;
				ws.Cells["B6"].Value = 2;
				ws.Cells["B7"].Value = 2;
				ws.Cells["B8"].Value = 2;
				ws.Cells["B9"].Value = 2;
				ws.Cells["B10"].Value = 2;
				ws.Cells["B11"].Value = 2;
				ws.Cells["B12"].Value = 2;
				ws.Cells["B13"].Value = 2;
				ws.Cells["B14"].Value = 2;
				ws.Cells["B15"].Value = 2;
				ws.Cells["B16"].Value = 2;
				ws.Cells["B17"].Value = 2;
				ws.Cells["B18"].Value = 2;
				ws.Cells["B19"].Value = 2;
				ws.Cells["B20"].Value = 2;
				ws.Cells["B21"].Value = 2;
				ws.Cells["B22"].Value = 2;
				ws.Cells["B23"].Value = 2;
				ws.Cells["B24"].Value = 2;
				ws.Cells["B25"].Value = 2;
				ws.Cells["B26"].Value = 2;
				ws.Cells["B27"].Value = 2;
				ws.Cells["B28"].Value = 2;
				ws.Cells["B29"].Value = 2;
				ws.Cells["B30"].Value = 2;
				ws.Cells["B31"].Value = 2;
				ws.Cells["B32"].Value = 2;
				ws.Cells["B33"].Value = 2;
				ws.Cells["B34"].Value = 2;
				ws.Cells["B35"].Value = 2;
				ws.Cells["B36"].Value = 2;
				ws.Cells["B37"].Value = 2;
				ws.Cells["B38"].Value = 2;
				ws.Cells["B39"].Value = 2;
				ws.Cells["B40"].Value = 2;
				ws.Cells["B41"].Value = 2;
				ws.Cells["B42"].Value = 2;
				ws.Cells["B43"].Value = 2;
				ws.Cells["B44"].Value = 2;
				ws.Cells["B45"].Value = 2;
				ws.Cells["B46"].Value = 2;
				ws.Cells["B47"].Value = 2;
				ws.Cells["B48"].Value = 2;
				ws.Cells["B49"].Value = 2;
				ws.Cells["B50"].Value = 2;
				ws.Cells["B51"].Value = 2;
				ws.Cells["B52"].Value = 2;
				ws.Cells["B53"].Value = 2;
				ws.Cells["B54"].Value = 2;
				ws.Cells["B55"].Value = 2;
				ws.Cells["B56"].Value = 2;
				ws.Cells["B57"].Value = 2;
				ws.Cells["B58"].Value = 2;
				ws.Cells["B59"].Value = 2;
				ws.Cells["B60"].Value = 2;
				ws.Cells["B61"].Value = 2;
				ws.Cells["B62"].Value = 2;
				ws.Cells["B63"].Value = 2;
				ws.Cells["B64"].Value = 2;
				ws.Cells["B65"].Value = 2;
				ws.Cells["B66"].Value = 2;
				ws.Cells["B67"].Value = 2;
				ws.Cells["B68"].Value = 2;
				ws.Cells["B69"].Value = 2;
				ws.Cells["B70"].Value = 2;
				ws.Cells["B71"].Value = 2;
				ws.Cells["B72"].Value = 2;
				ws.Cells["B73"].Value = 2;
				ws.Cells["B74"].Value = 2;
				ws.Cells["B75"].Value = 2;
				ws.Cells["B76"].Value = 2;
				ws.Cells["B77"].Value = 2;
				ws.Cells["B78"].Value = 2;
				ws.Cells["B79"].Value = 2;
				ws.Cells["B80"].Value = 2;
				ws.Cells["B81"].Value = 2;
				ws.Cells["B82"].Value = 2;
				ws.Cells["B83"].Value = 2;
				ws.Cells["B84"].Value = 2;
				ws.Cells["B85"].Value = 2;
				ws.Cells["B86"].Value = 2;
				ws.Cells["B87"].Value = 2;
				ws.Cells["B88"].Value = 2;
				ws.Cells["B89"].Value = 2;
				ws.Cells["B90"].Value = 2;
				ws.Cells["B91"].Value = 2;
				ws.Cells["B92"].Value = 2;
				ws.Cells["B93"].Value = 2;
				ws.Cells["B94"].Value = 2;
				ws.Cells["B95"].Value = 2;
				ws.Cells["B96"].Value = 2;
				ws.Cells["B97"].Value = 2;
				ws.Cells["B98"].Value = 2;
				ws.Cells["B99"].Value = 2;
				ws.Cells["B100"].Value = 2;
				ws.Cells["B101"].Value = 2;
				ws.Cells["B102"].Value = 2;
				ws.Cells["B103"].Value = 2;
				ws.Cells["B104"].Value = 2;
				ws.Cells["B105"].Value = 2;
				ws.Cells["B106"].Value = 2;
				ws.Cells["B107"].Value = 2;
				ws.Cells["B108"].Value = 2;
				ws.Cells["B109"].Value = 2;
				ws.Cells["B110"].Value = 2;
				ws.Cells["B111"].Value = 2;
				ws.Cells["B112"].Value = 2;
				ws.Cells["B113"].Value = 2;
				ws.Cells["B114"].Value = 2;
				ws.Cells["B115"].Value = 2;
				ws.Cells["B116"].Value = 2;
				ws.Cells["B117"].Value = 2;
				ws.Cells["B118"].Value = 2;
				ws.Cells["B119"].Value = 2;
				ws.Cells["B120"].Value = 2;
				ws.Cells["B121"].Value = 2;
				ws.Cells["B122"].Value = 2;
				ws.Cells["B123"].Value = 2;
				ws.Cells["B124"].Value = 2;
				ws.Cells["B125"].Value = 2;
				ws.Cells["B126"].Value = 2;
				ws.Cells["B127"].Value = 2;
				ws.Cells["B128"].Value = 2;
				ws.Cells["B129"].Value = 2;
				ws.Cells["B130"].Value = 2;
				ws.Cells["B131"].Value = 2;
				ws.Cells["B132"].Value = 2;
				ws.Cells["B133"].Value = 2;
				ws.Cells["B134"].Value = 2;
				ws.Cells["B135"].Value = 2;
				ws.Cells["B136"].Value = 2;
				ws.Cells["B137"].Value = 2;
				ws.Cells["B138"].Value = 2;
				ws.Cells["B139"].Value = 2;
				ws.Cells["B140"].Value = 2;
				ws.Cells["B141"].Value = 2;
				ws.Cells["B142"].Value = 2;
				ws.Cells["B143"].Value = 2;
				ws.Cells["B144"].Value = 2;
				ws.Cells["B145"].Value = 2;
				ws.Cells["B146"].Value = 2;
				ws.Cells["B147"].Value = 2;
				ws.Cells["B148"].Value = 2;
				ws.Cells["B149"].Value = 2;
				ws.Cells["B150"].Value = 2;
				ws.Cells["B151"].Value = 2;
				ws.Cells["B152"].Value = 2;
				ws.Cells["B153"].Value = 2;
				ws.Cells["B154"].Value = 2;
				ws.Cells["B155"].Value = 2;
				ws.Cells["B156"].Value = 2;
				ws.Cells["B157"].Value = 2;
				ws.Cells["B158"].Value = 2;
				ws.Cells["B159"].Value = 2;
				ws.Cells["B160"].Value = 2;
				ws.Cells["B161"].Value = 2;
				ws.Cells["B162"].Value = 2;
				ws.Cells["B163"].Value = 2;
				ws.Cells["B164"].Value = 2;
				ws.Cells["B165"].Value = 2;
				ws.Cells["B166"].Value = 2;
				ws.Cells["B167"].Value = 2;
				ws.Cells["B168"].Value = 2;
				ws.Cells["B169"].Value = 2;
				ws.Cells["B170"].Value = 2;
				ws.Cells["B171"].Value = 2;
				ws.Cells["B172"].Value = 2;
				ws.Cells["B173"].Value = 2;
				ws.Cells["B174"].Value = 2;
				ws.Cells["B175"].Value = 2;
				ws.Cells["B176"].Value = 2;
				ws.Cells["B177"].Value = 2;
				ws.Cells["B178"].Value = 2;
				ws.Cells["B179"].Value = 2;
				ws.Cells["B180"].Value = 2;
				ws.Cells["B181"].Value = 2;
				ws.Cells["B182"].Value = 2;
				ws.Cells["B183"].Value = 2;
				ws.Cells["B184"].Value = 2;
				ws.Cells["B185"].Value = 2;
				ws.Cells["B186"].Value = 2;
				ws.Cells["B187"].Value = 2;
				ws.Cells["B188"].Value = 2;
				ws.Cells["B189"].Value = 2;
				ws.Cells["B190"].Value = 2;
				ws.Cells["B191"].Value = 2;
				ws.Cells["B192"].Value = 2;
				ws.Cells["B193"].Value = 2;
				ws.Cells["B194"].Value = 2;
				ws.Cells["B195"].Value = 2;
				ws.Cells["B196"].Value = 2;
				ws.Cells["B197"].Value = 2;
				ws.Cells["B198"].Value = 2;
				ws.Cells["B199"].Value = 2;
				ws.Cells["B200"].Value = 2;
				ws.Cells["B201"].Value = 2;
				ws.Cells["B202"].Value = 2;
				ws.Cells["B203"].Value = 2;
				ws.Cells["B204"].Value = 2;
				ws.Cells["B205"].Value = 2;
				ws.Cells["B206"].Value = 2;
				ws.Cells["B207"].Value = 2;
				ws.Cells["B208"].Value = 2;
				ws.Cells["B209"].Value = 2;
				ws.Cells["B210"].Value = 2;
				ws.Cells["B211"].Value = 2;
				ws.Cells["B212"].Value = 2;
				ws.Cells["B213"].Value = 2;
				ws.Cells["B214"].Value = 2;
				ws.Cells["B215"].Value = 2;
				ws.Cells["B216"].Value = 2;
				ws.Cells["B217"].Value = 2;
				ws.Cells["B218"].Value = 2;
				ws.Cells["B219"].Value = 2;
				ws.Cells["B220"].Value = 2;
				ws.Cells["B221"].Value = 2;
				ws.Cells["B222"].Value = 2;
				ws.Cells["B223"].Value = 2;
				ws.Cells["B224"].Value = 2;
				ws.Cells["B225"].Value = 2;
				ws.Cells["B226"].Value = 2;
				ws.Cells["B227"].Value = 2;
				ws.Cells["B228"].Value = 2;
				ws.Cells["B229"].Value = 2;
				ws.Cells["B230"].Value = 2;
				ws.Cells["B231"].Value = 2;
				ws.Cells["B232"].Value = 2;
				ws.Cells["B233"].Value = 2;
				ws.Cells["B234"].Value = 2;
				ws.Cells["B235"].Value = 2;
				ws.Cells["B236"].Value = 2;
				ws.Cells["B237"].Value = 2;
				ws.Cells["B238"].Value = 2;
				ws.Cells["B239"].Value = 2;
				ws.Cells["B240"].Value = 2;
				ws.Cells["B241"].Value = 2;
				ws.Cells["B242"].Value = 2;
				ws.Cells["B243"].Value = 2;
				ws.Cells["B244"].Value = 2;
				ws.Cells["B245"].Value = 2;
				ws.Cells["B246"].Value = 2;
				ws.Cells["B247"].Value = 2;
				ws.Cells["B248"].Value = 2;
				ws.Cells["B249"].Value = 2;
				ws.Cells["B250"].Value = 2;
				ws.Cells["B251"].Value = 2;
				ws.Cells["B252"].Value = 2;
				ws.Cells["B253"].Value = 2;
				ws.Cells["B254"].Value = 2;
				ws.Cells["B255"].Value = 2;
				ws.Cells["B256"].Value = 2;
				ws.Cells["B257"].Value = 2;
				ws.Cells["B258"].Value = 2;
				ws.Cells["B259"].Value = 2;
				ws.Cells["B260"].Value = 2;
				ws.Cells["B261"].Value = 2;
				ws.Cells["B262"].Value = 2;
				ws.Cells["B263"].Value = 2;
				ws.Cells["B264"].Value = 2;
				ws.Cells["C1"].Formula = "B1* B2* B3* B4* B5* B6* B7* B8* B9* B10* B11* B12* B13* B14* B15* B16* B17* B18* B19* B20* B21* B22* B23* B24* B25* B26* B27* B28* B29* B30* B31* B32* B33* B34* B35* B36* B37* B38* B39* B40* B41* B42* B43* B44* B45* B46* B47* B48* B49* B50* B51* B52* B53* B54* B55* B56* B57* B58* B59* B60* B61* B62* B63* B64* B65* B66* B67* B68* B69* B70* B71* B72* B73* B74* B75* B76* B77* B78* B79* B80* B81* B82* B83* B84* B85* B86* B87* B88* B89* B90* B91* B92* B93* B94* B95* B96* B97* B98* B99* B100* B101* B102* B103* B104* B105* B106* B107* B108* B109* B110* B111* B112* B113* B114* B115* B116* B117* B118* B119* B120* B121* B122* B123* B124* B125* B126* B127* B128* B129* B130* B131* B132* B133* B134* B135* B136* B137* B138* B139* B140* B141* B142* B143* B144* B145* B146* B147* B148* B149* B150* B151* B152* B153* B154* B155* B156* B157* B158* B159* B160* B170* B171* B172* B173* B174* B175* B176* B187* B179* B180* B181* B182* B183* B184* B185* B186* B187* B188* B189* B190* B191* B192* B193* B194* B194* B195* B196* B197* B198* B199* B200* B201* B202* B203* B204* B205* B206* B207* B208* B209* B210* B211* B212* B213* B214* B215* B216* B217* B218* B219* B220* B221* B222* B223* B224* B225* B226* B227* B228* B229* B230* B231* B232* B233* B234* B235* B236* B237* B238* B239* B240* B241* B242* B243* B244* B245* B246* B247* B248* B249* B250* B251* B252* B253* B254* B255* B256* B257* B258* B259* B260* B261* B262* B263* B264";
				ws.Calculate();
				Assert.AreEqual(System.Math.Pow(2, 255), ws.Cells["C1"].Value);
			}
		}
		#endregion

		#region Operator Concat Tests
		[TestMethod]
		public void OperatorConcatShouldConcatTwoStrings()
		{
			var result = Operator.Concat.Apply(new CompileResult("a", DataType.String), new CompileResult("b", DataType.String));
			Assert.AreEqual("ab", result.Result);
		}

		[TestMethod]
		public void OperatorConcatShouldConcatANumberAndAString()
		{
			var result = Operator.Concat.Apply(new CompileResult(12, DataType.Integer), new CompileResult("b", DataType.String));
			Assert.AreEqual("12b", result.Result);
		}

		[TestMethod]
		public void OperatorConcatShouldConcatAnEmptyRange()
		{
			var file = new FileInfo("filename.xlsx");
			using (var package = new ExcelPackage(file))
			using (var sheet = package.Workbook.Worksheets.Add("NewSheet"))
			using (var excelDataProvider = new EpplusExcelDataProvider(package))
			{
				var emptyRange = excelDataProvider.GetRange("NewSheet", 2, 2, "B2");
				var result = Operator.Concat.Apply(new CompileResult(emptyRange, DataType.ExcelAddress), new CompileResult("b", DataType.String));
				Assert.AreEqual("b", result.Result);
				result = Operator.Concat.Apply(new CompileResult("b", DataType.String), new CompileResult(emptyRange, DataType.ExcelAddress));
				Assert.AreEqual("b", result.Result);
			}
		}
		#endregion

		#region Operator Equals Tests
		[TestMethod]
		public void OperatorEqShouldReturnTruefSuppliedValuesAreEqual()
		{
			var result = Operator.EqualsTo.Apply(new CompileResult(12, DataType.Integer), new CompileResult(12, DataType.Integer));
			Assert.IsTrue((bool)result.Result);
		}

		[TestMethod]
		public void OperatorEqShouldReturnFalsefSuppliedValuesDiffer()
		{
			var result = Operator.EqualsTo.Apply(new CompileResult(11, DataType.Integer), new CompileResult(12, DataType.Integer));
			Assert.IsFalse((bool)result.Result);
		}
		#endregion

		#region Operator NotEqualsTo Tests
		[TestMethod]
		public void OperatorNotEqualToShouldReturnTruefSuppliedValuesDiffer()
		{
			var result = Operator.NotEqualsTo.Apply(new CompileResult(11, DataType.Integer), new CompileResult(12, DataType.Integer));
			Assert.IsTrue((bool)result.Result);
		}

		[TestMethod]
		public void OperatorNotEqualToShouldReturnFalsefSuppliedValuesAreEqual()
		{
			var result = Operator.NotEqualsTo.Apply(new CompileResult(11, DataType.Integer), new CompileResult(11, DataType.Integer));
			Assert.IsFalse((bool)result.Result);
		}
		#endregion

		#region Operator GreaterThan Tests
		[TestMethod]
		public void OperatorGreaterThanToShouldReturnTrueIfLeftIsSetAndRightIsNull()
		{
			var result = Operator.GreaterThan.Apply(new CompileResult(11, DataType.Integer), new CompileResult(null, DataType.Empty));
			Assert.IsTrue((bool)result.Result);
		}

		[TestMethod]
		public void OperatorGreaterThanToShouldReturnTrueIfLeftIs11AndRightIs10()
		{
			var result = Operator.GreaterThan.Apply(new CompileResult(11, DataType.Integer), new CompileResult(10, DataType.Integer));
			Assert.IsTrue((bool)result.Result);
		}
		#endregion

		#region Operator Exp Tests
		[TestMethod]
		public void OperatorExpShouldReturnCorrectResult()
		{
			var result = Operator.Exp.Apply(new CompileResult(2, DataType.Integer), new CompileResult(3, DataType.Integer));
			Assert.AreEqual(8d, result.Result);
		}
		#endregion

		#region Numeric and Date String Comparison Tests
		[TestMethod]
		public void OperatorsActingOnNumericStrings()
		{
			double number1 = 42.0;
			double number2 = -143.75;
			CompileResult result1 = new CompileResult(number1.ToString("n"), DataType.String);
			CompileResult result2 = new CompileResult(number2.ToString("n"), DataType.String);
			var operatorResult = Operator.Concat.Apply(result1, result2);
			Assert.AreEqual($"{number1.ToString("n")}{number2.ToString("n")}", operatorResult.Result);
			operatorResult = Operator.Divide.Apply(result1, result2);
			Assert.AreEqual(number1 / number2, operatorResult.Result);
			operatorResult = Operator.Exp.Apply(result1, result2);
			Assert.AreEqual(Math.Pow(number1, number2), operatorResult.Result);
			operatorResult = Operator.Minus.Apply(result1, result2);
			Assert.AreEqual(number1 - number2, operatorResult.Result);
			operatorResult = Operator.Multiply.Apply(result1, result2);
			Assert.AreEqual(number1 * number2, operatorResult.Result);
			operatorResult = Operator.Percent.Apply(result1, result2);
			Assert.AreEqual(number1 * number2, operatorResult.Result);
			operatorResult = Operator.Plus.Apply(result1, result2);
			Assert.AreEqual(number1 + number2, operatorResult.Result);
			// Comparison operators always compare string-wise and don't parse out the actual numbers.
			operatorResult = Operator.NotEqualsTo.Apply(result1, new CompileResult(number1.ToString("n0"), DataType.String));
			Assert.IsTrue((bool)operatorResult.Result);
			operatorResult = Operator.EqualsTo.Apply(result1, new CompileResult(number1.ToString("n0"), DataType.String));
			Assert.IsFalse((bool)operatorResult.Result);
			operatorResult = Operator.GreaterThan.Apply(result1, result2);
			Assert.IsTrue((bool)operatorResult.Result);
			operatorResult = Operator.GreaterThanOrEqual.Apply(result1, result2);
			Assert.IsTrue((bool)operatorResult.Result);
			operatorResult = Operator.LessThan.Apply(result1, result2);
			Assert.IsFalse((bool)operatorResult.Result);
			operatorResult = Operator.LessThanOrEqual.Apply(result1, result2);
			Assert.IsFalse((bool)operatorResult.Result);
		}

		[TestMethod]
		public void OperatorsActingOnDateStrings()
		{
			const string dateFormat = "M-dd-yyyy";
			DateTime date1 = new DateTime(2015, 2, 20);
			DateTime date2 = new DateTime(2015, 12, 1);
			var numericDate1 = date1.ToOADate();
			var numericDate2 = date2.ToOADate();
			CompileResult result1 = new CompileResult(date1.ToString(dateFormat), DataType.String); // 2/20/2015
			CompileResult result2 = new CompileResult(date2.ToString(dateFormat), DataType.String); // 12/1/2015
			var operatorResult = Operator.Concat.Apply(result1, result2);
			Assert.AreEqual($"{date1.ToString(dateFormat)}{date2.ToString(dateFormat)}", operatorResult.Result);
			operatorResult = Operator.Divide.Apply(result1, result2);
			Assert.AreEqual(numericDate1 / numericDate2, operatorResult.Result);
			operatorResult = Operator.Exp.Apply(result1, result2);
			Assert.AreEqual(Math.Pow(numericDate1, numericDate2), operatorResult.Result);
			operatorResult = Operator.Minus.Apply(result1, result2);
			Assert.AreEqual(numericDate1 - numericDate2, operatorResult.Result);
			operatorResult = Operator.Multiply.Apply(result1, result2);
			Assert.AreEqual(numericDate1 * numericDate2, operatorResult.Result);
			operatorResult = Operator.Percent.Apply(result1, result2);
			Assert.AreEqual(numericDate1 * numericDate2, operatorResult.Result);
			operatorResult = Operator.Plus.Apply(result1, result2);
			Assert.AreEqual(numericDate1 + numericDate2, operatorResult.Result);
			// Comparison operators always compare string-wise and don't parse out the actual numbers.
			operatorResult = Operator.EqualsTo.Apply(result1, new CompileResult(date1.ToString("f"), DataType.String));
			Assert.IsFalse((bool)operatorResult.Result);
			operatorResult = Operator.NotEqualsTo.Apply(result1, new CompileResult(date1.ToString("f"), DataType.String));
			Assert.IsTrue((bool)operatorResult.Result);
			operatorResult = Operator.GreaterThan.Apply(result1, result2);
			Assert.IsTrue((bool)operatorResult.Result);
			operatorResult = Operator.GreaterThanOrEqual.Apply(result1, result2);
			Assert.IsTrue((bool)operatorResult.Result);
			operatorResult = Operator.LessThan.Apply(result1, result2);
			Assert.IsFalse((bool)operatorResult.Result);
			operatorResult = Operator.LessThanOrEqual.Apply(result1, result2);
			Assert.IsFalse((bool)operatorResult.Result);
		}

		[TestMethod]
		public void OperatorsActingOnGermanDateStrings()
		{
			var currentCulture = CultureInfo.CurrentCulture;
			var culture = new CultureInfo("de-DE");
			Thread.CurrentThread.CurrentCulture = culture;
			try
			{
				string dateFormat = culture.DateTimeFormat.ShortDatePattern;
				DateTime date1 = new DateTime(2015, 2, 20);
				DateTime date2 = new DateTime(2015, 12, 1);
				var numericDate1 = date1.ToOADate();
				var numericDate2 = date2.ToOADate();
				CompileResult result1 = new CompileResult(date1.ToString(dateFormat), DataType.String); // 20.02.2015
				CompileResult result2 = new CompileResult(date2.ToString(dateFormat), DataType.String); // 01.12.2015
				var operatorResult = Operator.Concat.Apply(result1, result2);
				Assert.AreEqual($"{date1.ToString(dateFormat)}{date2.ToString(dateFormat)}", operatorResult.Result);
				operatorResult = Operator.Divide.Apply(result1, result2);
				Assert.AreEqual(numericDate1 / numericDate2, operatorResult.Result);
				operatorResult = Operator.Exp.Apply(result1, result2);
				Assert.AreEqual(Math.Pow(numericDate1, numericDate2), operatorResult.Result);
				operatorResult = Operator.Minus.Apply(result1, result2);
				Assert.AreEqual(numericDate1 - numericDate2, operatorResult.Result);
				operatorResult = Operator.Multiply.Apply(result1, result2);
				Assert.AreEqual(numericDate1 * numericDate2, operatorResult.Result);
				operatorResult = Operator.Percent.Apply(result1, result2);
				Assert.AreEqual(numericDate1 * numericDate2, operatorResult.Result);
				operatorResult = Operator.Plus.Apply(result1, result2);
				Assert.AreEqual(numericDate1 + numericDate2, operatorResult.Result);
				// Comparison operators always compare strings string-wise and don't parse out the actual numbers.
				operatorResult = Operator.EqualsTo.Apply(result1, new CompileResult(date1.ToString("f"), DataType.String));
				Assert.IsFalse((bool)operatorResult.Result);
				operatorResult = Operator.NotEqualsTo.Apply(result1, new CompileResult(date1.ToString("f"), DataType.String));
				Assert.IsTrue((bool)operatorResult.Result);
				operatorResult = Operator.GreaterThan.Apply(result1, result2);
				Assert.IsTrue((bool)operatorResult.Result);
				operatorResult = Operator.GreaterThanOrEqual.Apply(result1, result2);
				Assert.IsTrue((bool)operatorResult.Result);
				operatorResult = Operator.LessThan.Apply(result1, result2);
				Assert.IsFalse((bool)operatorResult.Result);
				operatorResult = Operator.LessThanOrEqual.Apply(result1, result2);
				Assert.IsFalse((bool)operatorResult.Result);
			}
			finally
			{
				Thread.CurrentThread.CurrentCulture = currentCulture;
			}
		}
		#endregion
	}
}
