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
		[TestMethod]
		public void OperatorGreaterThanNumbersAreLessThanTextValuesAreLessThanLogicalValues()
		{
			Assert.AreEqual(false, Operator.GreaterThan.Apply(new CompileResult(1000000, DataType.Integer), new CompileResult("text", DataType.String)).Result);
			Assert.AreEqual(false, Operator.GreaterThan.Apply(new CompileResult(1000000, DataType.Integer), new CompileResult("1", DataType.String)).Result);
			Assert.AreEqual(false, Operator.GreaterThan.Apply(new CompileResult(1000000, DataType.Integer), new CompileResult(true, DataType.Boolean)).Result);
			Assert.AreEqual(false, Operator.GreaterThan.Apply(new CompileResult(1000000, DataType.Integer), new CompileResult(false, DataType.Boolean)).Result);
			Assert.AreEqual(true, Operator.GreaterThan.Apply(new CompileResult(true, DataType.Boolean), new CompileResult("10000", DataType.String)).Result);
		}

		// Numbers to numbers
		// Numbers to text
		// Numbers to logicals
		// Logicals to text
		// Logicals to logicals
		// text to text

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

		[TestMethod]
		public void OperatorMultiplyShouldThrowExceptionIfNonNumericOperand()
		{
			Operator.Multiply.Apply(new CompileResult(1, DataType.Integer), new CompileResult("a", DataType.String));
		}

		[TestMethod]
		public void OperatoMultiplyShouldMultiplyNumericStringAndNumber()
		{
			var result = Operator.Multiply.Apply(new CompileResult(1, DataType.Integer), new CompileResult("3", DataType.String));
			Assert.AreEqual(3d, result.Result);
		}

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

		[TestMethod]
		public void OperatorEqShouldReturnTruefSuppliedValuesAreEqual()
		{
			var result = Operator.Eq.Apply(new CompileResult(12, DataType.Integer), new CompileResult(12, DataType.Integer));
			Assert.IsTrue((bool)result.Result);
		}

		[TestMethod]
		public void OperatorEqShouldReturnFalsefSuppliedValuesDiffer()
		{
			var result = Operator.Eq.Apply(new CompileResult(11, DataType.Integer), new CompileResult(12, DataType.Integer));
			Assert.IsFalse((bool)result.Result);
		}

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

		[TestMethod]
		public void OperatorExpShouldReturnCorrectResult()
		{
			var result = Operator.Exp.Apply(new CompileResult(2, DataType.Integer), new CompileResult(3, DataType.Integer));
			Assert.AreEqual(8d, result.Result);
		}

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
			operatorResult = Operator.Eq.Apply(result1, new CompileResult(number1.ToString("n0"), DataType.String));
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
			operatorResult = Operator.Eq.Apply(result1, new CompileResult(date1.ToString("f"), DataType.String));
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
				// Comparison operators always compare string-wise and don't parse out the actual numbers.
				operatorResult = Operator.Eq.Apply(result1, new CompileResult(date1.ToString("f"), DataType.String));
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
	}
}
