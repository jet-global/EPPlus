using System;
using System.Globalization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace EPPlusTest.FormulaParsing.ExpressionGraph
{
	[TestClass]
	public class ExpressionConverterTests
	{
		#region ToStringExpression Tests
		[TestMethod]
		public void ToStringExpressionShouldConvertIntegerExpressionToStringExpression()
		{
			var integerExpression = new IntegerExpression("2");
			var result = new ExpressionConverter().ToStringExpression(integerExpression);
			Assert.IsInstanceOfType(result, typeof(StringExpression));
			Assert.AreEqual("2", result.Compile().Result);
		}

		[TestMethod]
		public void ToStringExpressionShouldCopyOperatorToStringExpression()
		{
			var integerExpression = new IntegerExpression("2") { Operator = Operator.Plus };
			var result = new ExpressionConverter().ToStringExpression(integerExpression);
			Assert.AreEqual(integerExpression.Operator, result.Operator);
		}

		[TestMethod]
		public void ToStringExpressionShouldConvertDecimalExpressionToStringExpression()
		{
			var decimalExpression = new DecimalExpression("2.5");
			var result = new ExpressionConverter().ToStringExpression(decimalExpression);
			Assert.IsInstanceOfType(result, typeof(StringExpression));
			Assert.AreEqual($"2{CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator}5", result.Compile().Result);
		}
		#endregion

		#region FromCompileResult Tests
		[TestMethod]
		public void FromCompileResultShouldCreateIntegerExpressionIfCompileResultIsInteger()
		{
			var compileResult = new CompileResult(1, DataType.Integer);
			var result = new ExpressionConverter().FromCompileResult(compileResult);
			Assert.IsInstanceOfType(result, typeof(IntegerExpression));
			Assert.AreEqual(1d, result.Compile().Result);
		}

		[TestMethod]
		public void FromCompileResultShouldCreateStringExpressionIfCompileResultIsString()
		{
			var compileResult = new CompileResult("abc", DataType.String);
			var result = new ExpressionConverter().FromCompileResult(compileResult);
			Assert.IsInstanceOfType(result, typeof(StringExpression));
			Assert.AreEqual("abc", result.Compile().Result);
		}

		[TestMethod]
		public void FromCompileResultShouldCreateDecimalExpressionIfCompileResultIsDouble()
		{
			var compileResult = new CompileResult(2.5d, DataType.Decimal);
			var result = new ExpressionConverter().FromCompileResult(compileResult);
			Assert.IsInstanceOfType(result, typeof(DecimalExpression));
			Assert.AreEqual(2.5d, result.Compile().Result);
		}

		[TestMethod]
		public void FromCompileResultShouldCreateDecimalExpressionIfCompileResultIsDecimal()
		{
			decimal input = 2.5m;
			double expected = Convert.ToDouble(input);
			var compileResult = new CompileResult(input, DataType.Decimal);
			var result = new ExpressionConverter().FromCompileResult(compileResult);
			Assert.IsInstanceOfType(result, typeof(DecimalExpression));
			Assert.AreEqual(expected, result.Compile().Result);
		}

		[TestMethod]
		public void FromCompileResultShouldCreateBooleanExpressionIfCompileResultIsBoolean()
		{
			var compileResult = new CompileResult("true", DataType.Boolean);
			var result = new ExpressionConverter().FromCompileResult(compileResult);
			Assert.IsInstanceOfType(result, typeof(BooleanExpression));
			Assert.IsTrue((bool)result.Compile().Result);
		}
		#endregion
	}
}
