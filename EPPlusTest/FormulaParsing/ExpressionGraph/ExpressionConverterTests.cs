using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace EPPlusTest.FormulaParsing.ExpressionGraph
{
	[TestClass]
	public class ExpressionConverterTests
	{
		#region FromCompileResult Tests
		[TestMethod]
		public void FromCompileResultInteger()
		{
			var compileResult = new CompileResult(1, DataType.Integer);
			var result = new ExpressionConverter().FromCompileResult(compileResult);
			Assert.IsInstanceOfType(result, typeof(IntegerExpression));
			Assert.AreEqual(1d, result.Compile().Result);
		}

		[TestMethod]
		public void FromCompileResultIntegerString()
		{
			var compileResult = new CompileResult("1", DataType.Integer);
			var result = new ExpressionConverter().FromCompileResult(compileResult);
			Assert.IsInstanceOfType(result, typeof(IntegerExpression));
			Assert.AreEqual(1d, result.Compile().Result);
		}
		[TestMethod]
		public void FromCompileResultIntegerUnknown()
		{
			var compileResult = new CompileResult(1, DataType.Unknown);
			var result = new ExpressionConverter().FromCompileResult(compileResult);
			Assert.IsInstanceOfType(result, typeof(IntegerExpression));
			Assert.AreEqual(1d, result.Compile().Result);
		}

		[TestMethod]
		public void FromCompileResultTime()
		{
			var compileResult = new CompileResult(1000, DataType.Time);
			var result = new ExpressionConverter().FromCompileResult(compileResult);
			Assert.IsInstanceOfType(result, typeof(DecimalExpression));
			Assert.AreEqual(1000d, result.Compile().Result);
		}

		[TestMethod]
		public void FromCompileResultTimeString()
		{
			var compileResult = new CompileResult("1000", DataType.Time);
			var result = new ExpressionConverter().FromCompileResult(compileResult);
			Assert.IsInstanceOfType(result, typeof(DecimalExpression));
			Assert.AreEqual(1000d, result.Compile().Result);
		}

		[TestMethod]
		public void FromCompileResultDecimal()
		{
			var compileResult = new CompileResult(2.5, DataType.Decimal);
			var result = new ExpressionConverter().FromCompileResult(compileResult);
			Assert.IsInstanceOfType(result, typeof(DecimalExpression));
			Assert.AreEqual(2.5d, result.Compile().Result);
		}

		[TestMethod]
		public void FromCompileResultDecimalString()
		{
			var compileResult = new CompileResult("2.5", DataType.Decimal);
			var result = new ExpressionConverter().FromCompileResult(compileResult);
			Assert.IsInstanceOfType(result, typeof(DecimalExpression));
			Assert.AreEqual(2.5d, result.Compile().Result);
		}

		[TestMethod]
		public void FromCompileResultDecimalUnknown()
		{
			var compileResult = new CompileResult(2.5, DataType.Unknown);
			var result = new ExpressionConverter().FromCompileResult(compileResult);
			Assert.IsInstanceOfType(result, typeof(DecimalExpression));
			Assert.AreEqual(2.5d, result.Compile().Result);
		}

		[TestMethod]
		public void FromCompileResultString()
		{
			var compileResult = new CompileResult("abc", DataType.String);
			var result = new ExpressionConverter().FromCompileResult(compileResult);
			Assert.IsInstanceOfType(result, typeof(StringExpression));
			Assert.AreEqual("abc", result.Compile().Result);
		}

		[TestMethod]
		public void FromCompileResultStringFromNumber()
		{
			var compileResult = new CompileResult(1000, DataType.String);
			var result = new ExpressionConverter().FromCompileResult(compileResult);
			Assert.IsInstanceOfType(result, typeof(StringExpression));
			Assert.AreEqual("1000", result.Compile().Result);
		}

		[TestMethod]
		public void FromCompileResultStringUnknown()
		{
			var compileResult = new CompileResult("abc", DataType.Unknown);
			var result = new ExpressionConverter().FromCompileResult(compileResult);
			Assert.IsInstanceOfType(result, typeof(StringExpression));
			Assert.AreEqual("abc", result.Compile().Result);
		}

		[TestMethod]
		public void FromCompileResultBoolean()
		{
			var compileResult = new CompileResult(true, DataType.Boolean);
			var result = new ExpressionConverter().FromCompileResult(compileResult);
			Assert.IsInstanceOfType(result, typeof(BooleanExpression));
			Assert.AreEqual(true, result.Compile().Result);
		}

		[TestMethod]
		public void FromCompileResultBooleanString()
		{
			var compileResult = new CompileResult("false", DataType.Boolean);
			var result = new ExpressionConverter().FromCompileResult(compileResult);
			Assert.IsInstanceOfType(result, typeof(BooleanExpression));
			Assert.AreEqual(false, result.Compile().Result);
		}

		[TestMethod]
		public void FromCompileResultBooleanUnknown()
		{
			var compileResult = new CompileResult(true, DataType.Unknown);
			var result = new ExpressionConverter().FromCompileResult(compileResult);
			Assert.IsInstanceOfType(result, typeof(BooleanExpression));
			Assert.AreEqual(true, result.Compile().Result);
		}

		[TestMethod]
		public void FromCompileResultDate()
		{
			var compileResult = new CompileResult(DateTime.Parse("1/1/2018"), DataType.Date);
			var result = new ExpressionConverter().FromCompileResult(compileResult);
			Assert.IsInstanceOfType(result, typeof(DateExpression));
			Assert.AreEqual(DateTime.Parse("1/1/2018"), result.Compile().Result);
		}

		[TestMethod]
		public void FromCompileResultDateString()
		{
			var compileResult = new CompileResult("1/1/2018", DataType.Date);
			var result = new ExpressionConverter().FromCompileResult(compileResult);
			Assert.IsInstanceOfType(result, typeof(DateExpression));
			Assert.AreEqual(DateTime.Parse("1/1/2018"), result.Compile().Result);
		}

		[TestMethod]
		public void FromCompileResultDateOADate()
		{
			var compileResult = new CompileResult(12345, DataType.Date);
			var result = new ExpressionConverter().FromCompileResult(compileResult);
			Assert.IsInstanceOfType(result, typeof(DateExpression));
			Assert.AreEqual(DateTime.Parse("10/18/1933"), result.Compile().Result);
		}

		[TestMethod]
		public void FromCompileResultDateOADateString()
		{
			var compileResult = new CompileResult("12345", DataType.Date);
			var result = new ExpressionConverter().FromCompileResult(compileResult);
			Assert.IsInstanceOfType(result, typeof(DateExpression));
			Assert.AreEqual(DateTime.Parse("10/18/1933"), result.Compile().Result);
		}

		[TestMethod]
		public void FromCompileResultDateUnknown()
		{
			var compileResult = new CompileResult(DateTime.Parse("1/1/2018"), DataType.Unknown);
			var result = new ExpressionConverter().FromCompileResult(compileResult);
			Assert.IsInstanceOfType(result, typeof(DateExpression));
			Assert.AreEqual(DateTime.Parse("1/1/2018"), result.Compile().Result);
		}

		[TestMethod]
		public void FromCompileResultExcelError()
		{
			var compileResult = new CompileResult(ExcelErrorValue.Parse("#N/A"), DataType.ExcelError);
			var result = new ExpressionConverter().FromCompileResult(compileResult);
			Assert.IsInstanceOfType(result, typeof(ExcelErrorExpression));
			Assert.AreEqual(ExcelErrorValue.Parse("#N/A"), result.Compile().Result);
		}

		[TestMethod]
		public void FromCompileResultExcelErrorString()
		{
			var compileResult = new CompileResult("#NAME?", DataType.ExcelError);
			var result = new ExpressionConverter().FromCompileResult(compileResult);
			Assert.IsInstanceOfType(result, typeof(ExcelErrorExpression));
			Assert.AreEqual(ExcelErrorValue.Parse("#NAME?"), result.Compile().Result);
		}

		[TestMethod]
		public void FromCompileResultErrorTypeEnum()
		{
			var compileResult = new CompileResult(eErrorType.Value, DataType.ExcelError);
			var result = new ExpressionConverter().FromCompileResult(compileResult);
			Assert.IsInstanceOfType(result, typeof(ExcelErrorExpression));
			Assert.AreEqual(ExcelErrorValue.Parse("#VALUE!"), result.Compile().Result);
		}

		[TestMethod]
		public void FromCompileResultExcelErrorUnknown()
		{
			var compileResult = new CompileResult(ExcelErrorValue.Parse("#N/A"), DataType.Unknown);
			var result = new ExpressionConverter().FromCompileResult(compileResult);
			Assert.IsInstanceOfType(result, typeof(ExcelErrorExpression));
			Assert.AreEqual(ExcelErrorValue.Parse("#N/A"), result.Compile().Result);
		}

		[TestMethod]
		public void FromCompileResultEmpty()
		{
			var compileResult = new CompileResult("whatever", DataType.Empty);
			var result = new ExpressionConverter().FromCompileResult(compileResult);
			Assert.IsInstanceOfType(result, typeof(IntegerExpression));
			Assert.AreEqual(0d, result.Compile().Result);
		}

		[TestMethod]
		public void FromCompileResultEmptyUnknown()
		{
			var compileResult = new CompileResult(null, DataType.Unknown);
			var result = new ExpressionConverter().FromCompileResult(compileResult);
			Assert.IsInstanceOfType(result, typeof(IntegerExpression));
			Assert.AreEqual(0d, result.Compile().Result);
		}

		[TestMethod]
		public void FromCompileResultExcelAddress()
		{
			var compileResult = new CompileResult("Sheet1!C3", DataType.ExcelAddress);
			var result = new ExpressionConverter().FromCompileResult(compileResult);
			Assert.IsInstanceOfType(result, typeof(StringExpression));
			Assert.AreEqual("Sheet1!C3", result.Compile().Result);
		}

		[TestMethod]
		public void FromCompileResultEnumerableInteger()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet1");
				sheet.Cells["C1"].Value = 1;
				sheet.Cells["C2"].Value = 2;
				sheet.Cells["C3"].Value = 3;
				var dataProvider = new EpplusExcelDataProvider(package);
				var range = dataProvider.GetRange("Sheet1", 1, 3, 3, 3);
				var compileResult = new CompileResult(range, DataType.Enumerable);
				var result = new ExpressionConverter().FromCompileResult(compileResult);
				Assert.IsInstanceOfType(result, typeof(IntegerExpression));
				Assert.AreEqual(1d, result.Compile().Result);
			}
		}

		[TestMethod]
		public void FromCompileResultEnumerableString()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet1");
				sheet.Cells["C1"].Value = "1";
				sheet.Cells["C2"].Value = "2";
				sheet.Cells["C3"].Value = "3";
				var dataProvider = new EpplusExcelDataProvider(package);
				var range = dataProvider.GetRange("Sheet1", 1, 3, 3, 3);
				var compileResult = new CompileResult(range, DataType.Enumerable);
				var result = new ExpressionConverter().FromCompileResult(compileResult);
				Assert.IsInstanceOfType(result, typeof(StringExpression));
				Assert.AreEqual("1", result.Compile().Result);
			}
		}

		[TestMethod]
		public void FromCompileResultEnumerableIntegerUnknown()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet1");
				sheet.Cells["C1"].Value = 1;
				sheet.Cells["C2"].Value = 2;
				sheet.Cells["C3"].Value = 3;
				var dataProvider = new EpplusExcelDataProvider(package);
				var range = dataProvider.GetRange("Sheet1", 1, 3, 3, 3);
				var compileResult = new CompileResult(range, DataType.Unknown);
				var result = new ExpressionConverter().FromCompileResult(compileResult);
				Assert.IsInstanceOfType(result, typeof(IntegerExpression));
				Assert.AreEqual(1d, result.Compile().Result);
			}
		}

		[TestMethod]
		public void FromCompileResultEnumerableList()
		{
			var compileResult = new CompileResult(new List<object> { 1, 2, 3 }, DataType.Enumerable);
			var result = new ExpressionConverter().FromCompileResult(compileResult);
			Assert.IsInstanceOfType(result, typeof(IntegerExpression));
			Assert.AreEqual(1d, result.Compile().Result);
		}

		[TestMethod]
		public void FromCompileResultEnumerableListUnknown()
		{
			var compileResult = new CompileResult(new List<object> { 1, 2, 3 }, DataType.Unknown);
			var result = new ExpressionConverter().FromCompileResult(compileResult);
			Assert.IsInstanceOfType(result, typeof(IntegerExpression));
			Assert.AreEqual(1d, result.Compile().Result);
		}
		#endregion
	}
}
