using System.Collections.Generic;
using System.Globalization;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace EPPlusTest.FormulaParsing.ExpressionGraph
{
	[TestClass]
	public class CompileResultFactoryTests
	{
		[TestMethod]
		public void CalculateUsingEuropeanDates()
		{
			var currentCulture = Thread.CurrentThread.CurrentCulture;
			try
			{
				var us = CultureInfo.CreateSpecificCulture("en-US");
				Thread.CurrentThread.CurrentCulture = us;
				var crf = new CompileResultFactory();
				var result = crf.Create("1/15/2014");
				var numeric = result.ResultNumeric;
				Assert.AreEqual(41654, numeric);
				var gb = CultureInfo.CreateSpecificCulture("en-GB");
				Thread.CurrentThread.CurrentCulture = gb;
				var euroResult = crf.Create("15/1/2014");
				var eNumeric = euroResult.ResultNumeric;
				Assert.AreEqual(41654, eNumeric);
			}
			finally
			{
				Thread.CurrentThread.CurrentCulture = currentCulture;
			}
		}

		[TestMethod]
		public void CreateErrorType()
		{
			var factory = new CompileResultFactory();
			Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), factory.Create("#VALUE!").Result);
			Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Name), factory.Create("#NAME?").Result);
			Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Div0), factory.Create("#DIV/0!").Result);
			Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), factory.Create("#N/A").Result);
			Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Null), factory.Create("#NULL!").Result);
			Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), factory.Create("#NUM!").Result);
			Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Ref), factory.Create("#REF!").Result);
		}

		[TestMethod]
		public void CreateEnumerableResult()
		{
			IEnumerable<object> enumerableObj = new List<string> { "asasd", "wrtdgff", "sdfsfds" };
			CompileResult compileResult = new CompileResultFactory().Create(enumerableObj);
			Assert.AreEqual(enumerableObj, compileResult.Result);
			Assert.AreEqual(DataType.Enumerable, compileResult.DataType);
		}

		[TestMethod]
		public void CreateListEnumerableResult()
		{
			List<string> listEnumerableObj = new List<string> { "asasd", "wrtdgff", "sdfsfds" };
			CompileResult compileResult = new CompileResultFactory().Create(listEnumerableObj);
			Assert.AreEqual(listEnumerableObj, compileResult.Result);
			Assert.AreEqual(DataType.Enumerable, compileResult.DataType);
		}

		[TestMethod]
		public void CreateStringResult()
		{
			string value = "SomeStringTextHere";
			CompileResult compileResult = new CompileResultFactory().Create(value);
			Assert.AreEqual(value, compileResult.Result);
			Assert.AreEqual(DataType.String, compileResult.DataType);
		}

		[TestMethod]
		public void CreateValueErrorStringResult()
		{
			string value = "#VALUE!";
			CompileResult compileResult = new CompileResultFactory().Create(value);
			Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), compileResult.Result);
			Assert.AreEqual(DataType.ExcelError, compileResult.DataType);
		}

		[TestMethod]
		public void CreateNameErrorStringResult()
		{
			string value = "#NAME?";
			CompileResult compileResult = new CompileResultFactory().Create(value);
			Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Name), compileResult.Result);
			Assert.AreEqual(DataType.ExcelError, compileResult.DataType);
		}

		[TestMethod]
		public void CreateIntResult()
		{
			int value = 1437;
			CompileResult compileResult = new CompileResultFactory().Create(value);
			Assert.AreEqual(value, compileResult.Result);
			Assert.AreEqual(DataType.Integer, compileResult.DataType);
		}

		[TestMethod]
		public void CreateDecimalResultFromDouble()
		{
			double value = 1437.756;
			CompileResult compileResult = new CompileResultFactory().Create(value);
			Assert.AreEqual(value, compileResult.Result);
			Assert.AreEqual(DataType.Decimal, compileResult.DataType);
		}

		[TestMethod]
		public void CreateDecimalResult()
		{
			decimal value = 1437.756m;
			CompileResult compileResult = new CompileResultFactory().Create(value);
			Assert.AreEqual(value, compileResult.Result);
			Assert.AreEqual(DataType.Decimal, compileResult.DataType);
		}

		[TestMethod]
		public void CreateResultByteConvertsToInt()
		{
			byte value = 7;
			CompileResult compileResult = new CompileResultFactory().Create(value);
			Assert.AreEqual(value, compileResult.Result);
			Assert.AreEqual(DataType.Integer, compileResult.DataType);
		}
	}
}
