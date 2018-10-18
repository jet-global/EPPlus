using System;
using System.Globalization;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace EPPlusTest.FormulaParsing.ExpressionGraph
{
	[TestClass]
	public class CompileResultTests
	{
		#region TestMethods
		[TestMethod]
		public void NumericStringCompileResult()
		{
			var expected = 124.24;
			string numericString = expected.ToString("n");
			CompileResult result = new CompileResult(numericString, DataType.String);
			Assert.IsFalse(result.IsNumeric);
			Assert.IsTrue(result.IsNumericOrDateString);
			Assert.AreEqual(expected, result.ResultNumeric);
		}

		[TestMethod]
		public void DateStringCompileResult()
		{
			var expected = new DateTime(2013, 1, 15);
			string dateString = expected.ToString("d");
			CompileResult result = new CompileResult(dateString, DataType.String);
			Assert.IsFalse(result.IsNumeric);
			Assert.IsTrue(result.IsNumericOrDateString);
			Assert.AreEqual(expected.ToOADate(), result.ResultNumeric);
		}

		[TestMethod]
		public void DateCompileResultAsNumeric()
		{
			var date = DateTime.Now;
			var compileResult = new CompileResult(date, DataType.Date);
			Assert.IsTrue(compileResult.IsNumeric);
			Assert.AreEqual(date.ToOADate(), compileResult.ResultNumeric);
		}

		[TestMethod]
		public void IsNumericOrDateStringTest()
		{
			var compileResult = new CompileResult(null, DataType.String);
			Assert.IsFalse(compileResult.IsNumericOrDateString);
			Assert.AreEqual(0, compileResult.ResultNumeric);
			compileResult = new CompileResult(12d, DataType.Decimal);
			Assert.IsFalse(compileResult.IsNumericOrDateString);

			// Number
			compileResult = new CompileResult("1,234", DataType.String);
			Assert.IsTrue(compileResult.IsNumericOrDateString);
			Assert.AreEqual(1234, compileResult.ResultNumeric);

			// Date
			compileResult = new CompileResult("10/21/2018", DataType.String);
			Assert.IsTrue(compileResult.IsNumericOrDateString);
			Assert.AreEqual(43394, compileResult.ResultNumeric);

			var currentCulture = Thread.CurrentThread.CurrentCulture;
			try
			{
				// Ambiguous Date
				// In German, the "." is both the date seperator and the number group seperator.
				Thread.CurrentThread.CurrentCulture = new CultureInfo("de-DE");
				// 1.1 should compile to the OADate of january first.
				compileResult = new CompileResult("1.1", DataType.String);
				Assert.IsTrue(compileResult.IsNumericOrDateString);
				Assert.AreEqual(new DateTime(DateTime.Today.Year, 1, 1).ToOADate(), compileResult.ResultNumeric);

				// Ambiguous Number
				// In English, 1.1 parses to both a date and a number but should be a number.
				Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
				compileResult = new CompileResult("1.1", DataType.String);
				Assert.IsTrue(compileResult.IsNumericOrDateString);
				Assert.AreEqual(1.1d, compileResult.ResultNumeric);
			}
			finally
			{
				Thread.CurrentThread.CurrentCulture = currentCulture;
			}
		}

		[TestMethod]
		public void IsNumericOrDateStringNumberFormatsTest()
		{
			double expected = 1442.0;
			var compileResult = new CompileResult("1.442E+003", DataType.String);
			Assert.IsTrue(compileResult.IsNumericOrDateString);
			Assert.AreEqual(expected, compileResult.ResultNumeric);
			compileResult = new CompileResult("1442", DataType.String);
			Assert.IsTrue(compileResult.IsNumericOrDateString);
			Assert.AreEqual(expected, compileResult.ResultNumeric);
			compileResult = new CompileResult("1442.00", DataType.String);
			Assert.IsTrue(compileResult.IsNumericOrDateString);
			Assert.AreEqual(expected, compileResult.ResultNumeric);
			compileResult = new CompileResult("1,442.0", DataType.String);
			Assert.IsTrue(compileResult.IsNumericOrDateString);
			Assert.AreEqual(expected, compileResult.ResultNumeric);
			expected = -0.00526;
			compileResult = new CompileResult("-5.26E-003", DataType.String);
			Assert.IsTrue(compileResult.IsNumericOrDateString);
			Assert.AreEqual(expected, compileResult.ResultNumeric);
			compileResult = new CompileResult("-0", DataType.String);
			Assert.IsTrue(compileResult.IsNumericOrDateString);
			Assert.AreEqual(0d, compileResult.ResultNumeric);
			compileResult = new CompileResult("-0.005", DataType.String);
			Assert.IsTrue(compileResult.IsNumericOrDateString);
			Assert.AreEqual(-0.005d, compileResult.ResultNumeric);
			compileResult = new CompileResult("-0.005260", DataType.String);
			Assert.IsTrue(compileResult.IsNumericOrDateString);
			Assert.AreEqual(expected, compileResult.ResultNumeric);

			compileResult = new CompileResult("2,345,656,567", DataType.String);
			Assert.IsTrue(compileResult.IsNumericOrDateString);
			Assert.AreEqual(2345656567d, compileResult.ResultNumeric);
			compileResult = new CompileResult("123", DataType.String);
			Assert.IsTrue(compileResult.IsNumericOrDateString);
			Assert.AreEqual(123d, compileResult.ResultNumeric);
			compileResult = new CompileResult("12345", DataType.String);
			Assert.IsTrue(compileResult.IsNumericOrDateString);
			Assert.AreEqual(12345d, compileResult.ResultNumeric);
			compileResult = new CompileResult("-12345", DataType.String);
			Assert.IsTrue(compileResult.IsNumericOrDateString);
			Assert.AreEqual(-12345d, compileResult.ResultNumeric);
			compileResult = new CompileResult("-12,345", DataType.String);
			Assert.IsTrue(compileResult.IsNumericOrDateString);
			Assert.AreEqual(-12345d, compileResult.ResultNumeric);
			compileResult = new CompileResult("-122,345", DataType.String);
			Assert.IsTrue(compileResult.IsNumericOrDateString);
			Assert.AreEqual(-122345d, compileResult.ResultNumeric);

			var currentCulture = Thread.CurrentThread.CurrentCulture;
			try
			{
				Thread.CurrentThread.CurrentCulture = new TestCulture("en-US", new NumberFormatInfo { NumberGroupSizes = new int[4] { 2, 3, 4, 0 }, NumberGroupSeparator = "*" });
				compileResult = new CompileResult("1", DataType.String);
				Assert.IsTrue(compileResult.IsNumericOrDateString);
				Assert.AreEqual(1d, compileResult.ResultNumeric);
				compileResult = new CompileResult("12", DataType.String);
				Assert.IsTrue(compileResult.IsNumericOrDateString);
				Assert.AreEqual(12d, compileResult.ResultNumeric);
				compileResult = new CompileResult("123", DataType.String);
				Assert.IsTrue(compileResult.IsNumericOrDateString);
				Assert.AreEqual(123d, compileResult.ResultNumeric);
				compileResult = new CompileResult("1*23", DataType.String);
				Assert.IsTrue(compileResult.IsNumericOrDateString);
				Assert.AreEqual(123d, compileResult.ResultNumeric);
				compileResult = new CompileResult("123*23", DataType.String);
				Assert.IsTrue(compileResult.IsNumericOrDateString);
				Assert.AreEqual(12323d, compileResult.ResultNumeric);
				compileResult = new CompileResult("6*123*23", DataType.String);
				Assert.IsTrue(compileResult.IsNumericOrDateString);
				Assert.AreEqual(612323d, compileResult.ResultNumeric);
				compileResult = new CompileResult("9876*123*23", DataType.String);
				Assert.IsTrue(compileResult.IsNumericOrDateString);
				Assert.AreEqual(987612323d, compileResult.ResultNumeric);
				compileResult = new CompileResult("9*9876*123*23", DataType.String);
				Assert.IsTrue(compileResult.IsNumericOrDateString);
				Assert.AreEqual(9987612323d, compileResult.ResultNumeric);
				compileResult = new CompileResult("1231239*9876*123*23", DataType.String);
				Assert.IsTrue(compileResult.IsNumericOrDateString);
				Assert.AreEqual(1231239987612323d, compileResult.ResultNumeric);

				Thread.CurrentThread.CurrentCulture = new TestCulture("en-US", new NumberFormatInfo { NumberGroupSizes = new int[2] { 2, 3, }, NumberGroupSeparator = "*", NumberDecimalSeparator = "#", NegativeSign = "_" });
				compileResult = new CompileResult("_123*23#23", DataType.String);
				Assert.IsTrue(compileResult.IsNumericOrDateString);
				Assert.AreEqual(-12323.23d, compileResult.ResultNumeric);
				compileResult = new CompileResult("6*123*23#", DataType.String);
				Assert.IsTrue(compileResult.IsNumericOrDateString);
				Assert.AreEqual(612323d, compileResult.ResultNumeric);
				compileResult = new CompileResult("126*123*23#8989", DataType.String);
				Assert.IsTrue(compileResult.IsNumericOrDateString);
				Assert.AreEqual(12612323.8989d, compileResult.ResultNumeric);

				Thread.CurrentThread.CurrentCulture = new TestCulture("en-US", new NumberFormatInfo { NumberGroupSizes = new int[0] });
				compileResult = new CompileResult("2345", DataType.String);
				Assert.IsTrue(compileResult.IsNumericOrDateString);
				Assert.AreEqual(2345d, compileResult.ResultNumeric);

				// In German, the "." is both the date seperator and the number group seperator.
				Thread.CurrentThread.CurrentCulture = new CultureInfo("de-DE");
				// So 1.1 is january first and should compile to the OADate.
				compileResult = new CompileResult("1.1", DataType.String);
				Assert.IsTrue(compileResult.IsNumericOrDateString);
				Assert.AreEqual(new DateTime(DateTime.Today.Year, 1, 1).ToOADate(), compileResult.ResultNumeric);
			}
			finally
			{
				Thread.CurrentThread.CurrentCulture = currentCulture;
			}
		}
		#endregion

		#region Nested Classes
		private class TestCulture : CultureInfo
		{
			public TestCulture(string name, NumberFormatInfo numberFormat) : base(name)
			{
				this.NumberFormat = numberFormat;
			}
		}
		#endregion
	}
}
