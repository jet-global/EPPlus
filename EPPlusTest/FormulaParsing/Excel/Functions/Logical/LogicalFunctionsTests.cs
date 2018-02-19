using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;

namespace EPPlusTest.Excel.Functions
{
	[TestClass]
	public class LogicalFunctionsTests
	{
		private ParsingContext _parsingContext = ParsingContext.Create();

		#region Combination Logical Test Methods
		[TestMethod]
		public void IfNestedInIfErrorReturnsPoundValueToTheIfErrorInsteadOfThrowingException()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("New Sheet");
				sheet.Cells[2, 2].Formula = "5-\"asdf\"";
				sheet.Cells[2, 3].Formula = "IF(B2, \"Defs True\", \"Totes False\")";
				sheet.Cells[2, 4].Formula = "IFERROR(C2, \"It's an Error\")";
				sheet.Cells["E14"].Value = 0;
				sheet.Cells["F14"].Value = "text";
				sheet.Cells[2, 5].Formula = "IFERROR(IF(E14-F14<0,E14-F14,0),\"Error Occurred\")";
				sheet.Calculate();
				Assert.AreEqual("#VALUE!", sheet.Cells[2, 2].Value.ToString());
				Assert.AreEqual("#VALUE!", sheet.Cells[2, 3].Value.ToString());
				Assert.AreEqual("It's an Error", sheet.Cells[2, 4].Value);
				Assert.AreEqual("Error Occurred", sheet.Cells[2, 5].Value);
			}
		}
		#endregion

		#region NOT Test Methods
		[TestMethod]
		public void NotShouldReturnFalseIfArgumentIsTrue()
		{
			var func = new Not();
			var args = FunctionsHelper.CreateArgs(true);
			var result = func.Execute(args, _parsingContext);
			Assert.IsFalse((bool)result.Result);
		}

		[TestMethod]
		public void NotShouldReturnTrueIfArgumentIs0()
		{
			var func = new Not();
			var args = FunctionsHelper.CreateArgs(0);
			var result = func.Execute(args, _parsingContext);
			Assert.IsTrue((bool)result.Result);
		}

		[TestMethod]
		public void NotShouldReturnFalseIfArgumentIs1()
		{
			var func = new Not();
			var args = FunctionsHelper.CreateArgs(1);
			var result = func.Execute(args, _parsingContext);
			Assert.IsFalse((bool)result.Result);
		}

		[TestMethod]
		public void NotShouldHandleExcelReference()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("sheet1");
				sheet.Cells["A1"].Value = false;
				sheet.Cells["A2"].Formula = "NOT(A1)";
				sheet.Calculate();
				Assert.IsTrue((bool)sheet.Cells["A2"].Value);
			}
		}

		[TestMethod]
		public void NotShouldHandleExcelReferenceToStringFalse()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("sheet1");
				sheet.Cells["A1"].Value = "false";
				sheet.Cells["A2"].Formula = "NOT(A1)";
				sheet.Calculate();
				Assert.IsTrue((bool)sheet.Cells["A2"].Value);
			}
		}

		[TestMethod]
		public void NotShouldHandleExcelReferenceToStringTrue()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("sheet1");
				sheet.Cells["A1"].Value = "TRUE";
				sheet.Cells["A2"].Formula = "NOT(A1)";
				sheet.Calculate();
				Assert.IsFalse((bool)sheet.Cells["A2"].Value);
			}
		}

		[TestMethod]
		public void NotWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Not();
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}
		#endregion

		#region OR Test Methods
		[TestMethod]
		public void OrShouldReturnTrueIfOneArgumentIsTrue()
		{
			var func = new Or();
			var args = FunctionsHelper.CreateArgs(true, false, false);
			var result = func.Execute(args, _parsingContext);
			Assert.IsTrue((bool)result.Result);
		}

		[TestMethod]
		public void OrShouldReturnTrueIfRangeArgumentIsTrue()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet1");
				sheet.Cells[2, 2].Value = 0;
				sheet.Cells[2, 3].Value = -1231230;
				sheet.Cells[2, 4].Value = false;
				sheet.Cells[3, 3].Formula = "OR(B2:D2)";
				sheet.Cells[3, 3].Calculate();
				Assert.IsTrue((bool)sheet.Cells[3, 3].Value);
			}
		}

		[TestMethod]
		public void OrShouldReturnFalseIfAllRangeArgumentValuesAreFalse()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet1");
				sheet.Cells[2, 2].Value = "strings are false.";
				sheet.Cells[2, 3].Value = string.Empty;
				sheet.Cells[2, 4].Value = false;
				sheet.Cells[3, 3].Formula = "OR(B2:D2)";
				sheet.Cells[3, 3].Calculate();
				Assert.IsFalse((bool)sheet.Cells[3, 3].Value);
			}
		}

		[TestMethod]
		public void OrShouldReturnTrueIfOneArgumentIsTrueString()
		{
			var func = new Or();
			var args = FunctionsHelper.CreateArgs("true", "FALSE", false);
			var result = func.Execute(args, _parsingContext);
			Assert.IsTrue((bool)result.Result);
		}

		[TestMethod]
		public void OrWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Or();
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}
		#endregion

		#region AND Test Methods
		[TestMethod]
		public void AndShouldHandleStringLiteralTrue()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("sheet1");
				sheet.Cells["A1"].Value = "tRuE";
				sheet.Cells["A2"].Formula = "AND(\"TRUE\", A1)";
				sheet.Calculate();
				Assert.IsTrue((bool)sheet.Cells["A2"].Value);
			}
		}

		[TestMethod]
		public void AndShouldReturnTrueIfAllArgumentsAreTrue()
		{
			var func = new And();
			var args = FunctionsHelper.CreateArgs(true, true, true);
			var result = func.Execute(args, _parsingContext);
			Assert.IsTrue((bool)result.Result);
		}

		[TestMethod]
		public void AndShouldReturnTrueIfAllArgumentsAreTrueOr1()
		{
			var func = new And();
			var args = FunctionsHelper.CreateArgs(true, true, 1, true, 1);
			var result = func.Execute(args, _parsingContext);
			Assert.IsTrue((bool)result.Result);
		}

		[TestMethod]
		public void AndShouldReturnFalseIfOneArgumentIsFalse()
		{
			var func = new And();
			var args = FunctionsHelper.CreateArgs(true, false, true);
			var result = func.Execute(args, _parsingContext);
			Assert.IsFalse((bool)result.Result);
		}

		[TestMethod]
		public void AndShouldReturnFalseIfOneArgumentIs0()
		{
			var func = new And();
			var args = FunctionsHelper.CreateArgs(true, 0, true);
			var result = func.Execute(args, _parsingContext);
			Assert.IsFalse((bool)result.Result);
		}

		[TestMethod]
		public void AndShouldReturnTrueIfRangeArgumentValuesAreTrue()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet1");
				sheet.Cells[2, 2].Value = 9999;
				sheet.Cells[2, 3].Value = -1231230;
				sheet.Cells[2, 4].Value = true;
				sheet.Cells[3, 3].Formula = "AND(B2:D2)";
				sheet.Cells[3, 3].Calculate();
				Assert.IsTrue((bool)sheet.Cells[3, 3].Value);
			}
		}

		[TestMethod]
		public void AndShouldReturnFalseIfAnyRangeArgumentValuesAreFalse()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet1");
				sheet.Cells[2, 2].Value = 0;
				sheet.Cells[2, 3].Value = 234234;
				sheet.Cells[2, 4].Value = true;
				sheet.Cells[3, 3].Formula = "AND(B2:D2)";
				sheet.Cells[3, 3].Calculate();
				Assert.IsFalse((bool)sheet.Cells[3, 3].Value);
			}
		}

		[TestMethod]
		public void AndWithInvalidArgumentReturnsPoundValue()
		{
			var func = new And();
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}
		#endregion

		#region IF Test Methods
		[TestMethod]
		public void IfShouldReturnCorrectResult()
		{
			var func = new If();
			var args = FunctionsHelper.CreateArgs(true, "A", "B");
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual("A", result.Result);
		}

		[TestMethod]
		public void IfWithInvalidArgumentReturnsPoundValue()
		{
			var func = new If();
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void IfWithErrorArgumentPropagatesError()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet1");
				sheet.Cells[2, 2].Value = "#NAME?";
				sheet.Cells[2, 3].Formula = "IF(B2=0, \"hide\",\"show\")";
				sheet.Cells[2, 4].Formula = "IF(SUM(B2)=0, \"hide\",\"show\")";
				sheet.Calculate();
				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Name), sheet.Cells[2, 3].Value as ExcelErrorValue);
				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Name), sheet.Cells[2, 4].Value as ExcelErrorValue);
			}
		}
		#endregion

		#region IFNA TestMethods
		[TestMethod]
		public void IfNaShouldReturnSecondArgIfCriteriaEvaluatesAsAnError2()
		{
			using (var package = new ExcelPackage())
			{
				var s1 = package.Workbook.Worksheets.Add("test");
				s1.Cells["A1"].Formula = "IFERROR(A2, \"hello\")";
				s1.Cells["A2"].Value = ExcelErrorValue.Create(eErrorType.NA);
				s1.Calculate();
				Assert.AreEqual("hello", s1.Cells["A1"].Value);
			}
		}

		[TestMethod]
		public void IfNaShouldReturnResultOfFormulaIfNoError()
		{
			using (var package = new ExcelPackage())
			{
				var s1 = package.Workbook.Worksheets.Add("test");
				s1.Cells["A1"].Formula = "IFNA(A2, \"hello\")";
				s1.Cells["A2"].Value = "hi there";
				s1.Calculate();
				Assert.AreEqual("hi there", s1.Cells["A1"].Value);
			}
		}

		[TestMethod]
		public void IfNaWithInvalidArgumentReturnsPoundValue()
		{
			var func = new IfNa();
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}
		#endregion
	}
}
