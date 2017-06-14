using System.IO;
using System.Linq;
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

		[TestMethod]
		public void IfShouldReturnCorrectResult()
		{
			var func = new If();
			var args = FunctionsHelper.CreateArgs(true, "A", "B");
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual("A", result.Result);
		}

		[TestMethod, Ignore]
		public void IfShouldIgnoreCase()
		{
			using (var pck = new ExcelPackage(new FileInfo(@"c:\temp\book1.xlsx")))
			{
				pck.Workbook.Calculate();
				Assert.AreEqual("Sant", pck.Workbook.Worksheets.First().Cells["C3"].Value);
			}
		}

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
		public void OrShouldReturnTrueIfOneArgumentIsTrue()
		{
			var func = new Or();
			var args = FunctionsHelper.CreateArgs(true, false, false);
			var result = func.Execute(args, _parsingContext);
			Assert.IsTrue((bool)result.Result);
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
		public void IfErrorShouldReturnSecondArgIfCriteriaEvaluatesAsAnError()
		{
			using (var package = new ExcelPackage())
			{
				var s1 = package.Workbook.Worksheets.Add("test");
				s1.Cells["A1"].Formula = "IFERROR(0/0, \"hello\")";
				s1.Calculate();
				Assert.AreEqual("hello", s1.Cells["A1"].Value);
			}
		}

		[TestMethod]
		public void IfErrorShouldReturnSecondArgIfCriteriaEvaluatesAsAnError2()
		{
			using (var package = new ExcelPackage())
			{
				var s1 = package.Workbook.Worksheets.Add("test");
				s1.Cells["A1"].Formula = "IFERROR(A2, \"hello\")";
				s1.Cells["A2"].Formula = "23/0";
				s1.Calculate();
				Assert.AreEqual("hello", s1.Cells["A1"].Value);
			}
		}

		[TestMethod]
		public void IfErrorShouldReturnResultOfFormulaIfNoError()
		{
			using (var package = new ExcelPackage())
			{
				var s1 = package.Workbook.Worksheets.Add("test");
				s1.Cells["A1"].Formula = "IFERROR(A2, \"hello\")";
				s1.Cells["A2"].Value = "hi there";
				s1.Calculate();
				Assert.AreEqual("hi there", s1.Cells["A1"].Value);
			}
		}

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
		public void AndWithInvalidArgumentReturnsPoundValue()
		{
			var func = new And();
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
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
		public void IfErrorWithInvalidArgumentReturnsPoundValue()
		{
			var func = new IfError();
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void IfErrorFunctionWithErrorValuesAsInputReturnsCorrectResults()
		{
			Assert.Fail("This test will fail until the IfError function is fixed. As of 6/14/2017 the IfError function behaves completely different from the Excel function.");
			var func = new IfError();
			var parsingContext = ParsingContext.Create();
			var argNA = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.NA),1);
			var argNAME = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Name),1);
			var argVALUE = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Value),1);
			var argNUM = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Num),1);
			var argDIV0 = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Div0),1);
			var argREF = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Ref),1);
			var resultNA = func.Execute(argNA, parsingContext);
			var resultNAME = func.Execute(argNAME, parsingContext);
			var resultVALUE = func.Execute(argVALUE, parsingContext);
			var resultNUM = func.Execute(argNUM, parsingContext);
			var resultDIV0 = func.Execute(argDIV0, parsingContext);
			var resultREF = func.Execute(argREF, parsingContext);
			Assert.AreEqual(1, resultNA.Result);
			Assert.AreEqual(1, resultNAME.Result);
			Assert.AreEqual(1, resultVALUE.Result);
			Assert.AreEqual(1, resultNUM.Result);
			Assert.AreEqual(1, resultDIV0.Result);
			Assert.AreEqual(1, resultREF.Result);
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

		[TestMethod]
		public void NotWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Not();
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
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

	}
}
