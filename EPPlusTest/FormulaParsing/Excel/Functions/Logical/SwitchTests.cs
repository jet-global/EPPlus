using System.Collections.Generic;
using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Logical
{
	[TestClass]
	public class SwitchTests
	{
		#region Class Variables
		private ParsingContext _parsingContext = ParsingContext.Create();
		#endregion

		#region TestMethods
		[TestMethod]
		public void SwitchFunctionZeroArguments()
		{
			var func = new Switch();
			var result = func.Execute(new List<FunctionArgument>(), _parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SwitchFunctionFewArguments()
		{
			var function = new Switch();
			var arguments = FunctionsHelper.CreateArgs(1, 1);
			var result = function.Execute(arguments, _parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SwitchFunctionTooManyArguments()
		{
			List<int> list = new List<int>(260);
			for (int i = 0; i < 260; i++)
			{
				list.Add(i);
			}
			var function = new Switch();
			var arguments = FunctionsHelper.CreateArgs(list);
			var result = function.Execute(arguments, _parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SwitchFunctionWithMatchStringArguments()
		{
			var function = new Switch();
			var arguments = FunctionsHelper.CreateArgs("one", "one", "Return one", "Default");
			var result = function.Execute(arguments, _parsingContext);
			Assert.AreEqual("Return one", result.Result);
		}

		[TestMethod]
		public void SwitchFunctionWithNoMatchWithDefault()
		{
			var function = new Switch();
			var arguments = FunctionsHelper.CreateArgs(1, 2, "Two", "Default");
			var result = function.Execute(arguments, _parsingContext);
			Assert.AreEqual("Default", result.Result);
		}

		[TestMethod]
		public void SwitchFunctionWithNoMatchNoDefault()
		{
			var function = new Switch();
			var arguments = FunctionsHelper.CreateArgs(1, 2, "Two");
			var result = function.Execute(arguments, _parsingContext);
			Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SwitchFunctionWithMatchNoDefault()
		{
			var function = new Switch();
			var arguments = FunctionsHelper.CreateArgs(1, 1, "One");
			var result = function.Execute(arguments, _parsingContext);
			Assert.AreEqual("One", result.Result);
		}

		[TestMethod]
		public void SwitchFunctionWithMatchDefault()
		{
			var function = new Switch();
			var arguments = FunctionsHelper.CreateArgs(1, 1, "One", "Default");
			var result = function.Execute(arguments, _parsingContext);
			Assert.AreEqual("One", result.Result);
		}

		[TestMethod]
		public void SwitchFunctionWithThirdValueMatch()
		{
			var function = new Switch();
			var arguments = FunctionsHelper.CreateArgs(1, 2, "two", 3, "third", 1, "One", "Default");
			var result = function.Execute(arguments, _parsingContext);
			Assert.AreEqual("One", result.Result);
		}

		[TestMethod]
		public void SwitchFunctionWithDoubleArguments()
		{
			var function = new Switch();
			var arguments = FunctionsHelper.CreateArgs(5.2, 2.1, "two", 3.1, "third", 5.2, "Five and two tenths", "Default");
			var result = function.Execute(arguments, _parsingContext);
			Assert.AreEqual("Five and two tenths", result.Result);
		}
		#endregion

		#region Integration Tests
		[TestMethod]
		public void SwitchFunctionIntegrationTrueTest()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet1");
				sheet.Cells[2, 2].Formula = @"SWITCH(1, 1, ""One"", ""No Match"")";
				sheet.Calculate();
				Assert.AreEqual("One", sheet.Cells[2, 2].Value);
			}
		}

		[TestMethod]
		public void SwitchFunctionMatchWithReferenceIntegrationTest()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet1");
				sheet.Cells["A1"].Formula = @"SWITCH(A3, 1, ""Sunday"", 2, ""Monday"", 3, ""Tuesday"")";
				sheet.Cells["A3"].Formula = "3";
				sheet.Calculate();
				Assert.AreEqual("Tuesday", sheet.Cells["A1"].Value);
			}
		}

		[TestMethod]
		public void SwitchFunctionNoMatchWithReferenceIntegrationTest()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet1");
				sheet.Cells["A1"].Formula = @"SWITCH(A3, 1, ""Sunday"", 2, ""Monday"", 3, ""Tuesday"", ""Default"")";
				sheet.Cells["A3"].Formula = "10";
				sheet.Calculate();
				Assert.AreEqual("Default", sheet.Cells["A1"].Value);
			}
		}

		[TestMethod]
		public void SwitchFunctionNestedMatchWithReferenceIntegrationTest()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet1");
				sheet.Cells["A1"].Formula = @"SWITCH(IF(A3>2, 1, 2), 1, ""Yes"", 2, ""No"")";
				sheet.Cells["A3"].Formula = "10";
				sheet.Calculate();
				Assert.AreEqual("Yes", sheet.Cells["A1"].Value);
			}
		}

		[TestMethod]
		public void SwitchFunctionNoMatchReturnDefaultWithDifferentTypeArgumentsIntegrationTest()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet1");
				sheet.Cells[2, 2].Formula = @"SWITCH(1, ""1"", ""Sunday"", 2, ""Monday"", 3, ""Tuesday"", ""Default"")";
				sheet.Calculate();
				Assert.AreEqual("Default", sheet.Cells[2, 2].Value);
			}
		}

		[TestMethod]
		public void SwitchFunctionNoMatchReturnDefaultWithDateArgumentsIntegrationTest()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet1");
				sheet.Cells[2, 2].Formula = @"SWITCH(1/1/2015, ""1/1/2015"", ""1st January"", ""Default"")";
				sheet.Calculate();
				Assert.AreEqual("Default", sheet.Cells[2, 2].Value);
			}
		}
		#endregion
	}
}