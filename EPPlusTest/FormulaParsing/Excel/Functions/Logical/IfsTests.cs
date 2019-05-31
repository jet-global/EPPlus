using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Logical
{
	[TestClass]
	public class IfsTests
	{
		#region Class Variables
		private ParsingContext _parsingContext = ParsingContext.Create();
		#endregion

		#region TestMethods
		[TestMethod]
		public void IfsFunctionSingleConditionTrue()
		{
			string expected = "ASDAS";
			var func = new Ifs();
			var args = FunctionsHelper.CreateArgs(true, expected);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(expected, (string)result.Result);
		}

		[TestMethod]
		public void IfsFunctionSingleConditionTrueString()
		{
			string expected = "ASDAS";
			var func = new Ifs();
			var args = FunctionsHelper.CreateArgs("true", expected);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(expected, (string)result.Result);
		}

		[TestMethod]
		public void IfsFunctionSingleConditionFalse()
		{
			string expected = "ASDAS";
			var func = new Ifs();
			var args = FunctionsHelper.CreateArgs(false, expected);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void IfsFunctionZeroArguments()
		{
			var func = new Ifs();
			var result = func.Execute(new List<FunctionArgument>(), _parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void IfsFunctionOddArgumentCount()
		{
			var func = new Ifs();
			var args = FunctionsHelper.CreateArgs(1, 2, 3, 4, 5);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}
		#endregion

		#region Integration Tests
		[TestMethod]
		public void IfsFunctionIntegrationTrueTest()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet1");
				sheet.Cells[2, 2].Formula = @"IFS(2=1, ""badbadbad"", 2=2, ""good"")";
				sheet.Calculate();
				Assert.AreEqual("good", sheet.Cells[2, 2].Value);
			}
		}
		#endregion
	}
}
