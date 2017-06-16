/*******************************************************************************
* You may amend and distribute as you like, but don't remove this header!
*
* EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
* See http://www.codeplex.com/EPPlus for details.
*
* Copyright (C) 2011-2017 Jan Källman, Matt Delaney, and others as noted in the source history.
*
* This library is free software; you can redistribute it and/or
* modify it under the terms of the GNU Lesser General Public
* License as published by the Free Software Foundation; either
* version 2.1 of the License, or (at your option) any later version.

* This library is distributed in the hope that it will be useful,
* but WITHOUT ANY WARRANTY; without even the implied warranty of
* MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
* See the GNU Lesser General Public License for more details.
*
* The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
* If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
*
* All code and executables are provided "as is" with no warranty either express or implied. 
* The author accepts no liability for any damage or loss of business that this product may cause.
*
* For code change notes, see the source control history.
*******************************************************************************/
using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Logical
{
	[TestClass]
	public class IfErrorTests : LogicalFunctionsTestBase
	{
		#region IfError Tests
		[TestMethod]
		public void IfErrorWithTooFewArgumentsReturnsPoundValue()
		{
			var function = new IfError();
			var arguments = FunctionsHelper.CreateArgs();
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void IfErrorWithNoErrorReturnsCorrectResult()
		{
			var function = new IfError();
			var arguments = FunctionsHelper.CreateArgs(1, "word");
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(1, result.Result);
		}

		[TestMethod]
		public void IfErrorWithErrorAndIntegerReturnsCorrectResult()
		{
			var function = new IfError();
			var arguments = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.NA), 1);
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(1, result.Result);
		}

		[TestMethod]
		public void IfErrorWithErrorAndDoubleReturnsCorrectResult()
		{
			var function = new IfError();
			var arguments = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Value), 1.5);
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(1.5, result.Result);
		}

		[TestMethod]
		public void IfErrorWithErrorAndStringReturnsCorrectResult()
		{
			var function = new IfError();
			var arguments = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Num), "word");
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual("word", result.Result);
		}

		[TestMethod]
		public void IfErrorWithErrorAndEmptyStringReturnsCorrectResult()
		{
			var function = new IfError();
			var arguments = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Div0), string.Empty);
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(string.Empty, result.Result);
		}

		[TestMethod]
		public void IfErrorWithErrorAndNullArgumentReturnsCorrectResult()
		{
			var function = new IfError();
			var arguments = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Name), null);
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(0, result.Result);
		}

		[TestMethod]
		public void IfErrorWithIntegerAndErrorReturnsCorrectResult()
		{
			var function = new IfError();
			var arguments = FunctionsHelper.CreateArgs(1, ExcelErrorValue.Create(eErrorType.NA));
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(1, result.Result);
		}

		[TestMethod]
		public void IfErrorWithDoubleAndErrorReturnsCorrectResult()
		{
			var function = new IfError();
			var arguments = FunctionsHelper.CreateArgs(1.5, ExcelErrorValue.Create(eErrorType.Value));
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(1.5, result.Result);
		}

		[TestMethod]
		public void IfErrorWithStringAndErrorReturnsCorrectResult()
		{
			var function = new IfError();
			var arguments = FunctionsHelper.CreateArgs("word", ExcelErrorValue.Create(eErrorType.Num));
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual("word", result.Result);
		}

		[TestMethod]
		public void IfErrorWithEmptyStringAndErrorReturnsCorrectResult()
		{
			var function = new IfError();
			var arguments = FunctionsHelper.CreateArgs(string.Empty, ExcelErrorValue.Create(eErrorType.Div0));
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(string.Empty, result.Result);
		}

		[TestMethod]
		public void IfErrorWithNullArgumentAndErrorReturnsCorrectResult()
		{
			var function = new IfError();
			var arguments = FunctionsHelper.CreateArgs(null, ExcelErrorValue.Create(eErrorType.Name));
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(0, result.Result);
		}

		[TestMethod]
		public void IfErrorWithNAErrorAndNameErrorReturnsCorrectResult()
		{
			var function = new IfError();
			var arguments = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.NA), ExcelErrorValue.Create(eErrorType.Name));
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(eErrorType.Name, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void IfErrorWithNameErrorAndNAErrorReturnsCorrectResult()
		{
			var function = new IfError();
			var arguments = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Name), ExcelErrorValue.Create(eErrorType.NA));
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void IfErrorWithErrorValuesAsInputReturnsCorrectResults()
		{
			var function = new IfError();
			var argumentsNA = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.NA), 1);
			var argumentsNAME = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Name), 1);
			var argumentsVALUE = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Value), 1);
			var argumentsNUM = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Num), 1);
			var argumentsDIV0 = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Div0), 1);
			var argumentsREF = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Ref), 1);
			var resultNA = function.Execute(argumentsNA, this.ParsingContext);
			var resultNAME = function.Execute(argumentsNAME, this.ParsingContext);
			var resultVALUE = function.Execute(argumentsVALUE, this.ParsingContext);
			var resultNUM = function.Execute(argumentsNUM, this.ParsingContext);
			var resultDIV0 = function.Execute(argumentsDIV0, this.ParsingContext);
			var resultREF = function.Execute(argumentsREF, this.ParsingContext);
			Assert.AreEqual(1, resultNA.Result);
			Assert.AreEqual(1, resultNAME.Result);
			Assert.AreEqual(1, resultVALUE.Result);
			Assert.AreEqual(1, resultNUM.Result);
			Assert.AreEqual(1, resultDIV0.Result);
			Assert.AreEqual(1, resultREF.Result);
		}

		[TestMethod]
		public void IfErrorInWorksheetWithNoErrorReturnsCorrectResult()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["A1"].Formula = "IFERROR(A2, \"hello\")";
				worksheet.Cells["A2"].Value = "word";
				worksheet.Calculate();
				Assert.AreEqual("word", worksheet.Cells["A1"].Value);
			}
		}

		[TestMethod]
		public void IfErrorInWorksheetWithStringAndErrorReturnsCorrectResult()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("test");
				worksheet.Cells["A1"].Formula = "IFERROR(\"hello\", 1/0)";
				worksheet.Calculate();
				Assert.AreEqual("hello", worksheet.Cells["A1"].Value);
			}
		}

		[TestMethod]
		public void IfErrorInWorksheetWithErrorAndStringReturnsCorrectResult()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("test");
				worksheet.Cells["A1"].Formula = "IFERROR(A2, \"hello\")";
				worksheet.Cells["A2"].Formula = "1/0";
				worksheet.Calculate();
				Assert.AreEqual("hello", worksheet.Cells["A1"].Value);
			}
		}

		[TestMethod]
		public void IfErrorInWorksheetBothArgumentsAsErrorsReturnsCorrectResult()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["A1"].Formula = "IFERROR(1/0, A2)";
				worksheet.Cells["A2"].Formula= "invalidFormulaToCreateNameError";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Name, ((ExcelErrorValue)worksheet.Cells["A1"].Value).Type);
			}
		}
		#endregion
	}
}
