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
using System;
using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math
{
	[TestClass]
	public class QuotientTests : MathFunctionsTestBase
	{
		#region Quotient Function (Execute) Tests
		[TestMethod]
		public void QuotientWithTwoPositiveIntegersReturnsCorrectValue()
		{
			var function = new Quotient();
			var result = function.Execute(FunctionsHelper.CreateArgs(10, 5), this.ParsingContext);
			Assert.AreEqual(2, result.Result);
		}

		[TestMethod]
		public void QuotientWithDoubleArgumentsReturnsCorrectValue()
		{
			var function = new Quotient();
			var result = function.Execute(FunctionsHelper.CreateArgs(6.3, 3.3), this.ParsingContext);
			Assert.AreEqual(1, result.Result);
		}

		[TestMethod]
		public void QuotientWithPositiveAndNegativeIntegerReturnsCorrectValue()
		{
			var function = new Quotient();
			var result = function.Execute(FunctionsHelper.CreateArgs(-10, 3), this.ParsingContext);
			Assert.AreEqual(-3, result.Result);
		}

		[TestMethod]
		public void QuotientWithDivisionByZeroReturnsDivZeroError()
		{
			var function = new Quotient();
			var result = function.Execute(FunctionsHelper.CreateArgs(1, 0), this.ParsingContext);
			Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void QuotientWithZeroAsNumeratorReturnsCorrectValue()
		{
			var function = new Quotient();
			var result = function.Execute(FunctionsHelper.CreateArgs(0, 1), this.ParsingContext);
			Assert.AreEqual(0, result.Result);
		}

		[TestMethod]
		public void QuotientWithNullFirstArgumentReturnsPoundNA()
		{
			var function = new Quotient();
			var result = function.Execute(FunctionsHelper.CreateArgs(null, 2), this.ParsingContext);
			Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void QuotientWithNullSecondArgumentReturnsPoundNA()
		{
			var function = new Quotient();
			var result = function.Execute(FunctionsHelper.CreateArgs(2, null), this.ParsingContext);
			Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void QuotientWithNoArgumentsReturnsPoundValue()
		{
			var function = new Quotient();
			var result = function.Execute(FunctionsHelper.CreateArgs(), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void QuotientWithDateFunctionArgumentsReturnsCorrectValue()
		{
			var function = new Quotient();
			var firstArgAsDate = new DateTime(2017, 5, 1);
			var secondArgAsDate = new DateTime(2017, 6, 1);
			var result = function.Execute(FunctionsHelper.CreateArgs(firstArgAsDate, secondArgAsDate), this.ParsingContext);
			Assert.AreEqual(0, result.Result);
		}

		[TestMethod]
		public void QuotientWithDatesAsStringArgumentsReturnsCorrectValue()
		{
			var function = new Quotient();
			var result = function.Execute(FunctionsHelper.CreateArgs("5/1/2017", "6/1/2017"), this.ParsingContext);
			Assert.AreEqual(0, result.Result);
		}

		[TestMethod]
		public void QuotientWithFractionsReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["B1"].Formula = "QUOTIENT((2/3),(1/3))";
				ws.Calculate();
				Assert.AreEqual(2, ws.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void QuotientWithStringArgumentsReturnsPoundValue()
		{
			var function = new Quotient();
			var result = function.Execute(FunctionsHelper.CreateArgs("string", "string"), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void QuotientWithNumericStringsReturnsCorrectValue()
		{
			var function = new Quotient();
			var result = function.Execute(FunctionsHelper.CreateArgs("10", "5"), this.ParsingContext);
			Assert.AreEqual(2, result.Result);
		}

		[TestMethod]
		public void QuotientShouldReturnCorrectResult()
		{
			var func = new Quotient();
			var args = FunctionsHelper.CreateArgs(5, 2);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(2, result.Result);
		}

		[TestMethod]
		public void QuotientShouldPoundDivZeroWhenDenomIs0()
		{
			var func = new Quotient();
			var args = FunctionsHelper.CreateArgs(1, 0);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(OfficeOpenXml.FormulaParsing.ExpressionGraph.DataType.ExcelError, result.DataType);
			Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)(result.Result)).Type);
		}

		[TestMethod]
		public void QuotientWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Quotient();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void QuotientFunctionWithErrorValuesAsInputReturnsTheInputErrorValue()
		{
			var func = new Quotient();
			var argNA = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.NA),0);
			var argNAME = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Name),0);
			var argVALUE = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Value),0);
			var argNUM = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Num),0);
			var argDIV0 = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Div0),0);
			var argREF = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Ref),0);
			var resultNA = func.Execute(argNA, this.ParsingContext);
			var resultNAME = func.Execute(argNAME, this.ParsingContext);
			var resultVALUE = func.Execute(argVALUE, this.ParsingContext);
			var resultNUM = func.Execute(argNUM, this.ParsingContext);
			var resultDIV0 = func.Execute(argDIV0, this.ParsingContext);
			var resultREF = func.Execute(argREF, this.ParsingContext);
			Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)resultNA.Result).Type);
			Assert.AreEqual(eErrorType.Name, ((ExcelErrorValue)resultNAME.Result).Type);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)resultVALUE.Result).Type);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)resultNUM.Result).Type);
			Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)resultDIV0.Result).Type);
			Assert.AreEqual(eErrorType.Ref, ((ExcelErrorValue)resultREF.Result).Type);
		}
		#endregion
	}
}
