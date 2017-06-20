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
using System.Linq;
using OfficeOpenXml.FormulaParsing.Excel;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math
{
	[TestClass]
	public class SignTests :MathFunctionsTestBase
	{
		#region Sign Function (Execute) Tests

		[TestMethod]
		public void SignWithNegativeIntegerReturnsCorrectValue()
		{
			var function = new Sign();
			var args = FunctionsHelper.CreateArgs(-2);
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(-1d, result.Result);
		}

		[TestMethod]
		public void SignWithPositiveIntegerReturnsCorrectValue()
		{
			var function = new Sign();
			var args = FunctionsHelper.CreateArgs(2);
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(1d, result.Result);
		}

		[TestMethod]
		public void SignWithInvalidArgumentReturnsPoundValue()
		{
			var function = new Sign();
			var args = FunctionsHelper.CreateArgs();
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SignWithPositiveDoubleReturnsCorrectValue()
		{
			var function = new Sign();
			var result = function.Execute(FunctionsHelper.CreateArgs(7.8), this.ParsingContext);
			Assert.AreEqual(1d, result.Result);
		}

		[TestMethod]
		public void SignWithNegativeDoubleReturnsCorrectValue()
		{
			var function = new Sign();
			var result = function.Execute(FunctionsHelper.CreateArgs(-9.9), this.ParsingContext);
			Assert.AreEqual(-1d, result.Result);
		}

		[TestMethod]
		public void SignWithAdditionInFunctionCallReturnsCorrectValue()
		{
			var function = new Sign();
			var result = function.Execute(FunctionsHelper.CreateArgs(4 + 7), this.ParsingContext);
			Assert.AreEqual(1d, result.Result);
		}

		[TestMethod]
		public void SignWithSubtractionInFunctionCallReturnsCorrectValue()
		{
			var function = new Sign();
			var result = function.Execute(FunctionsHelper.CreateArgs(1 - 5), this.ParsingContext);
			Assert.AreEqual(-1d, result.Result);
		}

		[TestMethod]
		public void SignWithMultiplicationInFunctionCallReturnsCorrectValue()
		{
			var function = new Sign();
			var result = function.Execute(FunctionsHelper.CreateArgs(8 * 9), this.ParsingContext);
			Assert.AreEqual(1d, result.Result);
		}

		[TestMethod]
		public void SignWithDivisionInFunctionCallReturnsCorrectValue()
		{
			var function = new Sign();
			var result = function.Execute(FunctionsHelper.CreateArgs(5 / 4), this.ParsingContext);
			Assert.AreEqual(1d, result.Result);
		}

		[TestMethod]
		public void SignWithDivisonByZeroReturnsPoundDivZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "SIGN(5/0)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B1"].Value).Type);
			}
		}

		[TestMethod]
		public void SignWithDateFunctionReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "SIGN(DATE(2017, 5, 7))";
				worksheet.Calculate();
				Assert.AreEqual(1d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void SignWithGeneralStringReturnsPoundValue()
		{
			var function = new Sign();
			var result = function.Execute(FunctionsHelper.CreateArgs("string"), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SignWithDatesAsStringsReturnsCorrectValue()
		{
			var function = new Sign();
			var result = function.Execute(FunctionsHelper.CreateArgs("5/5/2017"), this.ParsingContext);
			Assert.AreEqual(1d, result.Result);
		}

		[TestMethod]
		public void SignWithTrueBooleanReturnsCorrectValue()
		{
			var function = new Sign();
			var result = function.Execute(FunctionsHelper.CreateArgs(true), this.ParsingContext);
			Assert.AreEqual(1d, result.Result);
		}

		[TestMethod]
		public void SignWithFalseBooleanReturnsCorrectValue()
		{
			var function = new Sign();
			var result = function.Execute(FunctionsHelper.CreateArgs(false), this.ParsingContext);
			Assert.AreEqual(0d, result.Result);
		}

		[TestMethod]
		public void SignWithZeroInputReutrnsZero()
		{
			var function = new Sign();
			var result = function.Execute(FunctionsHelper.CreateArgs(0), this.ParsingContext);
			Assert.AreEqual(0d, result.Result);
		}

		[TestMethod]
		public void SignWithPositiveNumericStringReturnsCorrectValue()
		{
			var function = new Sign();
			var result = function.Execute(FunctionsHelper.CreateArgs("3"), this.ParsingContext);
			Assert.AreEqual(1d, result.Result);
		}

		[TestMethod]
		public void SignWithNegativeNumericStringReturnsCorrectValue()
		{
			var function = new Sign();
			var result = function.Execute(FunctionsHelper.CreateArgs("-5"), this.ParsingContext);
			Assert.AreEqual(-1d, result.Result);
		}

		[TestMethod]
		public void SignWithReferenceToEmptyCellReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "SIGN(A2)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B1"].Value);
			}
		}
		#endregion
	}
}
