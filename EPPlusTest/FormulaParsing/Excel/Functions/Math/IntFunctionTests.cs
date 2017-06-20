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
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Numeric;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math
{
	[TestClass]
	public class IntFunctionTests : MathFunctionsTestBase
	{
		#region IntFunction Function (Execute) Tests
		[TestMethod]
		public void IntFunctionShouldConvertTextToInteger()
		{
			var func = new IntFunction();
			var args = FunctionsHelper.CreateArgs("2");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(2, result.Result);
		}

		[TestMethod]
		public void IntFunctionWithInvalidArgumentReturnsPoundValue()
		{
			var func = new IntFunction();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void IntFunctionShouldConvertDecimalToInteger()
		{
			var func = new IntFunction();
			var args = FunctionsHelper.CreateArgs(2.88m);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(2, result.Result);
		}

		[TestMethod]
		public void IntFunctionShouldConvertNegativeDecimalToInteger()
		{
			var func = new IntFunction();
			var args = FunctionsHelper.CreateArgs(-2.88m);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(-3, result.Result);
		}

		[TestMethod]
		public void IntFunctionShouldConvertStringToInteger()
		{
			var func = new IntFunction();
			var args = FunctionsHelper.CreateArgs("-2.88");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(-3, result.Result);
		}

		[TestMethod]
		public void IntFunctionWithStringInputReturnsPoundValue()
		{
			var function = new IntFunction();
			var result = function.Execute(FunctionsHelper.CreateArgs("string"), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void IntFunctionWithDateFromDateFunctionReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "INT(DATE(2017, 6, 12))";
				worksheet.Calculate();
				Assert.AreEqual(42898, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void IntFunctionWithDateAsStringReturnsCorrectValue()
		{
			var function = new IntFunction();
			var result = function.Execute(FunctionsHelper.CreateArgs("5/5/2017"), this.ParsingContext);
			Assert.AreEqual(42860, result.Result);
		}

		[TestMethod]
		public void IntFunctionWithCellReferenceReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 15.9;
				worksheet.Cells["B2"].Formula = "INT(B1)";
				worksheet.Calculate();
				Assert.AreEqual(15, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void IntFunctionWithEmptyCellReferenceReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "INT(A2:A4)";
				worksheet.Calculate();
				Assert.AreEqual(0, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void IntFunctionWithTrueBooleanReturnsCorrectValue()
		{
			var function = new IntFunction();
			var result = function.Execute(FunctionsHelper.CreateArgs(true), this.ParsingContext);
			Assert.AreEqual(1, result.Result);
		}

		[TestMethod]
		public void IntFunctionWithFalseBooleanReturnsCorrectValue()
		{
			var function = new IntFunction();
			var result = function.Execute(FunctionsHelper.CreateArgs(false), this.ParsingContext);
			Assert.AreEqual(0, result.Result);
		}

		[TestMethod]
		public void IntFunctionWithPositiveFractionReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "INT((2/3))";
				worksheet.Calculate();
				Assert.AreEqual(0, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void IntFunctionWithNegativeFractionReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "INT((-2/3))";
				worksheet.Calculate();
				Assert.AreEqual(-1, worksheet.Cells["B1"].Value);
			}
		}
		#endregion
	}
}
