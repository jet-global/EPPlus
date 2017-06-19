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
	public class ModTests : MathFunctionsTestBase
	{
		#region Mod Function (Execute) Tests

		[TestMethod]
		public void ModWithNoInputsReturnsPoundValue()
		{
			var function = new Mod();
			var result = function.Execute(FunctionsHelper.CreateArgs(), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void ModWithPositiveInputsReturnsCorrectValue()
		{
			var function = new Mod();
			var result = function.Execute(FunctionsHelper.CreateArgs(15, 9), this.ParsingContext);
			Assert.AreEqual(6d, result.Result);
		}

		[TestMethod]
		public void ModWithNegativeNumberArgumentReturnsCorrectValue()
		{
			var function = new Mod();
			var result = function.Execute(FunctionsHelper.CreateArgs(-5, 7), this.ParsingContext);
			Assert.AreEqual(2d, result.Result);
		}

		[TestMethod]
		public void ModWithNegativeDivisorResturnsCorrectValue()
		{
			var function = new Mod();
			var result = function.Execute(FunctionsHelper.CreateArgs(5, -7), this.ParsingContext);
			Assert.AreEqual(-2d, result.Result);
		}

		[TestMethod]
		public void ModWithLargeNumbersReturnsCorrectValue()
		{
			var function = new Mod();
			var result = function.Execute(FunctionsHelper.CreateArgs(4, -10000), this.ParsingContext);
			Assert.AreEqual(-9996d, result.Result);
		}

		[TestMethod]
		public void ModWithNegativeIntegersReturnsCorrectValue()
		{
			var function = new Mod();
			var result = function.Execute(FunctionsHelper.CreateArgs(-78, -9), this.ParsingContext);
			Assert.AreEqual(-6d, result.Result);
		}

		[TestMethod]
		public void ModWithPositiveDoublesReturnsCorrectValue()
		{
			var function = new Mod();
			var result = function.Execute(FunctionsHelper.CreateArgs(15.9, 1.3), this.ParsingContext);
			Assert.AreEqual(0.3d, (double)result.Result, 0.000001);
		}

		[TestMethod]
		public void ModWithPositiveFractionsReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "MOD((2/3), (6/5))";
				worksheet.Calculate();
				Assert.AreEqual(0.66666667, (double)worksheet.Cells["B1"].Value, 0.000001);
			}
		}

		[TestMethod]
		public void ModWithZeroDivisorReturnsPoundDivZero()
		{
			var function = new Mod();
			var result = function.Execute(FunctionsHelper.CreateArgs(7, 0), this.ParsingContext);
			Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void ModWithGeneralStringInputReturnsPoundValue()
		{
			var function = new Mod();
			var firstInputString = function.Execute(FunctionsHelper.CreateArgs("string", 10), this.ParsingContext);
			var secondInputString = function.Execute(FunctionsHelper.CreateArgs(5, "string"), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)firstInputString.Result).Type);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)secondInputString.Result).Type);
		}

		[TestMethod]
		public void ModWithNumericStringInputReturnsCorrectValue()
		{
			var function = new Mod();
			var firstInputNumericString = function.Execute(FunctionsHelper.CreateArgs("6", 4), this.ParsingContext);
			var secondInputNumericString = function.Execute(FunctionsHelper.CreateArgs(7, "6"), this.ParsingContext);
			Assert.AreEqual(2d, firstInputNumericString.Result);
			Assert.AreEqual(1d, secondInputNumericString.Result);
		}

		[TestMethod]
		public void ModWithDateObjectInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "MOD(DATE(2017, 6, 15), 6)";
				worksheet.Cells["B2"].Formula = "MOD(9, DATE(2017, 6, 8))";
				worksheet.Calculate();
				Assert.AreEqual(1d, worksheet.Cells["B1"].Value);
				Assert.AreEqual(9d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void ModWithDateAsStringInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "MOD(\"5/2/2017\", 7)";
				worksheet.Cells["B2"].Formula = "MOD(7, \"5/8/2017\")";
				worksheet.Calculate();
				Assert.AreEqual(3d, worksheet.Cells["B1"].Value);
				Assert.AreEqual(7d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void ModWithBooleanAsTrueReturnsCorrectValue()
		{
			var function = new Mod();
			var numberBoolean = function.Execute(FunctionsHelper.CreateArgs(true, 5), this.ParsingContext);
			var divisorBoolean = function.Execute(FunctionsHelper.CreateArgs(10, true), this.ParsingContext);
			Assert.AreEqual(1d, numberBoolean.Result);
			Assert.AreEqual(0d, divisorBoolean.Result);
		}

		[TestMethod]
		public void ModWithBooleanAsFalseReturnsCorrectValueOrPoundDivZero()
		{
			var function = new Mod();
			var numberBoolean = function.Execute(FunctionsHelper.CreateArgs(false, 10), this.ParsingContext);
			var divisorBoolean = function.Execute(FunctionsHelper.CreateArgs(10, false), this.ParsingContext);
			Assert.AreEqual(0d, numberBoolean.Result);
			Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)divisorBoolean.Result).Type);
		}

		[TestMethod]
		public void ModWithReferenceToCellsReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 5;
				worksheet.Cells["B2"].Formula = "MOD(B1, 10)";
				worksheet.Calculate();
				Assert.AreEqual(5d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void ModShouldReturnCorrectResult()
		{
			var func = new Mod();
			var args = FunctionsHelper.CreateArgs(5, 2);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(1d, result.Result);
		}

		[TestMethod]
		public void ModWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Mod();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}
		#endregion
	}
}
