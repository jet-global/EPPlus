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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml;
namespace EPPlusTest.FormulaParsing.Excel.Functions.Math
{
	[TestClass]
	public class RoundDownTests : MathFunctionsTestBase
	{
		#region RoundDown Function (Execute) Tests
		[TestMethod]
		public void RoundDownWithNoInputsReturnsPoundValue()
		{
			var function = new Rounddown();
			var result = function.Execute(FunctionsHelper.CreateArgs(), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void RoundDownWithNoSecondInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "ROUNDDOWN(15, )";
				worksheet.Calculate();
				Assert.AreEqual(15d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void RoundDownWithNoFirstInputReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "ROUNDDOWN(,2)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void RoundDownWithPositiveSecondInputReturnsCorrectValue()
		{
			var function = new Rounddown();
			var result = function.Execute(FunctionsHelper.CreateArgs(12.356, 1), this.ParsingContext);
			Assert.AreEqual(12.3d, result.Result);
		}

		[TestMethod]
		public void RoundDownWithNegativeSecondInputReturnsCorrectValue()
		{
			var function = new Rounddown();
			var result = function.Execute(FunctionsHelper.CreateArgs(9333.23, -3), this.ParsingContext);
			Assert.AreEqual(9000d, result.Result);
		}

		[TestMethod]
		public void RoundDownWithZeroSecondInputReturnsCorrectValue()
		{
			var function = new Rounddown();
			var result = function.Execute(FunctionsHelper.CreateArgs(58.999999, 0), this.ParsingContext);
			Assert.AreEqual(58d, result.Result);
		}

		[TestMethod]
		public void RoundDownWithNumericStringSecondArgumentReturnsCorrectValue()
		{
			var function = new Rounddown();
			var result = function.Execute(FunctionsHelper.CreateArgs(15.351, "2"), this.ParsingContext);
			Assert.AreEqual(15.35, result.Result);
		}

		[TestMethod]
		public void RoundDownWithGeneralStringSecondArgumentReturnsPoundValue()
		{
			var function = new Rounddown();
			var result = function.Execute(FunctionsHelper.CreateArgs("string", 2), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void RoundDownWithSecondArgumentAsDateFunctinReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "ROUNDDOWN(23.564, DATE(2017, 6, 5))";
				worksheet.Calculate();
				Assert.AreEqual(23.564d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void RoundDownWithSecondArgumentAsDateAsStringReturnsCorrectValue()
		{
			var function = new Rounddown();
			var result = function.Execute(FunctionsHelper.CreateArgs(52.3, "5/5/2017"), this.ParsingContext);
			Assert.AreEqual(52.3d, result.Result);
		}

		[TestMethod]
		public void RoundDownWithFirstArgumentAsCellReferenceReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 16.235;
				worksheet.Cells["B2"].Formula = "ROUNDDOWN(B1, 2)";
				worksheet.Calculate();
				Assert.AreEqual(16.23d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void RoundDownWithSecondArgumentAsCellReferenceReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 2;
				worksheet.Cells["B2"].Formula = "ROUNDDOWN(16.235, B1)";
				worksheet.Calculate();
				Assert.AreEqual(16.23d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void RoundDownWithEmptyCellReferenceReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "ROUNDDOWN(A2, A3)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void RoundDownWithErrorValueAsInputReturnsRespectiveError()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "SQRT(-1)";
				worksheet.Cells["B2"].Formula = "ROUNDDOWN(12.3, B1)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)worksheet.Cells["B2"].Value).Type);
			}
		}

		[TestMethod]
		public void RoundDownWithPositiveIntegerAndPositiveSecondInputReturnsCorrectValue()
		{
			var function = new Rounddown();
			var result = function.Execute(FunctionsHelper.CreateArgs(56, 2), this.ParsingContext);
			Assert.AreEqual(56d, result.Result);
		}

		[TestMethod]
		public void RoundDownWithDoubleAndPositiveSecondInputReturnsCorrectValue()
		{
			var function = new Rounddown();
			var result = function.Execute(FunctionsHelper.CreateArgs(26.325655, 4), this.ParsingContext);
			Assert.AreEqual(26.3256d, result.Result);
		}

		[TestMethod]
		public void RoundDownWithFractionInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "ROUNDDOWN((2/3), 4)";
				worksheet.Calculate();
				Assert.AreEqual(0.6666d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void RoundDownWithNumericStringFirstInputReturnsCorrectValue()
		{
			var function = new Rounddown();
			var result = function.Execute(FunctionsHelper.CreateArgs("45546.546444", 3), this.ParsingContext);
			Assert.AreEqual(45546.546d, result.Result);
		}

		[TestMethod]
		public void RoundDownWithGeneralStringFirstInputReturnsCorrectValue()
		{
			var function = new Rounddown();
			var result = function.Execute(FunctionsHelper.CreateArgs("string", 3), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void RoundDownWithFirstInputAsDateFunctionInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B3"].Formula = "DATE(2017, 6, 5)";
				worksheet.Cells["B1"].Formula = "ROUNDDOWN(DATE(2017, 6, 5), 1)";
				worksheet.Calculate();
				Assert.AreEqual(worksheet.Cells["B3"].Value, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void RoundDownWithFirstInputAsDateAsStringReturnsCorrectValue()
		{
			var function = new Rounddown();
			var result = function.Execute(FunctionsHelper.CreateArgs("5/5/2017", 2), this.ParsingContext);
			Assert.AreEqual(42860d, result.Result);
		}

		[TestMethod]
		public void RoundDownWithNegativeDoublePositiveSecondArgReturnsCorrectValue()
		{
			var function = new Rounddown();
			var result = function.Execute(FunctionsHelper.CreateArgs(-5.326571, 3), this.ParsingContext);
			Assert.AreEqual(-5.326d, result.Result);
		}

		[TestMethod]
		public void RoundDownWithNegativeDoubleNegativeSecondArgReturnsCorrectValue()
		{
			var function = new Rounddown();
			var result = function.Execute(FunctionsHelper.CreateArgs(-3451, -2), this.ParsingContext);
			Assert.AreEqual(-3400d, result.Result);
		}

		[TestMethod]
		public void RoundDownWithNegativeIntegerPositiveSecondArgReturnsCorrectValue()
		{
			var function = new Rounddown();
			var result = function.Execute(FunctionsHelper.CreateArgs(-985, 2), this.ParsingContext);
			Assert.AreEqual(-985d, result.Result);
		}

		[TestMethod]
		public void RoundDownWithNegativeIntegerNegativeSecondArgReturnsCorrectValue()
		{
			var function = new Rounddown();
			var result = function.Execute(FunctionsHelper.CreateArgs(-98, -2), this.ParsingContext);
			Assert.AreEqual(0d, result.Result);
		}

		[TestMethod]
		public void RoundDownWithBooleanSecondInputReturnsCorrectValue()
		{
			var function = new Rounddown();
			var booleanTrue = function.Execute(FunctionsHelper.CreateArgs(12.345, true), this.ParsingContext);
			var booleanFalse = function.Execute(FunctionsHelper.CreateArgs(12.345, false), this.ParsingContext);
			Assert.AreEqual(12.3d, booleanTrue.Result);
			Assert.AreEqual(12d, booleanFalse.Result);
		}

		[TestMethod]
		public void RounddownShouldReturnCorrectResultWithPositiveNumber()
		{
			var func = new Rounddown();
			var args = FunctionsHelper.CreateArgs(9.999, 2);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(9.99, result.Result);
		}

		[TestMethod]
		public void RounddownShouldHandleNegativeNumber()
		{
			var func = new Rounddown();
			var args = FunctionsHelper.CreateArgs(-9.999, 2);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(-9.99, result.Result);
		}

		[TestMethod]
		public void RounddownShouldHandleNegativeNumDigits()
		{
			var func = new Rounddown();
			var args = FunctionsHelper.CreateArgs(999.999, -2);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(900d, result.Result);
		}

		[TestMethod]
		public void RounddownShouldReturn0IfNegativeNumDigitsIsTooLarge()
		{
			var func = new Rounddown();
			var args = FunctionsHelper.CreateArgs(999.999, -4);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(0d, result.Result);
		}

		[TestMethod]
		public void RounddownShouldHandleZeroNumDigits()
		{
			var func = new Rounddown();
			var args = FunctionsHelper.CreateArgs(999.999, 0);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(999d, result.Result);
		}
		#endregion
	}
}
