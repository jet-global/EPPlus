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
	public class CeilingTests : MathFunctionsTestBase
	{
		#region Ceiling Function (Execute) Tests
		[TestMethod]
		public void CeilingWithNoInputsReturnsPoundValue()
		{
			var function = new Ceiling();
			var result = function.Execute(FunctionsHelper.CreateArgs(), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void CeilingWithNoSecondInputReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "CEILING(10, )";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void CeilingWithNoFirstInputReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "CEILING(, 1000)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void CeilingWithPositiveInputsReturnsCorrectValue()
		{
			var function = new Ceiling();
			var result = function.Execute(FunctionsHelper.CreateArgs(3.7, 2), this.ParsingContext);
			Assert.AreEqual(4d, result.Result);
		}

		[TestMethod]
		public void CeilingWithNegativeInputsReturnsCorrectValue()
		{
			var function = new Ceiling();
			var result = function.Execute(FunctionsHelper.CreateArgs(-2.5, -2), this.ParsingContext);
			Assert.AreEqual(-4d, result.Result);
		}

		[TestMethod]
		public void CeilingWithPositiveFirstInputAndNegativeSecondInputReturnsPoundNum()
		{
			var function = new Ceiling();
			var result = function.Execute(FunctionsHelper.CreateArgs(52.3, -9), this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void CeilingWithFirstInputNegativeSecondInputPositiveReturnsCorrectValue()
		{
			var function = new Ceiling();
			var result = function.Execute(FunctionsHelper.CreateArgs(-15, 4), this.ParsingContext);
			Assert.AreEqual(-12d, result.Result);
		}

		[TestMethod]
		public void CeilingWithIntegerFirstInputReturnsCorrectValue()
		{
			var function = new Ceiling();
			var result = function.Execute(FunctionsHelper.CreateArgs(45, 2), this.ParsingContext);
			Assert.AreEqual(46d, result.Result);
		}

		[TestMethod]
		public void CeilingWithDoubleFirstInputReturnsCorrectValue()
		{
			var function = new Ceiling();
			var result = function.Execute(FunctionsHelper.CreateArgs(12.63, 3), this.ParsingContext);
			Assert.AreEqual(15d, result.Result);
		}

		[TestMethod]
		public void CeilingWithGeneralStringFirstInputReturnsCorrectValue()
		{
			var function = new Ceiling();
			var result = function.Execute(FunctionsHelper.CreateArgs("string", 3), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void CeilingWithNumbericStringFirstInputReturnsCorrectValue()
		{
			var function = new Ceiling();
			var result = function.Execute(FunctionsHelper.CreateArgs("45.6", 4), this.ParsingContext);
			Assert.AreEqual(48d, result.Result);
		}

		[TestMethod]
		public void CeilingWithDateFunctionFirstInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "CEILING(DATE(2017, 6, 5), 5)";
				worksheet.Calculate();
				Assert.AreEqual(42895d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void CeilingWithDateAsStringFirstInuptReturnsCorrectValue()
		{
			var function = new Ceiling();
			var result = function.Execute(FunctionsHelper.CreateArgs("5/5/2017", 4), this.ParsingContext);
			Assert.AreEqual(42860d, result.Result);
		}

		[TestMethod]
		public void CeilingWithBooleanFirstInputsReturnsCorrectValue()
		{
			var function = new Ceiling();
			var booleanTrue = function.Execute(FunctionsHelper.CreateArgs(true, 3), this.ParsingContext);
			var booleanFalse = function.Execute(FunctionsHelper.CreateArgs(false, 2), this.ParsingContext);
			Assert.AreEqual(3d, booleanTrue.Result);
			Assert.AreEqual(0d, booleanFalse.Result);
		}

		[TestMethod]
		public void CeilingWithCellReferenceFirstInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 15.26;
				worksheet.Cells["B2"].Formula = "CEILING(B1, 3)";
				worksheet.Calculate();
				Assert.AreEqual(18d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void CeilingWithErrorValueInputReturnsRespectiveError()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "SQRT(-1)";
				worksheet.Cells["B2"].Formula = "CEILING(B1, 3)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)worksheet.Cells["B2"].Value).Type);
			}
		}

		[TestMethod]
		public void CeilingWithSecondInputAsIntegerReturnsCorrectValue()
		{
			var function = new Ceiling();
			var result = function.Execute(FunctionsHelper.CreateArgs(6, 5), this.ParsingContext);
			Assert.AreEqual(10d, result.Result);
		}

		[TestMethod]
		public void CeilingWithDoubleSecondInputReturnsCorrectValue()
		{
			var function = new Ceiling();
			var result = function.Execute(FunctionsHelper.CreateArgs(45, 6.7), this.ParsingContext);
			Assert.AreEqual(46.9d, (double)result.Result, 0.000001);
		}

		[TestMethod]
		public void CeilingWithGeneralStringSecondInputReturnsPoundValue()
		{
			var function = new Ceiling();
			var result = function.Execute(FunctionsHelper.CreateArgs(34, "string"), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void CeilingWithNumericStringSecondInputReturnsCorrectValue()
		{
			var function = new Ceiling();
			var result = function.Execute(FunctionsHelper.CreateArgs(5, "2"), this.ParsingContext);
			Assert.AreEqual(6d, result.Result);
		}

		[TestMethod]
		public void CeilingWithDateFunctionSecondInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "CEILING(124.2, DATE(2017, 6, 15))";
				worksheet.Calculate();
				Assert.AreEqual(42901d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void CeilingWithDateAsStringSecondInputReturnsCorrectValue()
		{
			var function = new Ceiling();
			var result = function.Execute(FunctionsHelper.CreateArgs(12.3, "5/5/2017"), this.ParsingContext);
			Assert.AreEqual(42860d, result.Result);
		}

		[TestMethod]
		public void CeilingWithBooleanSecondInputReturnsCorrectValue()
		{
			var function = new Ceiling();
			var booleanTrue = function.Execute(FunctionsHelper.CreateArgs(56.672, true), this.ParsingContext);
			var booleanFalse = function.Execute(FunctionsHelper.CreateArgs(45.6, false), this.ParsingContext);
			Assert.AreEqual(57d, booleanTrue.Result);
			Assert.AreEqual(0d, booleanFalse.Result);
		}

		[TestMethod]
		public void CeilingWithCellReferenceSecondInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 15;
				worksheet.Cells["B2"].Formula = "CEILING(44, B1)";
				worksheet.Calculate();
				Assert.AreEqual(45d, worksheet.Cells["B2"].Value);
			}
		}

		//Below are the tests from EPPlus.
		[TestMethod]
		public void CeilingShouldRoundUpAccordingToParamsSignificanceLowerThan0()
		{
			var expectedValue = 22.36d;
			var func = new Ceiling();
			var args = FunctionsHelper.CreateArgs(22.35d, 0.01);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedValue, result.Result);
		}

		[TestMethod]
		public void CeilingShouldRoundTowardsZeroIfSignificanceAndNumberIsMinus0point1()
		{
			var expectedValue = -22.4d;
			var func = new Ceiling();
			var args = FunctionsHelper.CreateArgs(-22.35d, -0.1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedValue, System.Math.Round((double)result.Result, 2));
		}

		[TestMethod]
		public void CeilingShouldRoundUpAccordingToParamsSignificanceIs1()
		{
			var expectedValue = 23d;
			var func = new Ceiling();
			var args = FunctionsHelper.CreateArgs(22.35d, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedValue, result.Result);
		}

		[TestMethod]
		public void CeilingShouldRoundUpAccordingToParamsSignificanceIs10()
		{
			var expectedValue = 30d;
			var func = new Ceiling();
			var args = FunctionsHelper.CreateArgs(22.35d, 10);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedValue, result.Result);
		}

		[TestMethod]
		public void CeilingShouldRoundTowardsZeroIfSignificanceAndNumberIsNegative()
		{
			var expectedValue = -30d;
			var func = new Ceiling();
			var args = FunctionsHelper.CreateArgs(-22.35d, -10);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedValue, result.Result);
		}
		#endregion
	}
}
