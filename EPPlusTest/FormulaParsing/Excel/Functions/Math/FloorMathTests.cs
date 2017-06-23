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
	public class FloorMathTests : MathFunctionsTestBase
	{
		[TestMethod]
		public void FloorMathWithNoInputsReturnsPoundValue()
		{
			var function = new FloorMath();
			var result = function.Execute(FunctionsHelper.CreateArgs(), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}
		
		[TestMethod]
		public void FloorMathWithOneInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "FLOOR.MATH(15.67, , )";
				worksheet.Calculate();
				Assert.AreEqual(15d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void FloorMathWithTwoInputsReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "FLOOR.MATH(14.56, 3, )";
				worksheet.Calculate();
				Assert.AreEqual(12d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void FloorMathWithAllThreeInputsReturnsCorrectValue()
		{
			var function = new FloorMath();
			var result = function.Execute(FunctionsHelper.CreateArgs(14.56, 4, 6), this.ParsingContext);
			Assert.AreEqual(12d, result.Result);
		}

		[TestMethod]
		public void FloorMathWithNegativeNumberPositiveSigReturnsCorrectValue()
		{
			var function = new FloorMath();
			var result = function.Execute(FunctionsHelper.CreateArgs(-15.36, 2), this.ParsingContext);
			Assert.AreEqual(-16d, result.Result);
		}

		[TestMethod]
		public void FloorMathWithNegativeNumberNegativeSigReturnsCorrectValue()
		{
			var function = new FloorMath();
			var result = function.Execute(FunctionsHelper.CreateArgs(-15.6, -7), this.ParsingContext);
			Assert.AreEqual(-21d, result.Result);
		}

		[TestMethod]
		public void FloorMathWithPositiveNumberNegativeSigReturnsCorrectValue()
		{
			var function = new FloorMath();
			var result = function.Execute(FunctionsHelper.CreateArgs(15.7, -2), this.ParsingContext);
			Assert.AreEqual(14d, result.Result);
		}

		[TestMethod]
		public void FloorMathWithModeValueReturnsCorrectValue()
		{
			var function = new FloorMath();
			var result = function.Execute(FunctionsHelper.CreateArgs(-5.5, 2, -1), this.ParsingContext);
			Assert.AreEqual(-4d, result.Result);
		}

		[TestMethod]
		public void FloorMathWithNegativeNumberAndNoModeValueReturnsCorrectValue()
		{
			var function = new FloorMath();
			var result = function.Execute(FunctionsHelper.CreateArgs(-5.5, 2), this.ParsingContext);
			Assert.AreEqual(-6d, result.Result);
		}

		[TestMethod]
		public void FloorMathWithDecimalSignificanceReturnsCorrectValue()
		{
			var function = new FloorMath();
			var result = function.Execute(FunctionsHelper.CreateArgs(45.67, 2.34), this.ParsingContext);
			Assert.AreEqual(44.46d, (double)result.Result, 0.000001);
		}

		[TestMethod]
		public void FloorMathWithZeroModeReturnsCorrectValue()
		{
			var function = new FloorMath();
			var result = function.Execute(FunctionsHelper.CreateArgs(-5.5, 2, 0), this.ParsingContext);
			Assert.AreEqual(-6d, result.Result);
		}

		[TestMethod]
		public void FloorMathWithZeroSignificanceReturnsCorrectValue()
		{
			var function = new FloorMath();
			var result = function.Execute(FunctionsHelper.CreateArgs(56.7, 0), this.ParsingContext);
			Assert.AreEqual(0d, result.Result);
		}

		[TestMethod]
		public void FloorMathWithGeneralStringNumberReturnsPoundValue()
		{
			var function = new FloorMath();
			var result = function.Execute(FunctionsHelper.CreateArgs("String", 3, 5), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void FloorMathWithDateFunctionNumberReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "FLOOR.MATH(DATE(2017, 6, 14), 4)";
				worksheet.Calculate();
				Assert.AreEqual(42900d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void FloorMathWithDateAsStringNumberReturnsCorrectValue()
		{
			var function = new FloorMath();
			var result = function.Execute(FunctionsHelper.CreateArgs("5/5/2017", 4), this.ParsingContext);
			Assert.AreEqual(42860d, result.Result);
		}

		[TestMethod]
		public void FloorMathWithBooleanNumberInputReturnsZero()
		{
			var function = new FloorMath();
			var booleanTrue = function.Execute(FunctionsHelper.CreateArgs(true, 5), this.ParsingContext);
			var booleanFalse = function.Execute(FunctionsHelper.CreateArgs(false, 3), this.ParsingContext);
			Assert.AreEqual(0d, booleanTrue.Result);
			Assert.AreEqual(0d, booleanFalse.Result);
		}

		[TestMethod]
		public void FloorMathWithNumericStringNumberReturnsCorrectValue()
		{
			var function = new FloorMath();
			var result = function.Execute(FunctionsHelper.CreateArgs("5.67", 3), this.ParsingContext);
			Assert.AreEqual(3d, result.Result);
		}

		[TestMethod]
		public void FloorMathWithGeneralStringSignificanceReturnsCorrectValue()
		{
			var function = new FloorMath();
			var result = function.Execute(FunctionsHelper.CreateArgs(34.5, "String"), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void FloorMathWithDateFunctionSignificanceReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "FLOOR.MATH(34.5, DATE(2017, 6, 15))";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void FloorMathWithDateAsStringSignificanceReturnsCorrectValue()
		{
			var function = new FloorMath();
			var result = function.Execute(FunctionsHelper.CreateArgs(45548564.6, "5/5/2017"), this.ParsingContext);
			Assert.AreEqual(45517320d, result.Result);
		}

		[TestMethod]
		public void FloorMathWithBooleanSignificanceReturnsCorrectValue()
		{
			var function = new FloorMath();
			var booleanTrue = function.Execute(FunctionsHelper.CreateArgs(45.6, true), this.ParsingContext);
			var booleanFalse = function.Execute(FunctionsHelper.CreateArgs(45.6, false), this.ParsingContext);
			Assert.AreEqual(45d, booleanTrue.Result);
			Assert.AreEqual(0d, booleanFalse.Result);
		}

		[TestMethod]
		public void FloorMathWithNumericStringSignificanceReturnsCorrectValue()
		{
			var function = new FloorMath();
			var result = function.Execute(FunctionsHelper.CreateArgs(4.67, "2"), this.ParsingContext);
			Assert.AreEqual(4d, result.Result);
		}



		// Below are tests that deal with mode input.

		[TestMethod]
		public void FloorMathWithGeneralStringModeInputReturnsCorrectValue()
		{
			var function = new FloorMath();
			var result = function.Execute(FunctionsHelper.CreateArgs(34.5, 3, "string"), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void FloorMathWithDateFunctionModeInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "FLOOR.MATH(-45.5, 3, DATE(2017, 6, 15))";
				worksheet.Calculate();
				Assert.AreEqual(-45d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void FloorMathWithDateAsStringModeInputReturnsCorrectValue()
		{
			var function = new FloorMath();
			var result = function.Execute(FunctionsHelper.CreateArgs(-45.6, 3, "5/5/2017"), this.ParsingContext);
			Assert.AreEqual(-45d, result.Result);
		}

		[TestMethod]
		public void FloorMathWithBooleanModeInputReturnsCorrectValue()
		{
			var function = new FloorMath();
			var booleanTrue = function.Execute(FunctionsHelper.CreateArgs(-45.7, 3, true), this.ParsingContext);
			var booleanFalse = function.Execute(FunctionsHelper.CreateArgs(-45.7, 3, false), this.ParsingContext);
			Assert.AreEqual(-45d, booleanTrue.Result);
			Assert.AreEqual(-48d, booleanFalse.Result);
		}

		[TestMethod]
		public void FloorMathWithNumericStringModeInputReturnsCorrectValue()
		{
			var function = new FloorMath();
			var result = function.Execute(FunctionsHelper.CreateArgs(-45.7, 4, "2"), this.ParsingContext);
			Assert.AreEqual(-44d, result.Result);
		}

		[TestMethod]
		public void FloorMathWithFirstAndThirdInputsReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "FLOOR.MATH(45.6, ,7)";
				worksheet.Calculate();
				Assert.AreEqual(45d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void FloorMathWithOnlySecondAndThirdInputsReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "FLOOR.MATH(, 2, 3)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void FloorMathWithOnlyThirdInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "FLOOR.MATH(, , 2)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B1"].Value);
			}
		}
		
		[TestMethod]
		public void FloorMathWithOnlySecondInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "FLOOR.MATH(, 2, )";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B1"].Value);
			}
		}
	}
}
