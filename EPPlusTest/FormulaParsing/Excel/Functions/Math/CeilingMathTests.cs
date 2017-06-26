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
	public class CeilingMathTests : MathFunctionsTestBase
	{
		[TestMethod]
		public void CeilingMathWithNoInputsReturnsCorrectValue()
		{
			var function = new CeilingMath();
			var result = function.Execute(FunctionsHelper.CreateArgs(), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void CeilingMathWithFirstInputOnlyReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "CEILING.MATH(10.5, , )";
				worksheet.Calculate();
				Assert.AreEqual(11d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void CeilingMathWithFirstTwoInputsOnlyReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "CEILING.MATH(10.5, 2, )";
				worksheet.Calculate();
				Assert.AreEqual(12d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void CeilingMathWithAllThreeInputsReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "CEILING.MATH(10.5, 2, 1)";
				worksheet.Calculate();
				Assert.AreEqual(12d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void CeilingMathWithSecondAndThirdInputsReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "CEILING.MATH(, 2, 2)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void CeilingMathWithFirstAndThirdInputsReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "CEILING.MATH(10.5, , 3)";
				worksheet.Calculate();
				Assert.AreEqual(11d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void CeilingMathWithOnlyThirdInputReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "CEILING.MATH(, , 3)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void CeilingMathWithOnlySecondInputReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "CEILING.MATH(, 3, )";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void CeilingMathWithAllPositiveInputsReturnsCorrectValue()
		{
			var function = new CeilingMath();
			var result = function.Execute(FunctionsHelper.CreateArgs(10.5, 2, 1), this.ParsingContext);
			Assert.AreEqual(12d, result.Result);
		}

		[TestMethod]
		public void CeilingMathWithFirstTwoPositiveLastNegativeInputReturnsCorrectValue()
		{
			var function = new CeilingMath();
			var result = function.Execute(FunctionsHelper.CreateArgs(10.5, 2, -2), this.ParsingContext);
			Assert.AreEqual(12d, result.Result);
		}

		[TestMethod]
		public void CeilingMathWithMiddleNegativeInputReturnsCorrectValue()
		{
			var function = new CeilingMath();
			var result = function.Execute(FunctionsHelper.CreateArgs(10.5, -2, 2), this.ParsingContext);
			Assert.AreEqual(12d, result.Result);
		}

		[TestMethod]
		public void CeilingMathWithLastTwoNegativeInputsReturnsCorrectValue()
		{
			var function = new CeilingMath();
			var result = function.Execute(FunctionsHelper.CreateArgs(10.5, -2, -2), this.ParsingContext);
			Assert.AreEqual(12d, result.Result);
		}

		[TestMethod]
		public void CeilingMathWithAllNegativeInputsReturnsCorrectValue()
		{
			var function = new CeilingMath();
			var result = function.Execute(FunctionsHelper.CreateArgs(-10.5, -3, -1), this.ParsingContext);
			Assert.AreEqual(-12d, result.Result);
		}

		[TestMethod]
		public void CeilingMathWithFirstTwoInputsNegativeReturnsCorrectValue()
		{
			var function = new CeilingMath();
			var result = function.Execute(FunctionsHelper.CreateArgs(-10.5, -2, 2), this.ParsingContext);
			Assert.AreEqual(-12d, result.Result);
		}

		[TestMethod]
		public void CeilingMathWithMiddleInputPositiveReturnsCorrectValue()
		{
			var function = new CeilingMath();
			var result = function.Execute(FunctionsHelper.CreateArgs(-10.5, 2, -1), this.ParsingContext);
			Assert.AreEqual(-12d, result.Result);
		}

		[TestMethod]
		public void CeilingMathWithLastTwoInputsPositiveReturnsCorrectValue()
		{
			var function = new CeilingMath();
			var result = function.Execute(FunctionsHelper.CreateArgs(-10.5, 2, 2), this.ParsingContext);
			Assert.AreEqual(-12d, result.Result);
		}

		[TestMethod]
		public void CeilingMathWithIntegerNumberInputReturnsCorrectValue()
		{
			var function = new CeilingMath();
			var result = function.Execute(FunctionsHelper.CreateArgs(10, 3), this.ParsingContext);
			Assert.AreEqual(12d, result.Result);
		}

		[TestMethod]
		public void CeilingMathWithDoubleNumberInputReturnsCorrectValue()
		{
			var function = new CeilingMath();
			var result = function.Execute(FunctionsHelper.CreateArgs(10.5, 5), this.ParsingContext);
			Assert.AreEqual(15d, result.Result);
		}

		[TestMethod]
		public void CeilingMathWithFractionNumberInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "CEILING.MATH((2/3), 3)";
				worksheet.Calculate();
				Assert.AreEqual(3d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void CeilingMathWithStringNumberInputReturnsPoundValue()
		{
			var function = new CeilingMath();
			var result = function.Execute(FunctionsHelper.CreateArgs("string", 2), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void CeilingMathWithDateFunctionNumberInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "CEILING.MATH(DATE(2017, 5, 8), 4)";
				worksheet.Calculate();
				Assert.AreEqual(42864d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void CeilingMathWithDateAsStringNumberInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "CEILING.MATH(\"5/5/2017\", 3)";
				worksheet.Calculate();
				Assert.AreEqual(42861d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void CeilingMathWithBooleanNumberInputReturnsCorrectValue()
		{
			var function = new CeilingMath();
			var booleanTrue = function.Execute(FunctionsHelper.CreateArgs(true, 10), this.ParsingContext);
			var booleanFalse = function.Execute(FunctionsHelper.CreateArgs(false, 3), this.ParsingContext);
			Assert.AreEqual(10d, booleanTrue.Result);
			Assert.AreEqual(0d, booleanFalse.Result);
		}

		[TestMethod]
		public void CeilingMathWithIntegerSignificanceReturnsCorrectValue()
		{
			var function = new CeilingMath();
			var result = function.Execute(FunctionsHelper.CreateArgs(10.3, 5), this.ParsingContext);
			Assert.AreEqual(15d, result.Result);
		}

		[TestMethod]
		public void CeilingMathWithDoubleSignificanceReturnsCorrectValue()
		{
			var function = new CeilingMath();
			var result = function.Execute(FunctionsHelper.CreateArgs(10.4, 5.6), this.ParsingContext);
			Assert.AreEqual(11.2d, result.Result);
		}

		[TestMethod]
		public void CeilingMathWithFractionSignificanceReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "CEILING.MATH(10.6, (2/3))";
				worksheet.Calculate();
				Assert.AreEqual(10.66667d, (double)worksheet.Cells["B1"].Value, 0.000001);
			}
		}

		[TestMethod]
		public void CeilingMathWithStringSignificanceReturnsPoundValue()
		{
			var function = new CeilingMath();
			var result = function.Execute(FunctionsHelper.CreateArgs(10.5, "String"), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void CeilingMathWithDateFunctionSignificanceReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "CEILING.MATH(10.5, DATE(2017, 6, 5))";
				worksheet.Calculate();
				Assert.AreEqual(42891d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void CeilingMathWithDateAsStringSignificanceReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "CEILING.MATH(10.5, \"5/5/2017\")";
				worksheet.Calculate();
				Assert.AreEqual(42860d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void CeilingMathWithBooleanSignificanceReturnsCorrectValue()
		{
			var function = new CeilingMath();
			var booleanTrue = function.Execute(FunctionsHelper.CreateArgs(10.5, true), this.ParsingContext);
			var booleanFalse = function.Execute(FunctionsHelper.CreateArgs(10.4, false), this.ParsingContext);
			Assert.AreEqual(11d, booleanTrue.Result);
			Assert.AreEqual(0d, booleanFalse.Result);
		}

		[TestMethod]
		public void CeilingMathWithZeroSignificanceReturnsZero()
		{
			var function = new CeilingMath();
			var result = function.Execute(FunctionsHelper.CreateArgs(10.4, 0), this.ParsingContext);
			Assert.AreEqual(0d, result.Result);
		}

		[TestMethod]
		public void CeilingMathWithIntegerModeReturnsCorrectValue()
		{
			var function = new CeilingMath();
			var result = function.Execute(FunctionsHelper.CreateArgs(-10.3, 2, 2), this.ParsingContext);
			Assert.AreEqual(-12d, result.Result);
		}

		[TestMethod]
		public void CeilingMathWithDoubleModeReturnsCorrectValue()
		{
			var function = new CeilingMath();
			var result = function.Execute(FunctionsHelper.CreateArgs(-10.5, 2, 23.45), this.ParsingContext);
			Assert.AreEqual(-12d, result.Result);
		}

		[TestMethod]
		public void CeilingMathWithFractionModeReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "CEILING.MATH(-10.4, 2, (2/3))";
				worksheet.Calculate();
				Assert.AreEqual(-12d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void CeilingMathWithStringModeInputReturnsPoundValue()
		{
			var function = new CeilingMath();
			var result = function.Execute(FunctionsHelper.CreateArgs(-10.4, 2, "string"), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void CeilingMathWithDateFunctionModeInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "CEILING.MATH(-10.4, 2, DATE(2017, 6, 5))";
				worksheet.Calculate();
				Assert.AreEqual(-12d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void CeilingMathWithDateAsStringModeInputReturnsCorrectValue()
		{
			var function = new CeilingMath();
			var result = function.Execute(FunctionsHelper.CreateArgs(-10.4, 2, "5/5/2017"), this.ParsingContext);
			Assert.AreEqual(-12d, result.Result);
		}

		[TestMethod]
		public void CeilingMathWithBooleanModeInputReturnsCorrectValue()
		{
			var function = new CeilingMath();
			var booleanTrue = function.Execute(FunctionsHelper.CreateArgs(-10.4, 2, true), this.ParsingContext);
			var booleanFalse = function.Execute(FunctionsHelper.CreateArgs(-10.4, 2, false), this.ParsingContext);
			Assert.AreEqual(-12d, booleanTrue.Result);
			Assert.AreEqual(-10d, booleanFalse.Result);
		}

		[TestMethod]
		public void CeilingMathWithNegativeModeInputReturnsCorrectValue()
		{
			var function = new CeilingMath();
			var result = function.Execute(FunctionsHelper.CreateArgs(-10.4, 2, -10), this.ParsingContext);
			Assert.AreEqual(-12d, result.Result);
		}

		[TestMethod]
		public void CeilingMathWithZeroModeReturnsCorrectValue()
		{
			var function = new CeilingMath();
			var result = function.Execute(FunctionsHelper.CreateArgs(-10.3, 2, 0), this.ParsingContext);
			Assert.AreEqual(-10d, result.Result);
		}
	}
}
