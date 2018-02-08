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
	public class AbsTests : MathFunctionsTestBase
	{
		#region Abs Function (Execute) Tests
		[TestMethod]
		public void AbsWithNoInputsReturnsPoundValue()
		{
			var function = new Abs();
			var result = function.Execute(FunctionsHelper.CreateArgs(), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void AbsShouldReturnCorrectResult()
		{
			var expectedValue = 3d;
			var func = new Abs();
			var args = FunctionsHelper.CreateArgs(-3d);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedValue, result.Result);
		}

		[TestMethod]
		public void AbsWithPositiveIntegerReturnsCorrectResult()
		{
			var function = new Abs();
			var result = function.Execute(FunctionsHelper.CreateArgs(10), this.ParsingContext);
			Assert.AreEqual(10d, result.Result);
		}

		[TestMethod]
		public void AbsWithFractionInputReturnsCorrectResult()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "ABS(-(2/3))";
				worksheet.Calculate();
				Assert.AreEqual(.666667, (double)worksheet.Cells["B1"].Value, 0.000001);
			}
		}

		[TestMethod]
		public void AbsWithDoubleInputReturnsCorrectResult()
		{
			var function = new Abs();
			var result = function.Execute(FunctionsHelper.CreateArgs(4.67), this.ParsingContext);
			Assert.AreEqual(4.67d, result.Result);
		}

		[TestMethod]
		public void AbsWithGeneralStringInputReturnsPoundValue()
		{
			var function = new Abs();
			var result = function.Execute(FunctionsHelper.CreateArgs("string"), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SumPropagatesErrorTypes()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("sheet1");
				sheet.Cells["C2"].Formula = "VLOOKUP(D18,$F$7:$G$9,2,0)";
				sheet.Cells["C3"].Formula = "-ABS(C2)";
				sheet.Cells["C3"].Calculate();
				var actual = sheet.Cells["C3"].Value as ExcelErrorValue;
				Assert.AreEqual(eErrorType.NA, actual.Type);
			}
		}

		[TestMethod]
		public void AbsWithDateFunctionInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "ABS(DATE(2017, 6, 2))";
				worksheet.Calculate();
				Assert.AreEqual(42888d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void AbsWithDatesAsStringReturnsCorrectValue()
		{
			var function = new Abs();
			var result = function.Execute(FunctionsHelper.CreateArgs("5/5/2017"), this.ParsingContext);
			Assert.AreEqual(42860d, result.Result);
		}

		[TestMethod]
		public void AbsWithTrueBooleanReturnsCorrectValue()
		{
			var function = new Abs();
			var result = function.Execute(FunctionsHelper.CreateArgs(true), this.ParsingContext);
			Assert.AreEqual(1d, result.Result);
		}

		[TestMethod]
		public void AbsWithFalseBooleanReturnsCorrectValue()
		{
			var function = new Abs();
			var result = function.Execute(FunctionsHelper.CreateArgs(false), this.ParsingContext);
			Assert.AreEqual(0d, result.Result);
		}

		[TestMethod]
		public void AbsWithNumericStringReturnsCorrectValue()
		{
			var function = new Abs();
			var result = function.Execute(FunctionsHelper.CreateArgs("-89"), this.ParsingContext);
			Assert.AreEqual(89d, result.Result);
		}

		[TestMethod]
		public void AbsWithCellReferenceReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = -78;
				worksheet.Cells["B2"].Formula = "ABS(B1)";
				worksheet.Calculate();
				Assert.AreEqual(78d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void AbsWithEmptyCellReferenceReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "ABS(A2)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void AbsWithDivisionByZeroReturnsPoundDivZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "ABS(5/0)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B1"].Value).Type);
			}
		}
		#endregion
	}
}
