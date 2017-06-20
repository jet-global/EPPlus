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

			}
		}
		#endregion
	}
}
