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
	public class LnTests : MathFunctionsTestBase
	{
		#region Ln Function (Execute) Tests
		[TestMethod]
		public void LnWithPositiveIntegerReturnsCorrectValue()
		{
			var function = new Ln();
			var result = function.Execute(FunctionsHelper.CreateArgs(10), this.ParsingContext);
			Assert.AreEqual(2.302585093d, (double)result.Result, 0.000001);
		}

		[TestMethod]
		public void LnWithNegativeIntegerReturnsPoundNum()
		{
			var function = new Ln();
			var result = function.Execute(FunctionsHelper.CreateArgs(-1000), this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}
		
		[TestMethod]
		public void LnWithDoubleInputReturnsCorrectValue()
		{
			var function = new Ln();
			var result = function.Execute(FunctionsHelper.CreateArgs(105.654), this.ParsingContext);
			Assert.AreEqual(4.660169604d, (double)result.Result, 0.000001);
		}

		[TestMethod]
		public void LnWithExcelFractionsReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["B1"].Formula = "LN((5/4))";
				ws.Calculate();
				Assert.AreEqual(0.223143551d, (double)ws.Cells["B1"].Value, 0.000001);
			}
		}

		[TestMethod]
		public void LnWithDateFunctionInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["B1"].Formula = "LN(DATE(2017,6,15))";
				ws.Calculate();
				Assert.AreEqual(10.66665041d, (double)ws.Cells["B1"].Value, 0.000001);
			}
		}

		[TestMethod]
		public void LnWithZeroInputReturnsPoundNum()
		{
			var function = new Ln();
			var result = function.Execute(FunctionsHelper.CreateArgs(0), this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void LnWithDatesAsStringsReturnsCorrectValue()
		{
			var function = new Ln();
			var result = function.Execute(FunctionsHelper.CreateArgs("6/15/2017"), this.ParsingContext);
			Assert.AreEqual(10.66665041d, (double)result.Result, 0.000001);
		}

		[TestMethod]
		public void LnWithNumericStringReturnsCorrectValue()
		{
			var function = new Ln();
			var result = function.Execute(FunctionsHelper.CreateArgs("15"), this.ParsingContext);
			Assert.AreEqual(2.708050201d, (double)result.Result, 0.000001);
		}

		[TestMethod]
		public void LnWithOneInputReturnsCorrectValue()
		{
			var function = new Ln();
			var result = function.Execute(FunctionsHelper.CreateArgs(1), this.ParsingContext);
			Assert.AreEqual(0d, result.Result);
		}

		[TestMethod]
		public void LnWithNonRealNumberReturnsPoundNum()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["B1"].Formula = "LN(SQRT(-1))";
				ws.Calculate();
				Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)ws.Cells["B1"].Value).Type);
			}
		}

		[TestMethod]
		public void LnWithGeneralStringInputReturnsPoundValue()
		{
			var function = new Ln();
			var result = function.Execute(FunctionsHelper.CreateArgs("string"), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void LnWithNoArgumentsReturnsPoundValue()
		{
			var function = new Ln();
			var result = function.Execute(FunctionsHelper.CreateArgs(), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void LnWithEInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["B1"].Formula = "LN(EXP(1))";
				ws.Calculate();
				Assert.AreEqual(1d, ws.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void LnWithERaisedToAPowerInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["B1"].Formula = "LN(EXP(234))";
				ws.Calculate();
				Assert.AreEqual(234d, ws.Cells["B1"].Value);
			}
		}
		#endregion
	}
}
