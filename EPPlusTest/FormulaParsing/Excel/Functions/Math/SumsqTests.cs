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
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math
{
	[TestClass]
	public class SumsqTests : MathFunctionsTestBase
	{
		[TestMethod]
		public void SumsqWithIntegerInputsReturnsCorrectValue()
		{
			var function = new Sumsq();
			var result = function.Execute(FunctionsHelper.CreateArgs(2, 5), this.ParsingContext);
			Assert.AreEqual(29d, result.Result);
		}

		[TestMethod]
		public void SumsqWithDoubleInputsReturnsCorrectValue()
		{
			var function = new Sumsq();
			var result = function.Execute(FunctionsHelper.CreateArgs(2.5, 6.3), this.ParsingContext);
			Assert.AreEqual(45.94d, result.Result);
		}

		[TestMethod]
		public void SumsqWithGeneralStringInputReturnsCorrectValue()
		{
			var function = new Sumsq();
			var result = function.Execute(FunctionsHelper.CreateArgs("string", "string"), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SumsqWithNumericStringInputsReturnsCorrectValue()
		{
			var function = new Sumsq();
			var result = function.Execute(FunctionsHelper.CreateArgs("2", "4"), this.ParsingContext);
			Assert.AreEqual(20d, result.Result);
		}

		[TestMethod]
		public void SumsqWithDateFunctionInputsReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "SUMSQ(DATE(2017, 6, 15), 2)";
				worksheet.Calculate();
				Assert.AreEqual(1840495805d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void SumsqWithDatesAsStringsInputReturnsCorrectValue()
		{
			var function = new Sumsq();
			var result = function.Execute(FunctionsHelper.CreateArgs("5/5/2017", 6), this.ParsingContext);
			Assert.AreEqual(1836979636d, result.Result);
		}

		[TestMethod]
		public void SumsqWithLogicalValuesInputReturnsCorrectResult()
		{
			var function = new Sumsq();
			var booleanFalse = function.Execute(FunctionsHelper.CreateArgs(false, 3, 6), this.ParsingContext);
			var booleanTrue = function.Execute(FunctionsHelper.CreateArgs(true, 3, 4), this.ParsingContext);
			Assert.AreEqual(45d, booleanFalse.Result);
			Assert.AreEqual(26d, booleanTrue.Result);
		}

		[TestMethod]
		public void SumsqWithCellReferenceInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 5;
				worksheet.Cells["B2"].Formula = "SUMSQ(B1, 4)";
				worksheet.Calculate();
				Assert.AreEqual(41d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void SumsqWithIntegerCellReferenceReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 0;
			}
		}
	}
}
