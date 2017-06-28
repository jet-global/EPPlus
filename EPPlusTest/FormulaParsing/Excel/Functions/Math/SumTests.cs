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
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math
{
	[TestClass]
	public class SumTests : MathFunctionsTestBase
	{
		#region Sum Function (Execute) Tests
		[TestMethod]
		public void SumWithFourNumbersReturnsCorrectValue()
		{
			var function = new Sum();
			var result = function.Execute(FunctionsHelper.CreateArgs(1.5, 2, 3.5, 7), this.ParsingContext);
			Assert.AreEqual(14d, result.Result);
		}

		[TestMethod]
		public void SumWithFourNegativeNumbersReturnsCorrectValue()
		{
			var function = new Sum();
			var result = function.Execute(FunctionsHelper.CreateArgs(-1.5, -2, -3.5, -7), this.ParsingContext);
			Assert.AreEqual(-14d, result.Result);
		}

		[TestMethod]
		public void SumWithOneIntegerReturnsCorrectValue()
		{
			var function = new Sum();
			var result = function.Execute(FunctionsHelper.CreateArgs(2), this.ParsingContext);
			Assert.AreEqual(2d, result.Result);
		}

		[TestMethod]
		public void SumWithOneDoubleReturnsCorrectValue()
		{
			var function = new Sum();
			var result = function.Execute(FunctionsHelper.CreateArgs(2.5), this.ParsingContext);
			Assert.AreEqual(2.5d, result.Result);
		}

		[TestMethod]
		public void SumWithOneNumericStringReturnsCorrectValue()
		{
			var function = new Sum();
			var result = function.Execute(FunctionsHelper.CreateArgs("2"), this.ParsingContext);
			Assert.AreEqual(2d, result.Result);
		}

		[TestMethod]
		public void SumWithNonNumericStringReturnsPoundValue()
		{
			var function = new Sum();
			var result = function.Execute(FunctionsHelper.CreateArgs("word"), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SumWithDateInStringReturnsCorrectValue()
		{
			var function = new Sum();
			var result = function.Execute(FunctionsHelper.CreateArgs("6/16/2017"), this.ParsingContext);
			Assert.AreEqual(42902d, result.Result);
		}

		[TestMethod]
		public void SumWithBooleanValueReturnsCorrectValue()
		{
			var function = new Sum();
			var booleanTrue = function.Execute(FunctionsHelper.CreateArgs(true), this.ParsingContext);
			var booleanFalse = function.Execute(FunctionsHelper.CreateArgs(false), this.ParsingContext);
			Assert.AreEqual(1d, booleanTrue.Result);
			Assert.AreEqual(0d, booleanFalse.Result);
		}

		[TestMethod]
		public void SumWithErrorInputsReturnRespectiveErrors()
		{
			var func = new Sum();
			var argNA = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.NA), 5);
			var argNAME = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Name), 5);
			var argVALUE = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Value), 5);
			var argNUM = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Num), 5);
			var argDIV0 = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Div0), 5);
			var argREF = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Ref), 5);
			var resultNA = func.Execute(argNA, this.ParsingContext);
			var resultNAME = func.Execute(argNAME, this.ParsingContext);
			var resultVALUE = func.Execute(argVALUE, this.ParsingContext);
			var resultNUM = func.Execute(argNUM, this.ParsingContext);
			var resultDIV0 = func.Execute(argDIV0, this.ParsingContext);
			var resultREF = func.Execute(argREF, this.ParsingContext);
			Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)resultNA.Result).Type);
			Assert.AreEqual(eErrorType.Name, ((ExcelErrorValue)resultNAME.Result).Type);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)resultVALUE.Result).Type);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)resultNUM.Result).Type);
			Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)resultDIV0.Result).Type);
			Assert.AreEqual(eErrorType.Ref, ((ExcelErrorValue)resultREF.Result).Type);
		}

		[TestMethod]
		public void SumWithOneIntegerCellReferenceReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 2;
				worksheet.Cells["B2"].Formula = "SUM(B1)";
				worksheet.Calculate();
				Assert.AreEqual(2d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void SumWithOneDoubleCellReferenceReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 2.5;
				worksheet.Cells["B2"].Formula = "SUM(B1)";
				worksheet.Calculate();
				Assert.AreEqual(2.5d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void SumWithNumericStringCellReferenceReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "2";
				worksheet.Cells["B2"].Formula = "SUM(B1)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void SumWithNonNumericStringCellReferenceReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "word";
				worksheet.Cells["B2"].Formula = "SUM(B1)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void SumWithDateInStringCellReferenceReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "6/16/2017";
				worksheet.Cells["B2"].Formula = "SUM(B1)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void SumWithBooleanCellReferenceReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = true;
				worksheet.Cells["B2"].Formula = "SUM(B1)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void SumWithEmptyCellReferenceReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "SUM(A2)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B1"].Value);
			}
		}


		#endregion
	}
}
