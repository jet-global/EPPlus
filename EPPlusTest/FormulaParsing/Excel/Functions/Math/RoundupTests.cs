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
	public class RoundupTests : MathFunctionsTestBase
	{
		[TestMethod]
		public void RoundupWithNoArgsReturnsPoundValue()
		{
			var function = new Roundup();
			var result = function.Execute(FunctionsHelper.CreateArgs(), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void RoundupWithNoSecondInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "ROUNDUP(12.2, )";
				worksheet.Calculate();
				Assert.AreEqual(13d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void RoundupWithNoFirstInputReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "ROUNDUP(, 1)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void RoundupWithSecondInputGreaterThanZeroReturnsCorrectValue()
		{
			var function = new Roundup();
			var result = function.Execute(FunctionsHelper.CreateArgs(12.23256853, 4), this.ParsingContext);
			Assert.AreEqual(12.2326d, result.Result);
		}

		[TestMethod]
		public void RoundupWithSecondInputLessThanZeroReturnsCorrectValue()
		{
			var function = new Roundup();
			var result = function.Execute(FunctionsHelper.CreateArgs(114512, -3), this.ParsingContext);
			Assert.AreEqual(115000d, result.Result);
		}

		[TestMethod]
		public void RoundupWithSecondInputAsZeroReturnsCorrectValue()
		{
			var function = new Roundup();
			var result = function.Execute(FunctionsHelper.CreateArgs(12.568, 0), this.ParsingContext);
			Assert.AreEqual(13d, result.Result);
		}

		[TestMethod]
		public void RoundupWithSecondInputAsNumericStringReturnsCorrectValue()
		{
			var function = new Roundup();
			var result = function.Execute(FunctionsHelper.CreateArgs(25.364, "2"), this.ParsingContext);
			Assert.AreEqual(25.37d, result.Result);
		}

		[TestMethod]
		public void RoundupWithSecondInputAsGeneralStringReturnsPoundValue()
		{
			var function = new Roundup();
			var result = function.Execute(FunctionsHelper.CreateArgs(45.6, "string"), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void RoundupWithSecondInputAsDateFunctionResultReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "ROUNDUP(12.3654, DATE(2017, 6, 7))";
				worksheet.Calculate();
				Assert.AreEqual(12.3654d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void RoundupWithSecondInputAsDateAsStringReturnsCorrectValue()
		{
			var function = new Roundup();
			var result = function.Execute(FunctionsHelper.CreateArgs(12.3546, "5/5/2017"), this.ParsingContext);
			Assert.AreEqual(12.3546d, result.Result);
		}

		[TestMethod]
		public void RoundupWithSecondInputAsCellReferenceReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 2;
				worksheet.Cells["B2"].Formula = "ROUNDUP(26.32568, B1)";
				worksheet.Calculate();
				Assert.AreEqual(26.33d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void RoundupWithEmptyCellReferenceReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "ROUNDUP(A2,A3)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void RoundupWithInputsAsErrorValueReturnsRespectiveError()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["A2"].Formula = "SQRT(-1)";
				worksheet.Cells["B1"].Formula = "ROUNDUP(A2, 2)";
				worksheet.Cells["B2"].Formula = "ROUNDUP(34.45, A2)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)worksheet.Cells["B1"].Value).Type);
				Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)worksheet.Cells["B2"].Value).Type);
			}
		}

		[TestMethod]
		public void RoundupWithNumberAsCellReferenceReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 12.32568;
				worksheet.Cells["B2"].Formula = "ROUNDUP(B1, 2)";
				worksheet.Calculate();
				Assert.AreEqual(12.33d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void RoundupWithIntegerAndPositiveSecondInputReturnsCorrectValue()
		{
			var function = new Roundup();
			var result = function.Execute(FunctionsHelper.CreateArgs(25, 2), this.ParsingContext);
			Assert.AreEqual(25d, result.Result);
		}

		[TestMethod]
		public void RoundupWithDoubleAndPositiveSecondInputReturnsCorrectValue()
		{
			var function = new Roundup();
			var result = function.Execute(FunctionsHelper.CreateArgs(25.364, 2), this.ParsingContext);
			Assert.AreEqual(25.37d, result.Result);
		}

		[TestMethod]
		public void RoundupWithFractionReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "ROUNDUP((2/3), 3)";
				worksheet.Calculate();
				Assert.AreEqual(0.667d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void RoundupWithNumbericStringReturnsCorrectValue()
		{
			var function = new Roundup();
			var result = function.Execute(FunctionsHelper.CreateArgs("23.25623665", 3), this.ParsingContext);
			Assert.AreEqual(23.257d, result.Result);
		}

		[TestMethod]
		public void RoundupWithGeneralStringAsFirstInputReturnsPoundValue()
		{
			var function = new Roundup();
			var result = function.Execute(FunctionsHelper.CreateArgs("String", 3), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void RoundupWithNumberAsDateFunctionReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "ROUNDUP(DATE(2017, 6, 5), 3)";
				worksheet.Calculate();
				Assert.AreEqual(42891d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void RoundupWithNumberAsDateAsStringReturnsCorrectValue()
		{
			var function = new Roundup();
			var result = function.Execute(FunctionsHelper.CreateArgs("5/5/2017", 2), this.ParsingContext);
			Assert.AreEqual(42860d, result.Result);
		}

		[TestMethod]
		public void RoundupWithNegativeDoubleAndPositiveSecondInputReturnsCorrectValue()
		{
			var function = new Roundup();
			var result = function.Execute(FunctionsHelper.CreateArgs(-12.345, 2), this.ParsingContext);
			Assert.AreEqual(-12.35d, result.Result);
		}

		[TestMethod]
		public void RoundupWithNegativeDoubleAndNegativeSecondInputReturnsCorrectValue()
		{
			var function = new Roundup();
			var result = function.Execute(FunctionsHelper.CreateArgs(-1315.26, -2), this.ParsingContext);
			Assert.AreEqual(-1400d, result.Result);
		}

		[TestMethod]
		public void RoundupWithNegativeIntegerAndPositiveSecondInputReturnsCorrectValue()
		{
			var function = new Roundup();
			var result = function.Execute(FunctionsHelper.CreateArgs(-56, 3), this.ParsingContext);
			Assert.AreEqual(-56d, result.Result);
		}

		[TestMethod]
		public void RoundupWithNegativeIntegerAndNegativeSecondInputReturnsCorrectValue()
		{
			var function = new Roundup();
			var result = function.Execute(FunctionsHelper.CreateArgs(-45, -2), this.ParsingContext);
			Assert.AreEqual(-100d, result.Result);
		}

		[TestMethod]
		public void RoundupWithBooleanArgumentsReturnCorrectValues()
		{
			var function = new Roundup();
			var booleanTrue = function.Execute(FunctionsHelper.CreateArgs(2.35, true), this.ParsingContext);
			var booleanFalse = function.Execute(FunctionsHelper.CreateArgs(2.56, false), this.ParsingContext);
			Assert.AreEqual(2.4d, booleanTrue.Result);
			Assert.AreEqual(3d, booleanFalse.Result);
		}
	}
}
