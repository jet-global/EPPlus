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
	public class RoundTests : MathFunctionsTestBase
	{
		#region Round Function (Execute) Tests
		[TestMethod]
		public void RoundWithNoInputsReturnsPoundValue()
		{
			var function = new Round();
			var result = function.Execute(FunctionsHelper.CreateArgs(), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void RoundWithNoSecondInputReturnsCorrectNumber()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "ROUND(10, )";
				worksheet.Calculate();
				Assert.AreEqual(10d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void RoundWithNoFirstInputReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B2"].Formula = "ROUND(,3)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void RoundWithSecondInputGreaterThanZeroReturnsCorrectValue()
		{
			var function = new Round();
			var result = function.Execute(FunctionsHelper.CreateArgs(10.6673, 3), this.ParsingContext);
			Assert.AreEqual(10.667d, result.Result);
		}

		[TestMethod]
		public void RoundWithSecondInputLessThanZeroReturnsCorrectValue()
		{
			var function = new Round();
			var result = function.Execute(FunctionsHelper.CreateArgs(626.3, -3), this.ParsingContext);
			Assert.AreEqual(1000d, result.Result);
		}

		[TestMethod]
		public void RoundWithSecondInputIsZeroReturnsCorrectValue()
		{
			var function = new Round();
			var result = function.Execute(FunctionsHelper.CreateArgs(56.8, 0), this.ParsingContext);
			Assert.AreEqual(57d, result.Result);
		}

		[TestMethod]
		public void RoundWithSecondInputAsNumericStringReturnsCorrectValue()
		{
			var function = new Round();
			var result = function.Execute(FunctionsHelper.CreateArgs(10.5435, "2"), this.ParsingContext);
			Assert.AreEqual(10.54d, result.Result);
		}

		[TestMethod]
		public void RoundWithSecondInputAsGeneralStringReturnsCorrectValue()
		{
			var function = new Round();
			var result = function.Execute(FunctionsHelper.CreateArgs(10.3431, "string"), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void RoundWithSecondInputAsDateAsStringReturnsCorrectValue()
		{
			var function = new Round();
			var result = function.Execute(FunctionsHelper.CreateArgs(563.7, "5/5/2012"), this.ParsingContext);
			Assert.AreEqual(563.7d, result.Result);
		}

		[TestMethod]
		public void RoundWithSecondInputAsCellReferenceReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 3;
				worksheet.Cells["B2"].Formula = "ROUND(12.36589, B1)";
				worksheet.Calculate();
				Assert.AreEqual(12.366d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void RoundWithEmptyCellReferenceReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "ROUND(A2, A3)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void RoundWithFirstInputAsCellReferenceReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 12.36589;
				worksheet.Cells["B2"].Formula = "ROUND(B1, 3)";
				worksheet.Calculate();
				Assert.AreEqual(12.366d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void RoundWithErrorValueAsInputReturnsRespectiveError()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "SQRT(-1)";
				worksheet.Cells["B2"].Formula = "ROUND(B1, 2)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)worksheet.Cells["B2"].Value).Type);
			}
		}

		[TestMethod]
		public void RoundWithIntegerAndPositiveSecondArgReturnsCorrectValue()
		{
			var function = new Round();
			var result = function.Execute(FunctionsHelper.CreateArgs(12, 5), this.ParsingContext);
			Assert.AreEqual(12d, result.Result);
		}

		[TestMethod]
		public void RoundWithDoubleReturnsCorrectValue()
		{
			var function = new Round();
			var result = function.Execute(FunctionsHelper.CreateArgs(12.5623, 2), this.ParsingContext);
			Assert.AreEqual(12.56d, result.Result);
		}

		[TestMethod]
		public void RoundWithFractionInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "ROUND((2/3), 3)";
				worksheet.Calculate();
				Assert.AreEqual(0.667, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void RoundWithFirstInputAsNumericStringReturnsCorrectValue()
		{
			var function = new Round();
			var result = function.Execute(FunctionsHelper.CreateArgs("234.5667", 2), this.ParsingContext);
			Assert.AreEqual(234.57d, result.Result);
		}

		[TestMethod]
		public void RoundWithFirstInputAsGeneralStringReturnsPoundValue()
		{
			var function = new Round();
			var result = function.Execute(FunctionsHelper.CreateArgs("string", 3), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void RoundWithFirstInputAsDateAsStringReturnsCorrectValue()
		{
			var function = new Round();
			var result = function.Execute(FunctionsHelper.CreateArgs("5/5/2014", 3), this.ParsingContext);
			Assert.AreEqual(41764d, result.Result);
		}

		[TestMethod]
		public void RoundWithSecondArgumentAsDateObjectReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B2"].Formula = "ROUND(12.5, DATE(2013, 6, 2))";
				worksheet.Calculate();
				Assert.AreEqual(12.5d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void RoundWithFirstArgumentAsDateObjectReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "DATE(2017, 5, 2)";
				worksheet.Cells["B2"].Formula = "ROUND(DATE(2017, 5, 2), 3)";
				worksheet.Calculate();
				Assert.AreEqual(worksheet.Cells["B1"].Value, worksheet.Cells["B2"].Value);
			}
		}
		#endregion
	}
}
