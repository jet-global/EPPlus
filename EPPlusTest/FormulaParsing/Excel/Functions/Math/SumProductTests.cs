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
	public class SumProductTests : MathFunctionsTestBase
	{
		[TestMethod]
		public void SumProductWithIntegerInputsReturnsCorrectValue()
		{
			var function = new SumProduct();
			var result = function.Execute(FunctionsHelper.CreateArgs(4, 5, 6), this.ParsingContext);
			Assert.AreEqual(120d, result.Result);
		}

		[TestMethod]
		public void SumProductWithDoubleInputsReturnsCorrectValue()
		{
			var function = new SumProduct();
			var result = function.Execute(FunctionsHelper.CreateArgs(5.6, 5.7), this.ParsingContext);
			Assert.AreEqual(31.92d, (double)result.Result, 0.000001);
		}

		[TestMethod]
		public void SumProductWithGeneralStringInputReturnsPoundValue()
		{
			var function = new SumProduct();
			var result = function.Execute(FunctionsHelper.CreateArgs("string", "string"), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SumProductWithNumericStringInputReturnsPoundValue()
		{
			var function = new SumProduct();
			var result = function.Execute(FunctionsHelper.CreateArgs("2", "3"), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SumProductWithTrueBooleanInputReturnsPoundValue()
		{
			var function = new SumProduct();
			var result = function.Execute(FunctionsHelper.CreateArgs(true, 3), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SumProductWithFalseBooleanInputReturnsPoundValue()
		{
			var function = new SumProduct();
			var result = function.Execute(FunctionsHelper.CreateArgs(3, false), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SumProductWithDateFunctionInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "SUMPRODUCT(DATE(2017, 5, 2), 2)";
				worksheet.Calculate();
				Assert.AreEqual(85714d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void SumProductWithDateAsStringInputReturnsCorrectValue()
		{
			var function = new SumProduct();
			var result = function.Execute(FunctionsHelper.CreateArgs("5/5/2017", 2), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SumProductWithEmptyCellReferenceReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "SUMPRODUCT(A2)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)worksheet.Cells["B1"].Value).Type);
			}
		}

		[TestMethod]
		public void SumProductWithIntegerCellReferenceReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 5;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Formula = "SUMPRODUCT(B1, B2)";
				worksheet.Calculate();
				Assert.AreEqual(10d, worksheet.Cells["B3"].Value);
			}
		}

		[TestMethod]
		public void SumProductWithDoubleCellReferenceReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 2.3;
				worksheet.Cells["B2"].Value = 4.5;
				worksheet.Cells["B3"].Formula = "SUMPRODUCT(B1, B2)";
				worksheet.Calculate();
				Assert.AreEqual(10.35d, worksheet.Cells["B3"].Value);
			}
		}

		[TestMethod]
		public void SumProductWithGeneralStringCellReferenceReturnsPoundValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "String";
				worksheet.Cells["B2"].Value = "String";
				worksheet.Cells["B3"].Formula = "SUMPRODUCT(B1, B2)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)worksheet.Cells["B3"].Value).Type);
			}
		}

		[TestMethod]
		public void SumProductWithNumericStringCellReferenceReturnsPoundValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "3";
				worksheet.Cells["B2"].Value = "4";
				worksheet.Cells["B3"].Formula = "SUMPRODUCT(B1, B2)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)worksheet.Cells["B3"].Value).Type);
			}
		}

		[TestMethod]
		public void SumProductWithTrueBooleanCellReferenceReturnsPoundValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = true;
				worksheet.Cells["B2"].Value = 4;
				worksheet.Cells["B3"].Formula = "SUMPRODUCT(B1, B2)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)worksheet.Cells["B3"].Value).Type);
			}
		}

		[TestMethod]
		public void SumProductWithFalseBooleanCellReferenceReturnsPoundValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 3;
				worksheet.Cells["B2"].Value = false;
				worksheet.Cells["B3"].Formula = "SUMPRODUCT(B1, B2)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)worksheet.Cells["B3"].Value).Type);
			}
		}

		[TestMethod]
		public void SumProductWithDateFunctionCellReferenceReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "DATE(2017, 6, 15)";
				worksheet.Cells["B2"].Value = 4;
				worksheet.Cells["B3"].Formula = "SUMPRODUCT(B1, B2)";
				worksheet.Calculate();
				Assert.AreEqual(171604d, worksheet.Cells["B3"].Value);
			}
		}

		[TestMethod]
		public void SumProductWithDateAsStringCellReferenceReturnsPoundValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "5/5/2017";
				worksheet.Cells["B2"].Value = 4;
				worksheet.Cells["B3"].Formula = "SUMPRODUCT(B1, B2)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)worksheet.Cells["B3"].Value).Type);
			}
		}

		[TestMethod]
		public void SumProductWithIntegerArrayInputReurnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 5;
				worksheet.Cells["B2"].Value = 6;
				worksheet.Cells["B3"].Value = 7;
				worksheet.Cells["B4"].Value = 8;
				worksheet.Cells["B5"].Formula = "SUMPRODUCT(B1:B2, B3:B4)";
				worksheet.Calculate();
				Assert.AreEqual(83d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void SumProdctWithDoubleArrayInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 5.6;
				worksheet.Cells["B2"].Value = 3.5;
				worksheet.Cells["B3"].Value = 4.6;
				worksheet.Cells["B4"].Value = 4.2;
				worksheet.Cells["B5"].Formula = "SUMPRODUCT(B1:B2, B3:B4)";
				worksheet.Calculate();
				Assert.AreEqual(40.46d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void SumProductWithGeneralStringArrayInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "String";
				worksheet.Cells["B2"].Value = "String";
				worksheet.Cells["B3"].Value = "String";
				worksheet.Cells["B4"].Value = "String";
				worksheet.Cells["B5"].Formula = "SUMPRODUCT(B1:B2, B3:B4)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void SumProductWithNumericStringArrayInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "2";
				worksheet.Cells["B2"].Value = "4";
				worksheet.Cells["B3"].Value = "3";
				worksheet.Cells["B4"].Value = "3";
				worksheet.Cells["B5"].Formula = "SUMPRODUCT(B1:B2, B3:B4)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void SumProductWithTrueBooleanArrayInputReturnCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = true;
				worksheet.Cells["B2"].Value = 3;
				worksheet.Cells["B3"].Value = 6;
				worksheet.Cells["B4"].Value = true;
				worksheet.Cells["B5"].Formula = "SUMPRODUCT(B1:B2, B3:B4)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void SumProductWithFalseBooleanArrayInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 4;
				worksheet.Cells["B2"].Value = false;
				worksheet.Cells["B3"].Value = 9;
				worksheet.Cells["B4"].Value = false;
				worksheet.Cells["B5"].Formula = "SUMPRODUCT(B1:B2, B3:B4)";
				worksheet.Calculate();
				Assert.AreEqual(36d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void SumProductWithMixedBooleanArrayInputsReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = true;
				worksheet.Cells["B2"].Value = false;
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["B4"].Value = false;
				worksheet.Cells["B5"].Formula = "SUMPRODUCT(B1:B2, B3:B4)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void SumProductWithDifferentArraySizeInputsReturnsPoundValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 5;
				worksheet.Cells["B2"].Value = 6;
				worksheet.Cells["B3"].Value = 7;
				worksheet.Cells["B4"].Value = 8;
				worksheet.Cells["B5"].Value = 3;
				worksheet.Cells["B6"].Formula = "SUMPRODUCT(B1:B2, B3:B5)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)worksheet.Cells["B6"].Value).Type);
			}
		}

		[TestMethod]
		public void SumProductWithArrayOfMixedTypesReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "4";
				worksheet.Cells["B2"].Value = 4;
				worksheet.Cells["B3"].Value = true;
				worksheet.Cells["B4"].Value = "string";
				worksheet.Cells["B5"].Formula = "SUMPRODUCT(B1:B2, B3:B4)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void SumProductWithDateFunctionArrayInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "DATE(2017, 6, 2)";
				worksheet.Cells["B2"].Formula = "DATE(2017, 7, 1)";
				worksheet.Cells["B3"].Formula = "DATE(2017, 6, 21)";
				worksheet.Cells["B4"].Formula = "DATE(2017, 5, 5)";
				worksheet.Cells["B5"].Formula = "SUMPRODUCT(B1:B2, B3:B4)";
				worksheet.Calculate();
				Assert.AreEqual(3679618036d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void SumProductWithDateAsStringArrayInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "4/5/2017";
				worksheet.Cells["B2"].Value = "5/5/2017";
				worksheet.Cells["B3"].Value = "6/9/2015";
				worksheet.Cells["B4"].Value = "8/8/2015";
				worksheet.Cells["B5"].Formula = "SUMPRODUCT(B1:B2, B3:B4)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void SumProductWithAnyMissingArgumentReturnsPoundValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "SUMPRODUCT(3, , 5)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)worksheet.Cells["B1"].Value).Type);
			}
		}

		[TestMethod]
		public void SumProductWithArgumentArrayReferenceMissingReturnsPoundValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 4;
				worksheet.Cells["B2"].Value = 5;
				worksheet.Cells["B3"].Value = 5;
				worksheet.Cells["B4"].Value = 6;
				worksheet.Cells["B5"].Formula = "SUMPRODUCT(B1:B2, , B3:B4)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)worksheet.Cells["B5"].Value).Type);
			}
		}

		[TestMethod]
		public void SumProductWithFirstArgumentMissingReturnsPoundValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "SUMPRODUCT(, 3, 4)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)worksheet.Cells["B1"].Value).Type);
			}
		}

		[TestMethod]
		public void SumProductWithReferenceToErrorReturnsRespectiveError()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "notaformula";
				worksheet.Cells["B2"].Formula = "SUMPRODUCT(B1)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Name, ((ExcelErrorValue)worksheet.Cells["B2"].Value).Type);
			}
		}

		[TestMethod]
		public void SumProductWithErrorInArrayInputReturnsRespecitveError()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 3;
				worksheet.Cells["B2"].Formula = "notaformula";
				worksheet.Cells["B3"].Value = 5;
				worksheet.Cells["B4"].Value = 7;
				worksheet.Cells["B5"].Formula = "SUMPRODUCT(B1:B2, B4:B4)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Name, ((ExcelErrorValue)worksheet.Cells["B5"].Value).Type);
			}
		}

		[TestMethod]
		public void SumProductWithArrayDirectInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "SUMPRODUCT({4, 5, 6}, {4, 2, 4})";
				worksheet.Calculate();
				Assert.AreEqual(50d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void SumProductWithInvalidArgumentReturnsPoundValue()
		{
			var func = new SumProduct();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

	}
}
