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
using OfficeOpenXml.FormulaParsing.Excel;
using System.Linq;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math
{
	[TestClass]
	public class MaxTests : MathFunctionsTestBase
	{
		#region Max Function (Execute) Tests
		[TestMethod]
		public void MaxWithNoInputsReturnsPoundValue()
		{
			var function = new Max();
			var result = function.Execute(FunctionsHelper.CreateArgs(), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void MaxWithIntegerInputReturnsCorrectValue()
		{
			var function = new Max();
			var result = function.Execute(FunctionsHelper.CreateArgs(12, 2, 99, 1, 25), this.ParsingContext);
			Assert.AreEqual(99d, result.Result);
		}

		[TestMethod]
		public void MaxWithDoublesInputReturnsCorrectValue()
		{
			var function = new Max();
			var result = function.Execute(FunctionsHelper.CreateArgs(2.5, 26.5, 123.87, 658.64, 55.6), this.ParsingContext);
			Assert.AreEqual(658.64d, result.Result);
		}

		[TestMethod]
		public void MaxWithFractionsInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "MAX((2/3),(9/8),(2/55))";
				worksheet.Calculate();
				Assert.AreEqual(1.125d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void MaxWithStringInputReturnsPoundValue()
		{
			var function = new Max();
			var result = function.Execute(FunctionsHelper.CreateArgs("string", "string"), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void MaxWithDateObjectInputsReturnsCorrectValue()
		{
			var function = new Max();
			var dateObject1 = new DateTime(2017, 6, 2);
			var dateObject2 = new DateTime(2017, 6, 15);
			var dateObjectAsOADate1 = new DateTime(2017, 6, 2).ToOADate();
			var dateObjectAsOADate2 = new DateTime(2017, 6, 15).ToOADate();
			
			var result1 = function.Execute(FunctionsHelper.CreateArgs(dateObject1, dateObject2), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(dateObjectAsOADate1, dateObjectAsOADate2), this.ParsingContext);
			Assert.AreEqual(42901d, result1.Result);
			Assert.AreEqual(42901d, result2.Result);
		}

		[TestMethod]
		public void MaxWithDatesAsStringsInputReturnsCorrectValue()
		{
			var function = new Max();
			var result = function.Execute(FunctionsHelper.CreateArgs("5/2/2017", "6/25/2017"), this.ParsingContext);
			Assert.AreEqual(42911d, result.Result);
		}

		[TestMethod]
		public void MaxWithReferenceToEmptyCellReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "MAX(B2:B4)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void MaxWithReferenceToCellWithLogicalValueReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 0.5;
				worksheet.Cells["B2"].Value = 0.2;
				worksheet.Cells["B3"].Value = true;
				worksheet.Cells["B10"].Formula = "MAX(B1:B3)";
				worksheet.Calculate();
				Assert.AreEqual(0.5d, worksheet.Cells["B10"].Value);
			}
		}

		[TestMethod]
		public void MaxWithReferenceToCellWithNumericStringReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 0.5;
				worksheet.Cells["B2"].Value = "1";
				worksheet.Cells["B3"].Value = 0.2;
				worksheet.Cells["B4"].Formula = "MAX(B1:B3)";
				worksheet.Calculate();
				Assert.AreEqual(0.5d, worksheet.Cells["B4"].Value);
			}
		}

		[TestMethod]
		public void MaxWithArrayWithLogicalValueReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "MAX({0.5, 0.1, TRUE})";
				worksheet.Calculate();
				Assert.AreEqual(0.5d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void MaxWithArrayWithNumericStringReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "MAX({1, 3, \"78\"})";
				worksheet.Calculate();
				Assert.AreEqual(3d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void MaxWithReferenceToCellsWithStringsReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "string";
				worksheet.Cells["B2"].Value = "String";
				worksheet.Cells["B3"].Formula = "MAX(B1:B2)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B3"].Value);
			}
		}

		[TestMethod]
		public void MaxWithReferenceToCellsWithDateObjectsReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "DATE(2017, 5, 12)";
				worksheet.Cells["B2"].Formula = "DATE(2017, 6, 2)";
				worksheet.Cells["B3"].Formula = "DATE(2017, 5, 15)";
				worksheet.Cells["B4"].Value = 42867;
				worksheet.Cells["B5"].Value = 42888;
				worksheet.Cells["B6"].Value = 42870;
				worksheet.Cells["B7"].Formula = "MAX(B1:B3)";
				worksheet.Cells["B8"].Formula = "MAX(B4:B6)";
				worksheet.Calculate();
				Assert.AreEqual(42888d, worksheet.Cells["B7"].Value);
				Assert.AreEqual(42888d, worksheet.Cells["B8"].Value);
			}
		}

		[TestMethod]
		public void MaxWithReferenceToCellsWithDatesAsStringsReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "5/2/2017";
				worksheet.Cells["B2"].Value = "6/5/2017";
				worksheet.Cells["B3"].Value = "6/2/2017";
				worksheet.Cells["B4"].Formula = "MAX(B1:B3)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B4"].Value);
			}
		}

		[TestMethod]
		public void MaxWithArrayWithStringsReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "MAX({\"string\", \"string\", \"string\"})";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void MaxWithDecimalArgumentTest()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B2"].Formula = "MAX(C3)";
				worksheet.Cells["C3"].Value = (decimal)26.000;
				worksheet.Calculate();
				Assert.AreEqual(26d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void MaxWithDecimalArgumentsRangeTest()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B2"].Formula = "MAX(C3:G3)";
				worksheet.Cells["C3"].Value = (decimal)26.000;
				worksheet.Cells["D3"].Value = (double)19.000;
				worksheet.Cells["E3"].Value = (decimal)43.020;
				worksheet.Cells["G3"].Value = (int)12;
				worksheet.Calculate();
				Assert.AreEqual(43.020d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void MaxWithMultipleRangeArgumentsRangeTest()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B2"].Formula = "MAX(C3:E3, F3:H3)";
				worksheet.Cells["C3"].Value = 1;
				worksheet.Cells["D3"].Value = 2;
				worksheet.Cells["E3"].Value = "34";
				worksheet.Cells["F3"].Value = 4;
				worksheet.Cells["G3"].Value = "15";
				worksheet.Cells["H3"].Value = 6;
				worksheet.Calculate();
				Assert.AreEqual(6d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void MaxWithDateTimeArgumentsRangeTest()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B2"].Formula = "MAX(C3:G3)";
				worksheet.Cells["C3"].Value = (decimal)42500.00;
				worksheet.Cells["D3"].Value = new DateTime(2016, 5, 23); // 42513 in OADate.
				worksheet.Cells["E3"].Value = (double)42000.020;
				worksheet.Cells["G3"].Value = (int)42512;
				worksheet.Calculate();
				Assert.AreEqual(42513d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void MaxWithArrayWithDatesAsStringsReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "MAX({\"5/2/2017\", \"6/2/2017\"})";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void MaxWithArrayWithMixedTypesReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "TRUE";
				worksheet.Cells["B2"].Value = "String";
				worksheet.Cells["B3"].Value = 5;
				worksheet.Cells["B4"].Value = 8;
				worksheet.Cells["B5"].Formula = "MAX(B1:B4)";
				worksheet.Cells["B6"].Formula = "MAX({TRUE, \"string\", 3, 5, 8})";
				worksheet.Cells["B7"].Formula = "MAX(TRUE, 0.6, 0.7)";
				worksheet.Cells["B8"].Formula = "MAX(TRUE, \"string\", 0.8, 1)";
				worksheet.Calculate();
				Assert.AreEqual(8d, worksheet.Cells["B5"].Value);
				Assert.AreEqual(8d, worksheet.Cells["B6"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B7"].Value);
				Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)worksheet.Cells["B8"].Value).Type);
			}
		}

		[TestMethod]
		public void MaxIgnoresNonNumericReferencedStrings()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells[2, 2].Value = "abc";
				worksheet.Cells[2, 3].Value = 123;
				worksheet.Cells[2, 6].Formula = "MAX(B2,C2)";
				worksheet.Calculate();
				Assert.AreEqual(123d, worksheet.Cells[2, 6].Value);
			}
		}

		[TestMethod]
		public void MaxWithEmptyValueReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells[2, 3].Value = 123;
				worksheet.Cells[2, 4].Value = 345;
				worksheet.Cells[2, 6].Formula = "MAX(B2,C2,D2)";
				worksheet.Calculate();
				Assert.AreEqual(345d, worksheet.Cells[2, 6].Value);
			}
		}

		[TestMethod]
		public void MaxWithLargeRangeReturnsCorrectResult()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				for (int i = 1; i < 270; i++)
				{
					for (int j = 1; j < 2; j++)
					{
						worksheet.Cells[i, j].Value = 4;
					}
				}
				worksheet.Cells["C1"].Formula = "MAX(A1:A270)";
				worksheet.Cells["C2"].Formula = "MAX(A1:A270, 1, 2, 3, 1.5)";
				worksheet.Cells["C3"].Formula = "MAX(A1:A270, 1, 2, 37, 1.5)";
				worksheet.Calculate();
				Assert.AreEqual(4d, worksheet.Cells["C1"].Value);
				Assert.AreEqual(4d, worksheet.Cells["C2"].Value);
				Assert.AreEqual(37d, worksheet.Cells["C3"].Value);
			}
		}

		[TestMethod]
		public void MaxShouldCalculateCorrectResult()
		{
			var function = new Max();
			var arguments = FunctionsHelper.CreateArgs(4, 2, 5, 2);
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(5d, result.Result);
		}

		[TestMethod]
		public void MaxShouldIgnoreHiddenValuesIfIgnoreHiddenValuesIsTrue()
		{
			var function = new Max();
			function.IgnoreHiddenValues = true;
			var args = FunctionsHelper.CreateArgs(4, 2, 5, 2);
			args.ElementAt(2).SetExcelStateFlag(ExcelCellState.HiddenCell);
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(4d, result.Result);
		}

		[TestMethod]
		public void MaxWithInvalidArgumentReturnsPoundValue()
		{
			var function = new Max();
			var arguments = FunctionsHelper.CreateArgs();
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}
		#endregion
	}
}
