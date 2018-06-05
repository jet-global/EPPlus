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
	public class MinTests : MathFunctionsTestBase
	{
		#region Min Function (Execute) Tests
		[TestMethod]
		public void MinWithNoArgumentsReturnsPoundValue()
		{
			var function = new Min();
			var result = function.Execute(FunctionsHelper.CreateArgs(), this.ParsingContext);
		}

		[TestMethod]
		public void MinWithReferenceToEmptyCellReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "MIN(A2:A4)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void MinWithReferenceToCellWithLogicalValueReturnsCorrectResult()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 5;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = "TRUE";
				worksheet.Cells["B4"].Formula = "MIN(B1:B3)";
				worksheet.Calculate();
				Assert.AreEqual(2d, worksheet.Cells["B4"].Value);
			}
		}

		[TestMethod]
		public void MinWithReferenceToCellWithNumericStringReturnsCorrectResult()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "1";
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 5;
				worksheet.Cells["B4"].Formula = "MIN(B1:B3)";
				worksheet.Calculate();
				Assert.AreEqual(2d, worksheet.Cells["B4"].Value);
			}
		}

		[TestMethod]
		public void MinWithArrayWithLogicalValueReturnsCorrectResult()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "MIN({5, 6, TRUE, 2})";
				worksheet.Calculate();
				Assert.AreEqual(2d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void MinWithArrayWithNumericStringReturnsCorrectResult()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "MIN({5, 6, \"2\"})";
				worksheet.Calculate();
				Assert.AreEqual(5d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void MinWithLargeRangeReturnsCorrectValue()
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
				worksheet.Cells["C2"].Formula = "MIN(A1:A270)";
				worksheet.Cells["C3"].Formula = "MIN(A1:A270, 5, 6, 7, 8)";
				worksheet.Cells["C4"].Formula = "MIN(A1:A270, 5, 6, -3, 8)";
				worksheet.Calculate();
				Assert.AreEqual(4d, worksheet.Cells["C2"].Value);
				Assert.AreEqual(4d, worksheet.Cells["C3"].Value);
				Assert.AreEqual(-3d, worksheet.Cells["C4"].Value);
			}
		}

		[TestMethod]
		public void MinWithReferenceToCellsWithStringsReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "string";
				worksheet.Cells["B2"].Value = "string";
				worksheet.Cells["B3"].Value = "string";
				worksheet.Cells["B4"].Formula = "MIN(B1:B3)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B4"].Value);
			}
		}

		[TestMethod]
		public void MinWithReferenceToDateObjectsReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "DATE(2017, 5, 12)";
				worksheet.Cells["B2"].Formula = "DATE(2017, 6, 2)";
				worksheet.Cells["B3"].Formula = "DATE(2017, 5, 15)";
				worksheet.Cells["B4"].Formula = "MIN(B1:B3)";
				worksheet.Calculate();
				Assert.AreEqual(42867d, worksheet.Cells["B4"].Value);
			}
		}

		[TestMethod]
		public void MinWithReferenceToCellWithDateObjectAsOADate()
		{
			var function = new Min();
			var dateObject1 = new DateTime(2017, 5, 12).ToOADate();
			var dateObject2 = new DateTime(2017, 6, 2).ToOADate();
			var dateObject3 = new DateTime(2017, 5, 15).ToOADate();
			var result = function.Execute(FunctionsHelper.CreateArgs(dateObject1, dateObject2, dateObject3), this.ParsingContext);
			Assert.AreEqual(42867d, result.Result);
		}

		[TestMethod]
		public void MinWithReferenceToCellsWithDatesAsStringsReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "5/2/2017";
				worksheet.Cells["B2"].Value = "6/10/2017";
				worksheet.Cells["B3"].Formula = "MIN(B1:B2)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B3"].Value);
			}
		}

		[TestMethod]
		public void MinWithArrayWithStringsReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "MIN({\"string\", \"string\"})";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void MinWithArrayWithDatesAsStringsReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "MIN({\"5/2/2017\", \"6/10/2017\"})";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void MinIgnoresNonNumericReferencedStrings()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells[2, 2].Value = "abc";
				worksheet.Cells[2, 3].Value = 123;
				worksheet.Cells[2, 6].Formula = "MIN(B2,C2)";
				worksheet.Calculate();
				Assert.AreEqual(123d, worksheet.Cells[2, 6].Value);
			}
		}

		[TestMethod]
		public void MinWithEmptyValueReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells[2, 3].Value = 123;
				worksheet.Cells[2, 4].Value = 345;
				worksheet.Cells[2, 6].Formula = "MIN(B2,C2,D2)";
				worksheet.Calculate();
				Assert.AreEqual(123d, worksheet.Cells[2, 6].Value);
			}
		}

		[TestMethod]
		public void MinWithIntegerInputReturnsCorrectValue()
		{
			var function = new Min();
			var result = function.Execute(FunctionsHelper.CreateArgs(5, 9, 66, 13), this.ParsingContext);
			Assert.AreEqual(5d, result.Result);
		}

		[TestMethod]
		public void MinWithDoublesInputReturnsCorrectValue()
		{
			var function = new Min();
			var result = function.Execute(FunctionsHelper.CreateArgs(5.6, 23.89, 11.5, 2.6), this.ParsingContext);
			Assert.AreEqual(2.6d, result.Result);
		}

		[TestMethod]
		public void MinWithFractionsInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "MIN((9/8), (7/9), (1/3))";
				worksheet.Calculate();
				Assert.AreEqual(0.33333333d, (double)worksheet.Cells["B1"].Value, 0.00000001);
			}
		}

		[TestMethod]
		public void MinWithStringsInputReturnsPoundValue()
		{
			var function = new Min();
			var result = function.Execute(FunctionsHelper.CreateArgs("string", "string"), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void MinWithDateObjectsInputReturnsCorrectValue()
		{
			var function = new Min();
			var dateObject1 = new DateTime(2017, 6, 15);
			var dateObject2 = new DateTime(2017, 5, 18);
			var result = function.Execute(FunctionsHelper.CreateArgs(dateObject1, dateObject2), this.ParsingContext);
			Assert.AreEqual(42873d, result.Result);
		}

		[TestMethod]
		public void MinWithDatesAsStringsInputReturnsCorrectValue()
		{
			var function = new Min();
			var result = function.Execute(FunctionsHelper.CreateArgs("5/2/2017", "6/25/2017"), this.ParsingContext);
			Assert.AreEqual(42857d, result.Result);
		}

		[TestMethod]
		public void MinWithMixTypesReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = true;
				worksheet.Cells["B2"].Value = "string";
				worksheet.Cells["B3"].Value = 5;
				worksheet.Cells["B4"].Value = 3;
				worksheet.Cells["B5"].Formula = "MIN(B1:B4)";
				worksheet.Cells["B6"].Formula = "MIN(TRUE, \"STRING\", 5, 3)";
				worksheet.Calculate();
				Assert.AreEqual(3d, worksheet.Cells["B5"].Value);
				Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)worksheet.Cells["B6"].Value).Type);
			}
		}

		[TestMethod]
		public void MinWithLogicalValInArgumentListReturnsCorrectValue()
		{
			var function = new Min();
			var result = function.Execute(FunctionsHelper.CreateArgs(true, 5, 66, 8), this.ParsingContext);
			Assert.AreEqual(1d, result.Result);
		}

		[TestMethod]
		public void MinShouldCalculateCorrectResult()
		{
			var function = new Min();
			var arguments = FunctionsHelper.CreateArgs(4, 2, 5, 2);
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(2d, result.Result);
		}

		[TestMethod]
		public void MinShouldIgnoreHiddenValuesIfIgnoreHiddenValuesIsTrue()
		{
			var function = new Min();
			function.IgnoreHiddenValues = true;
			var arguments = FunctionsHelper.CreateArgs(4, 2, 5, 3);
			arguments.ElementAt(1).SetExcelStateFlag(ExcelCellState.HiddenCell);
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(3d, result.Result);
		}

		[TestMethod]
		public void MinWithInvalidArgumentReturnsPoundValue()
		{
			var function = new Min();
			var arguments = FunctionsHelper.CreateArgs();
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void MinWithOneArgumentReturnsCorrectValue()
		{
			var function = new Min();
			var result = function.Execute(FunctionsHelper.CreateArgs(10), this.ParsingContext);
			Assert.AreEqual(10d, result.Result);
		}
		#endregion
	}
}
