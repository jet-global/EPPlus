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
	public class MaxaTests : MathFunctionsTestBase
	{
		#region Maxa Function (Execute) Tests
		[TestMethod]
		public void MaxaWithNoArgumentsReturnsPoundValue()
		{
			var function = new Maxa();
			var result = function.Execute(FunctionsHelper.CreateArgs(), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void MaxaWithMaxInputReturnsCorrectValue()
		{
			//This functionality is different than that of Excel's. Normally when entering too many arguments
			//Excel won't let you compute the function however in EPPlus it will return a #NA!
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				for (int i = 1; i < 260; i++)
				{
					for (int j = 1; j < 2; j++)
					{
						worksheet.Cells[i, j].Value = 5;
					}
				}
				worksheet.Cells["A10"].Value = 25;
				worksheet.Cells["C1"].Formula = "MAXA(A1:A255)";
				worksheet.Cells["C2"].Formula = "MAXA(A1:A259)";
				worksheet.Calculate();
				Assert.AreEqual(25d, worksheet.Cells["C1"].Value);
				Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)worksheet.Cells["C2"].Value).Type);
			}
		}

		[TestMethod]
		public void MaxaWithReferenceToEmptyCellReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "MAXA(A2:A3)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void MaxaWithReferenceToCellWithLogicalValueReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = true;
				worksheet.Cells["B2"].Value = 0.5;
				worksheet.Cells["B3"].Formula = "FALSE";
				worksheet.Cells["B4"].Formula = "MAXA(B1:B3)";
				worksheet.Calculate();
				Assert.AreEqual(1d, worksheet.Cells["B4"].Value);
			}
		}

		[TestMethod]
		public void MaxaWithReferenceToCellWithNumericStringReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 2;
				worksheet.Cells["B2"].Value = 5;
				worksheet.Cells["B3"].Value = "1";
				worksheet.Cells["B4"].Formula = "MAXA(B1:B3)";
				worksheet.Calculate();
				Assert.AreEqual(5d, worksheet.Cells["B4"].Value);
			}
		}

		[TestMethod]
		public void MaxaWithReferenceToCelWithGenericStringReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "string";
				worksheet.Cells["B2"].Value = "String";
				worksheet.Cells["B3"].Formula = "MAXA(B1:B2)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B3"].Value);
			}
		}
		
		[TestMethod]
		public void MaxaWithReferenceToCellsWithDateObjectsReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "DATE(2017, 6, 2)";
				worksheet.Cells["B2"].Formula = "DATE(2017, 8, 15)";
				worksheet.Cells["B3"].Formula = "MAXA(B1:B2)";
				worksheet.Calculate();
				Assert.AreEqual(42962d, worksheet.Cells["B3"].Value);
			}
		}

		[TestMethod]
		public void MaxaWithReferenceToCellsWithDatesAsStringsReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "6/2/2017";
				worksheet.Cells["B2"].Value = "8/15/2017";
				worksheet.Cells["B3"].Formula = "MAXA(B1:B2)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B3"].Value);
			}
		}

		[TestMethod]
		public void MaxaWithReferenceToDoublesReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 1.5;
				worksheet.Cells["B2"].Value = 6.3;
				worksheet.Cells["B3"].Value = 0.26;
				worksheet.Cells["B4"].Formula = "MAXA(B1:B3)";
				worksheet.Calculate();
				Assert.AreEqual(6.3d, worksheet.Cells["B4"].Value);
			}
		}

		[TestMethod]
		public void MaxaWithReferenceToIntegersReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 8;
				worksheet.Cells["B2"].Value = 5;
				worksheet.Cells["B3"].Value = 2;
				worksheet.Cells["B4"].Formula = "MAXA(B1:B3)";
				worksheet.Calculate();
				Assert.AreEqual(8d, worksheet.Cells["B4"].Value);
			}
		}

		[TestMethod]
		public void MaxaWithReferenceToMixedTypesReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = true;
				worksheet.Cells["B2"].Value = 0.5;
				worksheet.Cells["B3"].Formula = "MAXA(B1:B2)";
				worksheet.Calculate();
				Assert.AreEqual(1d, worksheet.Cells["B3"].Value);
			}
		}

		[TestMethod]
		public void MaxaWithArrayWithLogicalValueReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "MAXA({0.5, 0.2, TRUE})";
				worksheet.Calculate();
				Assert.AreEqual(0.5d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void MaxaWithArrayWithNumericStringReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "MAXA({0.55, 0.2, \"1\"})";
				worksheet.Calculate();
				Assert.AreEqual(0.55d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void MaxaWithArrayWithStringsReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "MAXA({3, 6, \"string\"})";
				worksheet.Calculate();
				Assert.AreEqual(6d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void MaxaWithArrayWithDatesAsStringsReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "MAXA({\"6/2/2017\", \"5/15/2017\"})";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void MaxaWithArrayWithDoublesReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "MAXA({1.5, 0.3, 6.5})";
				worksheet.Calculate();
				Assert.AreEqual(6.5d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void MaxaWithArrayWithIntegersReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "MAXA({5, 10, 4})";
				worksheet.Calculate();
				Assert.AreEqual(10d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void MaxaWithArrayWithMixedTypesReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "MAXA({TRUE, 0.55, 0.2, \"1.2\"})";
				worksheet.Calculate();
				Assert.AreEqual(0.55d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void MaxaWithIntegerInputReturnsCorrectValue()
		{
			var function = new Maxa();
			var result = function.Execute(FunctionsHelper.CreateArgs(15, 22, 13), this.ParsingContext);
			Assert.AreEqual(22d, result.Result);
		}

		[TestMethod]
		public void MaxaWithDoubleInputReturnsCorrectValue()
		{
			var function = new Maxa();
			var result = function.Execute(FunctionsHelper.CreateArgs(1.3, 6.2, 2.5), this.ParsingContext);
			Assert.AreEqual(6.2d, result.Result);
		}

		[TestMethod]
		public void MaxaWithFractionInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "MAXA((2/3), (9/8), (2/10))";
				worksheet.Calculate();
				Assert.AreEqual(1.125d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void MaxaWithStringsInputReturnsPoundValue()
		{
			var function = new Maxa();
			var result = function.Execute(FunctionsHelper.CreateArgs("string", "string"), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void MaxaWithDateObjectInputsReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "MAXA(DATE(2017, 6, 12), DATE(2017, 5, 10))";
				worksheet.Calculate();
				Assert.AreEqual(42898d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void MaxaWithDatesAsStringsReturnsCorrectValue()
		{
			var function = new Maxa();
			var result = function.Execute(FunctionsHelper.CreateArgs("5/5/2017", "6/8/2017"), this.ParsingContext);
			Assert.AreEqual(42894d, result.Result);
		}

		[TestMethod]
		public void MaxaWithMixedTypesReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "MAXA(TRUE, 5, 3, \"0.5\")";
				worksheet.Cells["B2"].Formula = "MAXA(TRUE, 5, \"string\")";
				worksheet.Calculate();
				Assert.AreEqual(5d, worksheet.Cells["B1"].Value);
				Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)worksheet.Cells["B2"].Value).Type);
			}
		}

		[TestMethod]
		public void MaxaShouldCalculateCorrectResult()
		{
			var func = new Maxa();
			var args = FunctionsHelper.CreateArgs(-1, 0, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(1d, result.Result);
		}

		[TestMethod]
		public void MaxaShouldCalculateCorrectResultUsingBool()
		{
			var func = new Maxa();
			var args = FunctionsHelper.CreateArgs(-1, 0, true);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(1d, result.Result);
		}

		[TestMethod]
		public void MaxaUsingStringShouldReturnPoundValue()
		{
			var func = new Maxa();
			var args = FunctionsHelper.CreateArgs(-1, "test");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void MaxaWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Maxa();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}
		#endregion
	}
}
