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
	public class MinaTests : MathFunctionsTestBase
	{
		#region Mina Function (Execute) Tests
		[TestMethod]
		public void MinaWithNoInputsReturnsPoundValue()
		{
			var function = new Mina();
			var result = function.Execute(FunctionsHelper.CreateArgs(), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void MinaWithMaxInputReturnsCorrectValue()
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
				worksheet.Cells["A10"].Value = 2;
				worksheet.Cells["C1"].Formula = "MINA(A1:A255)";
				worksheet.Cells["C2"].Formula = "MINA(A1:A259)";
				worksheet.Calculate();
				Assert.AreEqual(2d, worksheet.Cells["C1"].Value);
				Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)worksheet.Cells["C2"].Value).Type);
			}
		}

		[TestMethod]
		public void MinaWithReferenceToAnEmptyCellReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "MINA(A3:A6)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void MinaWithReferenceToCellWithLogicalValueReturnsCorrectResult()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = true;
				worksheet.Cells["B2"].Value = 0.5;
				worksheet.Cells["C1"].Formula = "MINA(B1:B2)";
				worksheet.Calculate();
				Assert.AreEqual(0.5d, worksheet.Cells["C1"].Value);
			}
		}

		[TestMethod]
		public void MinaWithReferenceToCellWithNumericStringReturnsCorrectResult()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 2;
				worksheet.Cells["B2"].Value = 5;
				worksheet.Cells["B3"].Value = "1";
				worksheet.Cells["C1"].Formula = "MINA(B1:B3)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["C1"].Value);
			}
		}

		[TestMethod]
		public void MinaWithReferenceToCellsWithStringsReturnsCorrectResult()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "String";
				worksheet.Cells["B2"].Value = "string";
				worksheet.Cells["C1"].Formula = "MINA(B1:B2)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["C1"].Value);
			}
		}

		[TestMethod]
		public void MinaWithReferenceToCellsWithDateObjectsReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "DATE(2017, 6, 2)";
				worksheet.Cells["B2"].Formula = "DATE(2017, 8, 15)";
				worksheet.Cells["C1"].Formula = "MINA(B1:B2)";
				worksheet.Calculate();
				Assert.AreEqual(42888d, worksheet.Cells["C1"].Value);
			}
		}

		[TestMethod]
		public void MinaWithReferenceToCellsWithDatesAsStringsReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "6/2/2017";
				worksheet.Cells["B2"].Value = "8/15/2017";
				worksheet.Cells["C1"].Formula = "MINA(B1:B2)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["C1"].Value);
			}
		}

		[TestMethod]
		public void MinaWithReferenceToMixedTypesReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = true;
				worksheet.Cells["B2"].Value = 5;
				worksheet.Cells["B3"].Value = "string";
				worksheet.Cells["C1"].Formula = "MINA(B1:B2)";
				worksheet.Cells["C2"].Formula = "MINA(B1:B3)";
				worksheet.Calculate();
				Assert.AreEqual(1d, worksheet.Cells["C1"].Value);
				Assert.AreEqual(0d, worksheet.Cells["C2"].Value);
			}
		}

		[TestMethod]
		public void MinaWithReferenceToDoublesReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 1.5;
				worksheet.Cells["B2"].Value = 6.3;
				worksheet.Cells["B3"].Value = 0.26;
				worksheet.Cells["C1"].Formula = "MINA(B1:B3)";
				worksheet.Calculate();
				Assert.AreEqual(0.26d, worksheet.Cells["C1"].Value);
			}
		}

		[TestMethod]
		public void MinaWithReferenceToIntegersReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 5;
				worksheet.Cells["B2"].Value = 8;
				worksheet.Cells["B3"].Value = 2;
				worksheet.Cells["C1"].Formula = "MINA(B1:B3)";
				worksheet.Calculate();
				Assert.AreEqual(2d, worksheet.Cells["C1"].Value);
			}
		}

		[TestMethod]
		public void MinaWithArrayWithLogicalValueReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["C1"].Formula = "MINA({5, 2, TRUE})";
				worksheet.Calculate();
				Assert.AreEqual(2d, worksheet.Cells["C1"].Value);
			}
		}

		[TestMethod]
		public void MinaWithArrayWithNumericStringReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["C1"].Formula = "MINA({5, 2, \"1\"})";
				worksheet.Calculate();
				Assert.AreEqual(2d, worksheet.Cells["C1"].Value);
			}
		}

		[TestMethod]
		public void MinaWithArrayWithStringsReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["C1"].Formula = "MINA({3, 6, \"string\"})";
				worksheet.Calculate();
				Assert.AreEqual(3d, worksheet.Cells["C1"].Value);
			}
		}

		[TestMethod]
		public void MinaWithArrayWithDatesAsStringsReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["C1"].Formula = "MINA({\"6/2/2017\", \"5/15/2017\"})";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["C1"].Value);
			}
		}

		[TestMethod]
		public void MinaWithArrayWithMixedTypesReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["C1"].Formula = "MINA({TRUE, 5, 2})";
				worksheet.Cells["C2"].Formula = "MINA({TRUE, 5, \"2\"})";
				worksheet.Cells["C3"].Formula = "MINA({TRUE, 5, 2, \"string\"})";
				worksheet.Calculate();
				Assert.AreEqual(2d, worksheet.Cells["C1"].Value);
				Assert.AreEqual(5d, worksheet.Cells["C2"].Value);
				Assert.AreEqual(2d, worksheet.Cells["C3"].Value);
			}
		}

		[TestMethod]
		public void MinaWithArrayWithDoublesReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["C1"].Formula = "MINA({1.5, 0.3, 6.5})";
				worksheet.Calculate();
				Assert.AreEqual(0.3d, worksheet.Cells["C1"].Value);
			}
		}

		[TestMethod]
		public void MinaWithArrayWithIntegersReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["C1"].Formula = "MINA({5, 10, 4})";
				worksheet.Calculate();
				Assert.AreEqual(4d, worksheet.Cells["C1"].Value);
			}
		}

		[TestMethod]
		public void MinaWithIntegerInputReturnsCorrectValue()
		{
			var function = new Mina();
			var result = function.Execute(FunctionsHelper.CreateArgs(5, 6, 12, 88), this.ParsingContext);
			Assert.AreEqual(5d, result.Result);
		}

		[TestMethod]
		public void MinaWithDoublesInputReturnsCorrectValue()
		{
			var function = new Mina();
			var result = function.Execute(FunctionsHelper.CreateArgs(1.2, 2.2, 3.3), this.ParsingContext);
			Assert.AreEqual(1.2d, result.Result);
		}

		[TestMethod]
		public void MinaWithFractionsInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["C1"].Formula = "MINA((2/3),(9/8),(2/10))";
				worksheet.Calculate();
				Assert.AreEqual(0.2d, worksheet.Cells["C1"].Value);
			}
		}

		[TestMethod]
		public void MinaWithStringInputReturnsCorrectValue()
		{
			var function = new Mina();
			var result = function.Execute(FunctionsHelper.CreateArgs("string", "string"), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void MinaWithDateObjectsInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["C1"].Formula = "MINA(DATE(2017, 6, 12), DATE(2017, 5, 10))";
				worksheet.Calculate();
				Assert.AreEqual(42865d, worksheet.Cells["C1"].Value);
			}
		}

		[TestMethod]
		public void MinaWithDatesAsStringsInputReturnsCorrectValue()
		{
			var function = new Mina();
			var result = function.Execute(FunctionsHelper.CreateArgs("5/5/2017", "6/8/2017"), this.ParsingContext);
			Assert.AreEqual(42860d, result.Result);
		}

		[TestMethod]
		public void MinaWithMixedTypesReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["C1"].Formula = "MINA(TRUE, 3, 5)";
				worksheet.Cells["C2"].Formula = "MINA(TRUE, 3, 5, \"0.5\")";
				worksheet.Cells["C3"].Formula = "MINA(TRUE, 5, \"string\")";
				worksheet.Calculate();
				Assert.AreEqual(1d, worksheet.Cells["C1"].Value);
				Assert.AreEqual(0.5d, worksheet.Cells["C2"].Value);
				Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)worksheet.Cells["C3"].Value).Type);
			}
		}

		//This is an already existing EPPlus Test.
		[TestMethod]
		public void MinaWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Mina();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}
		#endregion
	}
}
