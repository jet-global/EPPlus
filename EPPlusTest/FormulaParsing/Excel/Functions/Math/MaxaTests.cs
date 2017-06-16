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
		public void MaxaWithNoInputsReturnsPoundValue()
		{
			var function = new Maxa();
			var result = function.Execute(FunctionsHelper.CreateArgs(), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void MaxaWithReferenceToEmptyCellsReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "MAXA(A2:A4)";
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
				worksheet.Cells["B1"].Value = 0.6;
				worksheet.Cells["B2"].Value = 0.3;
				worksheet.Cells["B3"].Value = "TRUE";
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
				worksheet.Cells["B1"].Value = 5;
				worksheet.Cells["B2"].Value = 3;
				worksheet.Cells["B3"].Value = "67";
				worksheet.Cells["B4"].Formula = "MAXA(B1:B3)";
				worksheet.Calculate();
				Assert.AreEqual(5d, worksheet.Cells["B4"].Value);
			}
		}

		[TestMethod]
		public void MaxaWithArrayWithLogicalValueReturnsCorrectResult()
		{
			var function = new Maxa();
			var result = function.Execute(FunctionsHelper.CreateArgs(0.6, 0.1, true), this.ParsingContext);
			Assert.AreEqual(0.6d, result.Result);

			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "MAXA({0.6, TRUE, 0.1})";
				worksheet.Calculate();
				Assert.AreEqual(0.6d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void MaxaWithArrayWithNumericStringReturnsCorrectValue()
		{
			var function = new Maxa();
			var result = function.Execute(FunctionsHelper.CreateArgs("90", 79, 4), this.ParsingContext);
			Assert.AreEqual(79d, result.Result);

			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "MAXA({\"90\", 79, 4})";
				worksheet.Calculate();
				Assert.AreEqual(79d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void MaxaWithReferenceToCellsWithStringsReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "String";
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
				worksheet.Cells["B1"].Formula = "DATE(2017, 5 ,12)";
				worksheet.Cells["B2"].Formula = "DATE(2017, 6, 2)";
				worksheet.Cells["B3"].Formula = "DATE(2017, 5, 15)";
				worksheet.Cells["B4"].Formula = "MAXA(B1:B3)";
				worksheet.Calculate();
				Assert.AreEqual(42888d, worksheet.Cells["B4"].Value);
			}
		}

		[TestMethod]
		public void MaxaWithReferenceToCellsWithDateObjectsAsOADatesReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 42867;
				worksheet.Cells["B2"].Value = 42888;
				worksheet.Cells["B3"].Value = 42870;
				worksheet.Cells["B4"].Formula = "MAXA(B1:B3)";
				worksheet.Calculate();
				Assert.AreEqual(42888d, worksheet.Cells["B4"].Value);
			}
		}

		[TestMethod]
		public void MaxaWithReferenceToCellsWithDatesAsStringsReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "5/5/2017";
				worksheet.Cells["B2"].Value = "6/2/2017";
				worksheet.Cells["B3"].Value = "5/15/2017";
				worksheet.Cells["B4"].Formula = "MINA(B1:B3)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B4"].Value);
			}
		}

		[TestMethod]
		public void MaxaWithArrayWithStringsReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var workbook = package.Workbook.Worksheets.Add("Sheet1");
				workbook.Cells["B1"].Formula = "MAXA({\"string\", \"string\"})";
				workbook.Calculate();
				Assert.AreEqual(0d, workbook.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void MaxaWithArrayWithDatesAsStringsReturnsZero()
		{
			var function = new Maxa();
			var result = function.Execute(FunctionsHelper.CreateArgs("5/5/2017", "6/2/2017"), this.ParsingContext);
			Assert.AreEqual(0d, result.Result);

			using (var package = new ExcelPackage())
			{
				var workbook = package.Workbook.Worksheets.Add("Sheet1");
				workbook.Cells["B1"].Formula = "MAXA(\"5/5/2017\", \"6/2/2017\")";
				workbook.Calculate();
				Assert.AreEqual(0d, workbook.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void MaxaWithIntegerInputReturnsCorrectValue()
		{
			var function = new Maxa();
			var result = function.Execute(FunctionsHelper.CreateArgs(5, 16, 20, 17), this.ParsingContext);
			Assert.AreEqual(20d, result.Result);
		}

		[TestMethod]
		public void MaxaWithDoublesInputReturnsCorrectValue()
		{
			var function = new Maxa();
			var result = function.Execute(FunctionsHelper.CreateArgs(123.3, 12.6, 9.22), this.ParsingContext);
			Assert.AreEqual(123.3d, result.Result);
		}

		[TestMethod]
		public void MaxaWithFractionsInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var workbook = package.Workbook.Worksheets.Add("Sheet1");
				workbook.Cells["B1"].Formula = "MAXA((2/3),(9/8),(2/55))";
				workbook.Calculate();
				Assert.AreEqual(1.125d, workbook.Cells["B1"].Value);
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
				worksheet.Cells["B1"].Formula = "MAXA(DATE(2017, 6, 15), DATE(2017, 5, 18))";
				worksheet.Calculate();
				Assert.AreEqual(42901d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void MaxaWithDatesAsStringsInputsReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "MAXA(\"5/2/2017\", \"6/25/2017\")";
				worksheet.Calculate();
				Assert.AreEqual(42911d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void MaxaWithMixedTypesReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "MAXA(TRUE, 5, , 8, \"16\")";
				worksheet.Calculate();
				Assert.AreEqual(16d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void MaxaWithLogicValueInListOfArgumentsReturnsCorrectValue()
		{
			var function = new Maxa();
			var result = function.Execute(FunctionsHelper.CreateArgs(0.5, "TRUE", 0.6), this.ParsingContext);
			Assert.AreEqual(1d, result.Result);

			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "MAXA(TRUE, 0.5, 0.6)";
				worksheet.Calculate();
				Assert.AreEqual(1d, worksheet.Cells["B1"].Value);
			}
		}
		#endregion
	}
}
