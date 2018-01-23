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
	public class RankEqTests : MathFunctionsTestBase
	{
		[TestMethod]
		public void RankEqWithNoInputsReturnsPoundValue()
		{
			var function = new Rank();
			var result = function.Execute(FunctionsHelper.CreateArgs(), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void RankEqWithTwoInputsReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 5;
				worksheet.Cells["B2"].Value = 6;
				worksheet.Cells["B3"].Value = 7;
				worksheet.Cells["B4"].Value = 42901;
				worksheet.Cells["B5"].Formula = "RANK.EQ(5, B1:B4)";
				worksheet.Calculate();
				Assert.AreEqual(4d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void RankEqWithGeneralStringFirstInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 4;
				worksheet.Cells["B2"].Value = 3;
				worksheet.Cells["B3"].Value = 2;
				worksheet.Cells["B4"].Formula = "RANK.EQ(\"string\", B1:B3)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)worksheet.Cells["B4"].Value).Type);
			}
		}

		[TestMethod]
		public void RankEqWithMissingFirstInputReturnsPoundNA()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 4;
				worksheet.Cells["B2"].Value = 3;
				worksheet.Cells["B3"].Value = 2;
				worksheet.Cells["B4"].Formula = "RANK.EQ( , B1:B3)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)worksheet.Cells["B4"].Value).Type);
			}
		}

		[TestMethod]
		public void RankEqWithNumericStringFirstInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 5;
				worksheet.Cells["B2"].Value = 6;
				worksheet.Cells["B3"].Value = 7;
				worksheet.Cells["B4"].Value = 42901;
				worksheet.Cells["B5"].Formula = "RANK.EQ(\"5\", B1:B4)";
				worksheet.Calculate();
				Assert.AreEqual(4d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void RankEqWithNumberNotInArrayReturnsPoundNA()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 2;
				worksheet.Cells["B2"].Value = 5;
				worksheet.Cells["B3"].Formula = "RANK.EQ(3, B1:B2)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)worksheet.Cells["B3"].Value).Type);
			}
		}

		[TestMethod]
		public void RankEqWithNumberAsDateFunctionReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 5;
				worksheet.Cells["B2"].Value = 6;
				worksheet.Cells["B3"].Value = 7;
				worksheet.Cells["B4"].Value = 42901;
				worksheet.Cells["B5"].Formula = "RANK.EQ(DATE(2017, 6, 15), B1:B4)";
				worksheet.Calculate();
				Assert.AreEqual(1d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void RankEqWithNumberAsDateAsStringReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 5;
				worksheet.Cells["B2"].Value = 6;
				worksheet.Cells["B3"].Value = 7;
				worksheet.Cells["B4"].Value = 42901;
				worksheet.Cells["B5"].Formula = "RANK.EQ(\"6/15/2017\", B1:B4)";
				worksheet.Calculate();
				Assert.AreEqual(1d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void RankEqWithBooleanNumberInputsReturnCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 0;
				worksheet.Cells["B2"].Value = 1;
				worksheet.Cells["B3"].Value = 7;
				worksheet.Cells["B10"].Formula = "RANK.EQ(TRUE, B1:B3)";
				worksheet.Cells["B11"].Formula = "RANK.EQ(FALSE, B1:B3)";
				worksheet.Calculate();
				Assert.AreEqual(2d, worksheet.Cells["B10"].Value);
				Assert.AreEqual(3d, worksheet.Cells["B11"].Value);
			}
		}

		[TestMethod]
		public void RankEqWithIntegerCellReferenceReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 4;
				worksheet.Cells["B2"].Value = 5;
				worksheet.Cells["B3"].Value = 2;
				worksheet.Cells["B4"].Value = 4;
				worksheet.Cells["B5"].Value = 6;
				worksheet.Cells["C1"].Formula = "RANK.EQ(5, B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(2d, worksheet.Cells["C1"].Value);
			}
		}

		[TestMethod]
		public void RankEqWithDoubleCellReferenceReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 4.5;
				worksheet.Cells["B2"].Value = 6.78;
				worksheet.Cells["B3"].Value = 3.14;
				worksheet.Cells["B4"].Value = 9.006;
				worksheet.Cells["B5"].Formula = "RANK.EQ(3.14, B1:B4)";
				worksheet.Calculate();
				Assert.AreEqual(4d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void RankEqWithBooleanCellReferenceReturnsPoundNA()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = true;
				worksheet.Cells["B2"].Value = false;
				worksheet.Cells["B3"].Value = true;
				worksheet.Cells["B4"].Formula = "RANK.EQ(TRUE, B1:B3)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)worksheet.Cells["B4"].Value).Type);
			}
		}

		[TestMethod]
		public void RankEqWithStringCellReferenceReturnsPoundNA()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "String";
				worksheet.Cells["B2"].Value = "String";
				worksheet.Cells["B3"].Value = "String";
				worksheet.Cells["B4"].Formula = "RANK.EQ(5, B1:B3)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)worksheet.Cells["B4"].Value).Type);
			}
		}

		[TestMethod]
		public void RankEqWithNumericStringCellReferenceReturnsPoundNA()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["A2"].Value = "8";
				worksheet.Cells["A3"].Value = "9";
				worksheet.Cells["A4"].Value = "5";
				worksheet.Cells["A5"].Value = "6";
				worksheet.Cells["B1"].Formula = "RANK.EQ(9, A2:A5)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)worksheet.Cells["B1"].Value).Type);
			}
		}

		[TestMethod]
		public void RankEqWithMixedTypesReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["A2"].Value = 6;
				worksheet.Cells["A3"].Value = false;
				worksheet.Cells["A4"].Value = 5.67;
				worksheet.Cells["A5"].Value = "cat";
				worksheet.Cells["A6"].Value = 3;
				worksheet.Cells["B1"].Formula = "RANK.EQ(3, A2:A6)";
				worksheet.Calculate();
				Assert.AreEqual(3d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void RankEqWithEmptyCellReferenceReturnsErrorValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "RANK.EQ(A2:A4)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)worksheet.Cells["B1"].Value).Type);
			}
		}

		[TestMethod]
		public void RankEqWithOrderAsZeroReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 4;
				worksheet.Cells["B2"].Value = 5;
				worksheet.Cells["B3"].Value = 2;
				worksheet.Cells["B4"].Formula = "RANK.EQ(5, B1:B3, 0)";
				worksheet.Calculate();
				Assert.AreEqual(1d, worksheet.Cells["B4"].Value);
			}
		}

		[TestMethod]
		public void RankEqWithNonZeroOrderReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 4;
				worksheet.Cells["B2"].Value = 5;
				worksheet.Cells["B3"].Value = 2;
				worksheet.Cells["B4"].Formula = "RANK.EQ(5, B1:B3, 450)";
				worksheet.Calculate();
				Assert.AreEqual(3d, worksheet.Cells["B4"].Value);
			}
		}

		[TestMethod]
		public void RankEqWithNumericStringReturnsPoundValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 4;
				worksheet.Cells["B2"].Value = 5;
				worksheet.Cells["B3"].Value = 2;
				worksheet.Cells["B4"].Formula = "RANK.EQ(5, B1:B3, \"3\")";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)worksheet.Cells["B4"].Value).Type);
			}
		}

		[TestMethod]
		public void RankEqWithGeneralStringReturnsPoundValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 4;
				worksheet.Cells["B2"].Value = 5;
				worksheet.Cells["B3"].Value = 2;
				worksheet.Cells["B4"].Formula = "RANK.EQ(5, B1:B3, \"string\")";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)worksheet.Cells["B4"].Value).Type);
			}
		}

		[TestMethod]
		public void RankEqWithOrderAsDateAsStringReturnsPoundValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 4;
				worksheet.Cells["B2"].Value = 5;
				worksheet.Cells["B3"].Value = 2;
				worksheet.Cells["B4"].Formula = "RANK.EQ(5, B1:B3, \"5/5/2017\")";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)worksheet.Cells["B4"].Value).Type);
			}
		}

		[TestMethod]
		public void RankEqWithBooleanOrderInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 4;
				worksheet.Cells["B2"].Value = 5;
				worksheet.Cells["B3"].Value = 2;
				worksheet.Cells["B4"].Formula = "RANK.EQ(5, B1:B3, TRUE)";
				worksheet.Cells["B5"].Formula = "RANK.EQ(5, B1:B3, FALSE)";
				worksheet.Calculate();
				Assert.AreEqual(3d, worksheet.Cells["B4"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void RankEqWithDateObjectCellReferenceReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "DATE(2017, 5, 5)";
				worksheet.Cells["B2"].Formula = "DATE(2017, 6, 15)";
				worksheet.Cells["B3"].Formula = "RANK.EQ(42901, B1:B2)";
				worksheet.Calculate();
				Assert.AreEqual(1d, worksheet.Cells["B3"].Value);
			}
		}

		[TestMethod]
		public void RankEqWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Rank();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void RankEqWithDuplicateValuesReturnsCorrectResult()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 5;
				worksheet.Cells["B3"].Value = 5;
				worksheet.Cells["B4"].Value = 7;
				worksheet.Cells["B5"].Formula = "RANK.EQ(5, B1:B4)";
				worksheet.Calculate();
				Assert.AreEqual(2d, worksheet.Cells["B5"].Value);

			}
		}
	}
}
