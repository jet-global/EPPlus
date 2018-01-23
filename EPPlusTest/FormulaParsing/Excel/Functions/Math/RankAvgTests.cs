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
	public class RankAvgTests : MathFunctionsTestBase
	{
		[TestMethod]
		public void RankAvgWithNoInputsReturnsPoundValue()
		{
			var function = new Rank(true);
			var result = function.Execute(FunctionsHelper.CreateArgs(), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void RankAvgWithTwoInputsReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 5;
				worksheet.Cells["B3"].Value = 5;
				worksheet.Cells["B4"].Value = 7;
				worksheet.Cells["B5"].Formula = "RANK.AVG(5, B1:B4)";
				worksheet.Calculate();
				Assert.AreEqual(2.5d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void RankAvgWithNumberAsGeneralStringReturnsPoundValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 4;
				worksheet.Cells["B2"].Value = 3;
				worksheet.Cells["B3"].Formula = "RANK.AVG(\"string\", B1:B2)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)worksheet.Cells["B3"].Value).Type);
			}
		}

		[TestMethod]
		public void RankAvgWithNoNumberInputRetunrnsPoundNA()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 3;
				worksheet.Cells["B2"].Value = 3;
				worksheet.Cells["B3"].Formula = "RANK.AVG(, B1:B2)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)worksheet.Cells["B3"].Value).Type);
			}
		}

		[TestMethod]
		public void RankAvgWithNumberAsNumericStringReturnsCorrectValue()
		{
			using (var pacakge = new ExcelPackage())
			{
				var worksheet = pacakge.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 5;
				worksheet.Cells["B3"].Value = 5;
				worksheet.Cells["B4"].Value = 7;
				worksheet.Cells["B5"].Formula = "RANK.AVG(\"5\", B1:B4)";
				worksheet.Calculate();
				Assert.AreEqual(2.5d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void RankAvgWithNumberNotInArrayReturnsPoundNA()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Formula = "RANK.AVG(4, B1:B2)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)worksheet.Cells["B3"].Value).Type);
			}
		}

		[TestMethod]
		public void RankAvgWithDateFunctionNumberInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 5;
				worksheet.Cells["B2"].Value = 7;
				worksheet.Cells["B3"].Value = 42901;
				worksheet.Cells["B4"].Formula = "RANK.AVG(DATE(2017, 6, 15), B1:B3)";
				worksheet.Calculate();
				Assert.AreEqual(1d, worksheet.Cells["B4"].Value);
			}
		}

		[TestMethod]
		public void RankAvgWithNumberInputAsDateAsStringReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 5;
				worksheet.Cells["B2"].Value = 7;
				worksheet.Cells["B3"].Value = 42901;
				worksheet.Cells["B4"].Formula = "RANK.AVG(\"6/15/2017\", B1:B3)";
				worksheet.Calculate();
				Assert.AreEqual(1d, worksheet.Cells["B4"].Value);
			}
		}

		[TestMethod]
		public void RankAvgWithBooleanNumberInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 0;
				worksheet.Cells["B2"].Value = 1;
				worksheet.Cells["B3"].Formula = "RANK.AVG(TRUE, B1:B2)";
				worksheet.Cells["B4"].Formula = "RANK.AVG(FALSE, B1:B2)";
				worksheet.Calculate();
				Assert.AreEqual(1d, worksheet.Cells["B3"].Value);
				Assert.AreEqual(2d, worksheet.Cells["B4"].Value);
			}
		}

		[TestMethod]
		public void RankAvgWithIntegerCellReferenceReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 4;
				worksheet.Cells["B2"].Value = 5;
				worksheet.Cells["B3"].Value = 2;
				worksheet.Cells["B4"].Value = 2;
				worksheet.Cells["B5"].Formula = "RANK.AVG(5, B1:B4)";
				worksheet.Calculate();
				Assert.AreEqual(1d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void RankAvgWithDoubleCellReferenceReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 4.5;
				worksheet.Cells["B2"].Value = 6.78;
				worksheet.Cells["B3"].Value = 3.14;
				worksheet.Cells["B4"].Value = 3.14;
				worksheet.Cells["B5"].Formula = "RANK.AVG(3.14, B1:B4)";
				worksheet.Calculate();
				Assert.AreEqual(3.5d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void RankAvgWithBooleanCellReferenceReturnsPoundNA()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = true;
				worksheet.Cells["B2"].Value = false;
				worksheet.Cells["B3"].Value = true;
				worksheet.Cells["B4"].Formula = "RANK(TRUE, B1:B3)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)worksheet.Cells["B4"].Value).Type);
			}
		}

		[TestMethod]
		public void RankAvgWithGeneralStringCellReferenceReturnsPoundNA()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "string";
				worksheet.Cells["B2"].Value = "string";
				worksheet.Cells["B3"].Value = "string";
				worksheet.Cells["B4"].Formula = "RANK.AVG(5, B1:B3)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)worksheet.Cells["B4"].Value).Type);
			}
		}

		[TestMethod]
		public void RankAvgWithNumericStringCellReferenceReturnsPoundNA()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "9";
				worksheet.Cells["B2"].Value = "8";
				worksheet.Cells["B3"].Value = "6";
				worksheet.Cells["B4"].Formula = "RANK.AVG(9, B1:B3)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)worksheet.Cells["B4"].Value).Type);
			}
		}

		[TestMethod]
		public void RankAvgWithMixedTypesCellReferenceReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = false;
				worksheet.Cells["B2"].Value = 5.67;
				worksheet.Cells["B3"].Value = "Cat";
				worksheet.Cells["B4"].Value = 3;
				worksheet.Cells["B5"].Formula = "RANK.AVG(3, B1:B4)";
				worksheet.Calculate();
				Assert.AreEqual(2d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void RankAvgWithEmptyCellReferenceReturnsPoundNA()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "RANK.AVG(4, A2:A4)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)worksheet.Cells["B1"].Value).Type);
			}
		}

		[TestMethod]
		public void RankAvgWithZeroOrderInputReturnsCorrectValue()
		{
			using (var packge = new ExcelPackage())
			{
				var worksheet = packge.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 4;
				worksheet.Cells["B2"].Value = 5;
				worksheet.Cells["B3"].Value = 2;
				worksheet.Cells["B4"].Value = 2;
				worksheet.Cells["B5"].Formula = "RANK.AVG(5, B1:B4, 0)";
				worksheet.Calculate();
				Assert.AreEqual(1d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void RankAvgWithNonZeroOrderInputReturnsCorrectValue()
		{
			using (var packge = new ExcelPackage())
			{
				var worksheet = packge.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 4;
				worksheet.Cells["B2"].Value = 5;
				worksheet.Cells["B3"].Value = 2;
				worksheet.Cells["B4"].Value = 2;
				worksheet.Cells["B5"].Formula = "RANK.AVG(5, B1:B4, 450)";
				worksheet.Calculate();
				Assert.AreEqual(4d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void RankAvgWithNumericStringOrderReturnsPoundValue()
		{
			using (var packge = new ExcelPackage())
			{
				var worksheet = packge.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 4;
				worksheet.Cells["B2"].Value = 5;
				worksheet.Cells["B3"].Value = 2;
				worksheet.Cells["B4"].Value = 2;
				worksheet.Cells["B5"].Formula = "RANK.AVG(5, B1:B4, \"4\")";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)worksheet.Cells["B5"].Value).Type);
			}
		}

		[TestMethod]
		public void RankAvgWithGeneralStringOrderReturnsPoundValue()
		{
			using (var packge = new ExcelPackage())
			{
				var worksheet = packge.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 4;
				worksheet.Cells["B2"].Value = 5;
				worksheet.Cells["B3"].Value = 2;
				worksheet.Cells["B4"].Value = 2;
				worksheet.Cells["B5"].Formula = "RANK.AVG(5, B1:B4, \"string\")";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)worksheet.Cells["B5"].Value).Type);
			}
		}

		[TestMethod]
		public void RankAvgWithDateAsStringOrderReturnsPoundValue()
		{
			using (var packge = new ExcelPackage())
			{
				var worksheet = packge.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 4;
				worksheet.Cells["B2"].Value = 5;
				worksheet.Cells["B3"].Value = 2;
				worksheet.Cells["B4"].Value = 2;
				worksheet.Cells["B5"].Formula = "RANK.AVG(5, B1:B4, \"5/5/2017\")";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)worksheet.Cells["B5"].Value).Type);
			}
		}

		[TestMethod]
		public void RankAvgWithBooleanOrderReturnsCorrectValues()
		{
			using (var packge = new ExcelPackage())
			{
				var worksheet = packge.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 4;
				worksheet.Cells["B2"].Value = 5;
				worksheet.Cells["B3"].Value = 2;
				worksheet.Cells["B4"].Value = 2;
				worksheet.Cells["B5"].Formula = "RANK.AVG(5, B1:B4, TRUE)";
				worksheet.Cells["B6"].Formula = "RANK.AVG(5, B1:B4, FALSE)";
				worksheet.Calculate();
				Assert.AreEqual(4d, worksheet.Cells["B5"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B6"].Value);
			}
		}
	}
}
