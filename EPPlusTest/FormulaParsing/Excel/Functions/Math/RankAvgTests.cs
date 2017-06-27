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
	}
}
