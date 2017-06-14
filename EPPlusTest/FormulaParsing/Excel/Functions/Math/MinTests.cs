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
				worksheet.Cells["B1"].Formula = "MIN(5, 6, TRUE, 2)";
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
				worksheet.Cells["B1"].Formula = "MIN(5, 6, \"2\")";
				worksheet.Calculate();
				Assert.AreEqual(5d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void MinWithMaxArgumentsReturnsCorrectValue()
		{
			// This functionality is different from that of Excel's. Normally when too many arguments are entered
			// into a function it won't let you calculate the function, however in EPPlus it will return a pound
			// NA error instead. 
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
				worksheet.Cells["C1"].Formula = "MIN(A1:A255)";
				worksheet.Cells["C2"].Formula = "MIN(A1:A270)";
				worksheet.Calculate();
				Assert.AreEqual(4d, worksheet.Cells["C1"].Value);
				Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)worksheet.Cells["C2"].Value).Type);
			}
		}

		[TestMethod]
		public void  MinWithReferenceToCellsWithStringsReturnsZero()
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
	}
}
