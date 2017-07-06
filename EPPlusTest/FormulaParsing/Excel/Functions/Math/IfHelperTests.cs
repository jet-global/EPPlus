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
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Numeric;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using static OfficeOpenXml.FormulaParsing.ExcelDataProvider;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math
{
	[TestClass]
	public class IfHelperTests : MathFunctionsTestBase 
	{
		#region IfHelper Function (Execute) Tests
		[TestMethod]
		public void CalculateCriteriaWithSameRowCellReferenceReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet2");
				var provider = new EpplusExcelDataProvider(package);
				this.ParsingContext.Scopes.NewScope(RangeAddress.Empty);
				worksheet.Cells["B1"].Value = 5;
				worksheet.Cells["B2"].Value = 10;
				worksheet.Cells["B3"].Value = 15;
				IRangeInfo testRange = provider.GetRange(worksheet.Name, 1, 2, 3, 2);
				IRangeInfo firstRange = provider.GetRange(worksheet.Name, 2, 2, 2, 2);
				var address = firstRange.Address;
				var result = IfHelper.CalculateCriteria(FunctionsHelper.CreateArgs(firstRange, testRange), worksheet, address._fromRow, address._fromCol);
				Assert.AreEqual(10, result);
			}
		}

		[TestMethod]
		public void CalculateCriteriaWithSameColumnCellReferenceReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				var provider = new EpplusExcelDataProvider(package);
				this.ParsingContext.Scopes.NewScope(RangeAddress.Empty);
				worksheet.Cells["E7"].Value = 5;
				worksheet.Cells["F7"].Value = 10;
				worksheet.Cells["G7"].Value = 15;
				IRangeInfo testRange = provider.GetRange(worksheet.Name, 7, 5, 7, 7);
				IRangeInfo firstRange = provider.GetRange(worksheet.Name, 6, 6, 6, 6);
				var address = firstRange.Address;
				var result = IfHelper.CalculateCriteria(FunctionsHelper.CreateArgs(firstRange, testRange), worksheet, address._fromRow, address._fromCol);
				Assert.AreEqual(10, result);
			}
		}

		[TestMethod]
		public void CalculateCriteriaWithNonMatchingRowReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet2");
				var provider = new EpplusExcelDataProvider(package);
				this.ParsingContext.Scopes.NewScope(RangeAddress.Empty);
				worksheet.Cells["B1"].Value = 5;
				worksheet.Cells["B2"].Value = 10;
				worksheet.Cells["B3"].Value = 15;
				IRangeInfo testRange = provider.GetRange(worksheet.Name, 1, 2, 3, 2);
				IRangeInfo firstRange = provider.GetRange(worksheet.Name, 5, 5, 5, 5);
				var address = firstRange.Address;
				var result = IfHelper.CalculateCriteria(FunctionsHelper.CreateArgs(firstRange, testRange), worksheet, address._fromRow, address._fromCol);
				Assert.AreEqual(0, result);
			}
		}

		[TestMethod]
		public void CalculateCriteriaWithNonMatchingColReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				var provider = new EpplusExcelDataProvider(package);
				this.ParsingContext.Scopes.NewScope(RangeAddress.Empty);
				worksheet.Cells["E7"].Value = 5;
				worksheet.Cells["F7"].Value = 10;
				worksheet.Cells["G7"].Value = 15;
				IRangeInfo testRange = provider.GetRange(worksheet.Name, 7, 5, 7, 7);
				IRangeInfo firstRange = provider.GetRange(worksheet.Name, 8, 8, 8, 8);
				var address = firstRange.Address;
				var result = IfHelper.CalculateCriteria(FunctionsHelper.CreateArgs(firstRange, testRange), worksheet, address._fromRow, address._fromCol);
				Assert.AreEqual(0, result);
			}
		}

		[TestMethod]
		public void CalculateCriteriaWithObjectReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				var provider = new EpplusExcelDataProvider(package);
				this.ParsingContext.Scopes.NewScope(RangeAddress.Empty);
				worksheet.Cells["E12"].Value = 1;
				worksheet.Cells["E13"].Value = 2;
				worksheet.Cells["E14"].Value = 3;
				worksheet.Cells["F12"].Value = 1;
				worksheet.Cells["F13"].Value = ">2";
				worksheet.Cells["F14"].Value = 3;
				worksheet.Cells["H13"].Formula = "SUMIF(E12:E14, F12:F14)";
				worksheet.Calculate();
				Assert.AreEqual(3d, worksheet.Cells["H13"].Value);
			}
		}
		#endregion
	}
}
