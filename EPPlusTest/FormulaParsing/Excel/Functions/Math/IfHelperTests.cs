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
using System.Globalization;
using System.Threading;
using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using static OfficeOpenXml.FormulaParsing.ExcelDataProvider;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math
{
	[TestClass]
	public class IfHelperTests : MathFunctionsTestBase 
	{
		#region IfHelper Function (Execute) Tests
		[TestMethod]
		public void ObjectMatchesCriteriaHandlesErrorValuesAsCriteria()
		{
			Assert.Fail("This test will fail until the IfHelper class has been fixed to properly parse error values in criteria for all cultures.");
			var currentCulture = CultureInfo.CurrentCulture;
			try
			{
				var us = CultureInfo.CreateSpecificCulture("en-US");
				Thread.CurrentThread.CurrentCulture = us;
				using (var package = new ExcelPackage())
				{
					var worksheet = package.Workbook.Worksheets.Add("Sheet1");
					worksheet.Cells["C2"].Formula = "DAY(\"word\")"; // Evaluates to #VALUE!.
					worksheet.Cells["C3"].Value = "#VALUE!";
					worksheet.Cells["C4"].Value = "#VALUE";
					worksheet.Cells["B2"].Formula = "COUNTIF(C2,\"=#VALUE!\")";
					worksheet.Cells["B3"].Formula = "COUNTIF(C3,\"=#VALUE!\")";
					worksheet.Cells["B4"].Formula = "COUNTIF(C2,\"#VALUE!\")";
					worksheet.Cells["B5"].Formula = "COUNTIF(C3,\"#VALUE!\")";
					worksheet.Cells["B6"].Formula = "COUNTIF(C2,\"=#VALUE\")";
					worksheet.Cells["B7"].Formula = "COUNTIF(C3,\"=#VALUE\")";
					worksheet.Cells["B8"].Formula = "COUNTIF(C4,\"=#VALUE\")";
					worksheet.Calculate();
					Assert.AreEqual(1, worksheet.Cells["B2"].Value);
					Assert.AreEqual(0, worksheet.Cells["B3"].Value);
					Assert.AreEqual(1, worksheet.Cells["B4"].Value);
					Assert.AreEqual(0, worksheet.Cells["B5"].Value);
					Assert.AreEqual(0, worksheet.Cells["B6"].Value);
					Assert.AreEqual(0, worksheet.Cells["B7"].Value);
					Assert.AreEqual(1, worksheet.Cells["B8"].Value);
				}
				var de = CultureInfo.CreateSpecificCulture("de-DE");
				Thread.CurrentThread.CurrentCulture = de;
				using (var package = new ExcelPackage())
				{
					var worksheet = package.Workbook.Worksheets.Add("Sheet1");
					worksheet.Cells["C2"].Formula = "DAY(\"word\")"; // Evaluates to #VALUE!.
					worksheet.Cells["C3"].Value = "#WERT!"; // Note that #WERT! is the German translation for #VALUE!.
					worksheet.Cells["C4"].Value = "#WERT";
					worksheet.Cells["B2"].Formula = "COUNTIF(C2,\"=#WERT!\")";
					worksheet.Cells["B3"].Formula = "COUNTIF(C3,\"=#WERT!\")";
					worksheet.Cells["B4"].Formula = "COUNTIF(C2,\"#WERT!\")";
					worksheet.Cells["B5"].Formula = "COUNTIF(C3,\"#WERT!\")";
					worksheet.Cells["B6"].Formula = "COUNTIF(C2,\"=#WERT\")";
					worksheet.Cells["B7"].Formula = "COUNTIF(C3,\"=#WERT\")";
					worksheet.Cells["B8"].Formula = "COUNTIF(C4,\"=#WERT\")";
					worksheet.Calculate();
					Assert.AreEqual(1d, worksheet.Cells["B2"].Value);
					Assert.AreEqual(0d, worksheet.Cells["B3"].Value);
					Assert.AreEqual(1d, worksheet.Cells["B4"].Value);
					Assert.AreEqual(0d, worksheet.Cells["B5"].Value);
					Assert.AreEqual(0d, worksheet.Cells["B6"].Value);
					Assert.AreEqual(0d, worksheet.Cells["B7"].Value);
					Assert.AreEqual(1d, worksheet.Cells["B8"].Value);
				}
			}
			finally
			{
				Thread.CurrentThread.CurrentCulture = currentCulture;
			}
		}

		[TestMethod]
		public void ObjectMatchesCriteriaHandlesBooleanValuesAsCriteria()
		{
			Assert.Fail("This test will fail until the IfHelper class has been fixed to properly parse booleans in criteria for all cultures.");
			var currentCulture = CultureInfo.CurrentCulture;
			try
			{
				var us = CultureInfo.CreateSpecificCulture("en-US");
				Thread.CurrentThread.CurrentCulture = us;
				using (var package = new ExcelPackage())
				{
					var worksheet = package.Workbook.Worksheets.Add("Sheet1");
					worksheet.Cells["C2"].Value = true;
					worksheet.Cells["C3"].Value = "TRUE";
					worksheet.Cells["B2"].Formula = "COUNTIF(C2, TRUE)";
					worksheet.Cells["B3"].Formula = "COUNTIF(C3, TRUE)";
					worksheet.Cells["B4"].Formula = "COUNTIF(C2,\"TRUE\")";
					worksheet.Cells["B5"].Formula = "COUNTIF(C3,\"TRUE\")";
					worksheet.Cells["B6"].Formula = "COUNTIF(C2,\"=TRUE\")";
					worksheet.Cells["B7"].Formula = "COUNTIF(C3,\"=TRUE\")";
					worksheet.Calculate();
					Assert.AreEqual(1d, worksheet.Cells["B2"].Value);
					Assert.AreEqual(0d, worksheet.Cells["B3"].Value);
					Assert.AreEqual(1d, worksheet.Cells["B4"].Value);
					Assert.AreEqual(0d, worksheet.Cells["B5"].Value);
					Assert.AreEqual(1d, worksheet.Cells["B6"].Value);
					Assert.AreEqual(0d, worksheet.Cells["B7"].Value);
				}
			}
			finally
			{
				Thread.CurrentThread.CurrentCulture = currentCulture;
			}
		}

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
