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
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math
{
	[TestClass]
	public class CountBlankTests
	{
		#region CountBlank Tests
		[TestMethod]
		public void CountBlankWithRangeOf0()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B2"].Formula = "COUNTBLANK(C2)";
				worksheet.Cells["C2"].Value = 0;
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void CountBlankWithEmptyString()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B2"].Formula = "COUNTBLANK(C2)";
				worksheet.Cells["C2"].Value = string.Empty;
				worksheet.Calculate();
				Assert.AreEqual(1d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void CountBlankWithEmptyCell()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B2"].Formula = "COUNTBLANK(C2)";
				worksheet.Cells["C2"].Value = null;
				worksheet.Calculate();
				Assert.AreEqual(1d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void CountBlankWithArrayWithEmptyStringFirst()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B2"].Formula = "COUNTBLANK(C2)";
				worksheet.Cells["C2"].Formula = "{\"\",\"word\"}";
				worksheet.Calculate();
				Assert.AreEqual(1d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void CountBlankWithArrayWithEmptyStringLast()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B2"].Formula = "COUNTBLANK(C2)";
				worksheet.Cells["C2"].Formula = "{\"word\",\"\"}";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void CountBlankWithNumber()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B2"].Formula = "COUNTBLANK(C2)";
				worksheet.Cells["C2"].Value = 1;
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void CountBlankNumericString()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B2"].Formula = "COUNTBLANK(C2)";
				worksheet.Cells["C2"].Value = "1";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void CountBlankWithNonNumericString()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B2"].Formula = "COUNTBLANK(C2)";
				worksheet.Cells["C2"].Value = "word";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void CountBlankWithBoolean()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B2"].Formula = "COUNTBLANK(C2)";
				worksheet.Cells["C2"].Value = true;
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void CountBlankWithDate()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B2"].Formula = "COUNTBLANK(C2)";
				worksheet.Cells["C2"].Value = (new DateTime(2017, 7, 10)).ToOADate();
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void CountBlankCountsDefaultEmptyCellsInCellRange()
		{
			// Note that EPPlus does not include cells that have not been explicitly set
			// in the group of cells in the IRangeInfo. i.e. for the specified range in this test
			// (A1:B4), only the cells that have had their value set (A1 and B2) will be included
			// in the IRangeInfo passed to the COUNTBLANK function.
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["A1"].Value = 1;
				worksheet.Cells["B2"].Value = string.Empty;
				worksheet.Cells["A5"].Formula = "COUNTBLANK(A1:B4)";
				worksheet.Calculate();
				Assert.AreEqual(7d, worksheet.Cells["A5"].Value);
			}
		}
		#endregion
	}
}
