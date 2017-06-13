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
	public class MaxTests: MathFunctionsTestBase
	{
		#region Max Function (Execute) Tests
		[TestMethod]
		public void MaxWithNoInputsReturnsPoundValue()
		{
			var function = new Max();
			var result = function.Execute(FunctionsHelper.CreateArgs(), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void MaxWithIntegerInputReturnsCorrectValue()
		{
			var function = new Max();
			var result = function.Execute(FunctionsHelper.CreateArgs(12, 2, 99, 1, 25), this.ParsingContext);
			Assert.AreEqual(99d, result.Result);
		}

		[TestMethod]
		public void MaxWithDoublesInputReturnsCorrectValue()
		{
			var function = new Max();
			var result = function.Execute(FunctionsHelper.CreateArgs(2.5,26.5,123.87,658.64,55.6), this.ParsingContext);
			Assert.AreEqual(658.64d, result.Result);
		}

		[TestMethod]
		public void MaxWithFractionsInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "MAX((2/3),(9/8),(2/55))";
				worksheet.Calculate();
				Assert.AreEqual(1.125d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void MaxWithStringInputReturnsPoundValue()
		{
			var function = new Max();
			var result = function.Execute(FunctionsHelper.CreateArgs("string", "string"), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void MaxWithDateObjectInputsReturnsCorrectValue()
		{
			var function = new Max();
			var dateObject1 = new DateTime(2017, 6, 2);
			var dateObject2 = new DateTime(2017, 6, 15);
			var dateObjectAsOADate1 = new DateTime(2017, 6, 2).ToOADate();
			var dateObjectAsOADate2 = new DateTime(2017, 6, 15).ToOADate();

			var result1 = function.Execute(FunctionsHelper.CreateArgs(dateObject1, dateObject2), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(dateObjectAsOADate1, dateObjectAsOADate2), this.ParsingContext);

			Assert.AreEqual(42901d, result1.Result);
			Assert.AreEqual(42901d, result2.Result);
		}

		[TestMethod]
		public void MaxWithDatesAsStringsInputReturnsCorrectValue()
		{
			var function = new Max();
			var result = function.Execute(FunctionsHelper.CreateArgs("5/2/2017", "6/25/2017"), this.ParsingContext);
			Assert.AreEqual(42911d, result.Result);
		}

		[TestMethod]
		public void MaxWithReferenceToEmptyCellReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "MAX(B2:B4)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void MaxWithReferenceToCellWithLogicalValueReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 0.5;
				worksheet.Cells["B2"].Value = 0.2;
				worksheet.Cells["B3"].Value = true;
				worksheet.Cells["B10"].Formula = "MAX(B1:B3)";
				worksheet.Calculate();
				Assert.AreEqual(0.5d, worksheet.Cells["B10"].Value);
			}
		}

		[TestMethod]
		public void MaxWithReferenceToCellWithNumericStringReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 0.5;
				worksheet.Cells["B2"].Value = "1";
				worksheet.Cells["B3"].Value = 0.2;
				worksheet.Cells["B4"].Formula = "MAX(B1:B3)";
				worksheet.Calculate();
				Assert.AreEqual(0.5d, worksheet.Cells["B4"].Value);
			}
		}

		[TestMethod]
		public void MaxWithArrayWithLogicalValueReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "MAX({0.5, 0.1, TRUE})";
				worksheet.Calculate();
				Assert.AreEqual(0.5d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void MaxWithArrayWithNumericStringReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "MAX({1, 3, \"78\"})";
				worksheet.Calculate();
				Assert.AreEqual(3d, worksheet.Cells["B1"].Value);
			}
		}
		#endregion
	}
}
