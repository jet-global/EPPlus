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
	public class MedianTests : MathFunctionsTestBase
	{
		#region Median Function (Execute) Tests

		[TestMethod]
		public void MedianWithNoArgumentsReturnsPoundValue()
		{
			var function = new Median();
			var result = function.Execute(FunctionsHelper.CreateArgs(), this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void MedianWithMaxArgumentsReturnsCorrectValue()
		{

		}

		[TestMethod]
		public void MedianWithOneInputReturnsCorrectValue()
		{
			var function = new Median();
			var result = function.Execute(FunctionsHelper.CreateArgs(15), this.ParsingContext);
			Assert.AreEqual(15d, result.Result);
		}

		[TestMethod]
		public void MedianWithNumericInputsReturnsCorrectValue()
		{
			var function = new Median();
			var result = function.Execute(FunctionsHelper.CreateArgs(16, 55, 19, 20), this.ParsingContext);
			Assert.AreEqual(19.5d, result.Result);
		}

		[TestMethod]
		public void MedianWithGenericStringInputReturnsPoundValue()
		{
			var function = new Median();
			var result = function.Execute(FunctionsHelper.CreateArgs("string", "string", "string"), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void MedianWithNumericStringInputReturnsCorrectValue()
		{
			var function = new Median();
			var result = function.Execute(FunctionsHelper.CreateArgs("16", "55", "19", "20"), this.ParsingContext);
			Assert.AreEqual(19.5d, result.Result);
		}

		[TestMethod]
		public void MedianWithReferenceToNumbersReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["B1"].Value = 16;
				ws.Cells["B2"].Value = 6;
				ws.Cells["B3"].Value = 5;
				ws.Cells["B10"].Formula = "MEDIAN(B1:B3)";
				ws.Calculate();
				Assert.AreEqual(6d, ws.Cells["B10"].Value);
			}
		}

		[TestMethod]
		public void MedianWithReferencesTypedOutReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["B1"].Value = 16;
				ws.Cells["B2"].Value = 6;
				ws.Cells["B3"].Value = 5;
				ws.Cells["B10"].Formula = "MEDIAN(B1,B2,B3)";
				ws.Calculate();
				Assert.AreEqual(6d, ws.Cells["B10"].Value);
			}
		}

		[TestMethod]
		public void MedianWithReferenceToNumericStringsReturnsPoundNum()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["B1"].Value = "5";
				ws.Cells["B2"].Value = "45";
				ws.Cells["B3"].Value = "76";
				ws.Cells["B10"].Formula = "MEDIAN(B1:B3)";
				ws.Calculate();
				Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)ws.Cells["B10"].Value).Type);
			}
		}

		[TestMethod]
		public void MedianWithReferencesToGeneralStringsReturnsPoundNum()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["B1"].Value = "string!";
				ws.Cells["B2"].Value = "string";
				ws.Cells["B3"].Value = "string";
				ws.Cells["B10"].Formula = "MEDIAN(B1:B3)";
				ws.Calculate();
				Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)ws.Cells["B10"].Value).Type);
			}
		}

		[TestMethod]
		public void MedianWithLogicInputsReturnsCorrectValue()
		{
			var function = new Median();
			var result = function.Execute(FunctionsHelper.CreateArgs(true, false), this.ParsingContext);
			Assert.AreEqual(0.5d, result.Result);
		}

		[TestMethod]
		public void MedianWithReferenceToLogicInputsReturnsPoundNum()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["B1"].Value = "TRUE";
				ws.Cells["B2"].Value = "FALSE";
				ws.Cells["B10"].Formula = "MEDIAN(B1:B2)";
				ws.Calculate();
				Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)ws.Cells["B10"].Value).Type);
			}
		}

		[TestMethod]
		public void MedianWithReferenceToCellsWithZeroReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["B1"].Value = 0;
				ws.Cells["B2"].Value = 16;
				ws.Cells["B3"].Value = 6;
				ws.Cells["B4"].Value = 5;
				ws.Cells["B10"].Formula = "MEDIAN(B1:B4)";
				ws.Calculate();
				Assert.AreEqual(5.5d, ws.Cells["B10"].Value);
			}
		}

		[TestMethod]
		public void MedianWithInputCellsThatHaveErrorsReturnsRespectiveError()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["B1"].Value = 99;
				ws.Cells["B2"].Value = 6;
				ws.Cells["B3"].Formula = "MEDIAN(\"strings\")";
				ws.Cells["B10"].Formula = "MEDIAN(B1:B3)";
				ws.Calculate();
				Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)ws.Cells["B10"].Value).Type);
			}
		}

		[TestMethod]
		public void MedianWithReferenceToEmptyCellsReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["B1"].Value = 16;
				ws.Cells["B2"].Value = 5;
				ws.Cells["B3"].Value = 6;
				ws.Cells["B10"].Formula = "MEDIAN(B1:B4)";
				ws.Calculate();
				Assert.AreEqual(6d, ws.Cells["B10"].Value);
			}
		}

		[TestMethod]
		public void MedianWithReferenceToStringsAndNumbersReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["B1"].Value = 5;
				ws.Cells["B2"].Value = 64;
				ws.Cells["B3"].Value = 0;
				ws.Cells["B4"].Value = "string";
				ws.Cells["B10"].Formula = "MEDIAN(B1:B4)";
				ws.Calculate();
				Assert.AreEqual(5d, ws.Cells["B10"].Value);
			}
		}

		[TestMethod]
		public void MedianWithDateObjectAsInputReturnsCorrectValue()
		{

		}
		
		[TestMethod]
		public void MedianWithDateAsStringReturnsCorrectValue()
		{

		}

		[TestMethod]
		public void MedianWithDoubleInputsReturnsCorrectValue()
		{
			var function = new Median();
			var result = function.Execute(FunctionsHelper.CreateArgs(5.5, 2.3, 15.6, 11.2), this.ParsingContext);
			Assert.AreEqual(8.35d, result.Result);
		}

		[TestMethod]
		public void MedianWithFractionInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
			}
		}
		#endregion
	}
}
