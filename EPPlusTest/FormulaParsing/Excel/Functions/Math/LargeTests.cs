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
	public class LargeTests : MathFunctionsTestBase
	{
		#region Large Function (Execute) Tests
		[TestMethod]
		public void LargeWithCorrectInputsReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["A2"].Value = 45;
				ws.Cells["A3"].Value = 2;
				ws.Cells["A4"].Value = 789;
				ws.Cells["A5"].Value = 3;
				ws.Cells["B1"].Formula = "LARGE(A2:A5, 2)";
				ws.Calculate();
				Assert.AreEqual(45d, ws.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void LargeWithNullFirstParameterReturnsPoundNum()
		{
			var function = new Large();
			var result = function.Execute(FunctionsHelper.CreateArgs(null, 4), this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void LargeWithNullSecondParameterReturnsPoundNum()
		{
			var function = new Large();
			var result = function.Execute(FunctionsHelper.CreateArgs(4, null), this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void LargeWithNegativeSecondInputReturnsPoundNum()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["A2"].Value = 45;
				ws.Cells["A3"].Value = 2;
				ws.Cells["A4"].Value = 789;
				ws.Cells["A5"].Value = 3;
				ws.Cells["B1"].Formula = "LARGE(A2:A5, -56)";
				ws.Calculate();
				Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)ws.Cells["B1"].Value).Type);
			}
		}

		[TestMethod]
		public void LargeWithSecondInputOutOfArrayIndexRangeReturnsPoundNum()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["A2"].Value = 45;
				ws.Cells["A3"].Value = 2;
				ws.Cells["A4"].Value = 789;
				ws.Cells["A5"].Value = 3;
				ws.Cells["B1"].Formula = "LARGE(A2:A5, 67)";
				ws.Calculate();
				Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)ws.Cells["B1"].Value).Type);
			}
		}

		[TestMethod]
		public void LargeWithSecondInputAsNumericStringReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["A2"].Value = 45;
				ws.Cells["A3"].Value = 2;
				ws.Cells["A4"].Value = 789;
				ws.Cells["A5"].Value = 3;
				ws.Cells["B1"].Formula = "LARGE(A2:A5, \"3\")";
				ws.Calculate();
				Assert.AreEqual(3d, ws.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void LargeWithSecondInputAsStringReturnsPoundValue()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["A2"].Value = 6;
				ws.Cells["A3"].Value = 67;
				ws.Cells["A4"].Value = 44;
				ws.Cells["B1"].Formula = "LARGE(A2:A4,\"string\")";
				ws.Calculate();
				Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)ws.Cells["B1"].Value).Type);
			}
		}

		[TestMethod]
		public void LargeWithArrayOfStringsReturnsPoundNum()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["A2"].Value = "string";
				ws.Cells["A3"].Value = "string";
				ws.Cells["A4"].Value = "String";
				ws.Cells["B1"].Formula = "LARGE(A2:A4, 2)";
				ws.Calculate();
				Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)ws.Cells["B1"].Value).Type);
			}
		}

		[TestMethod]
		public void LargeWithArrayOfNumericStringsReturnsPoundNum()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["A2"].Value = "2";
				ws.Cells["A3"].Value = "34";
				ws.Cells["A4"].Value = "77";
				ws.Cells["B1"].Formula = "LARGE(A2:A4, 3)";
				ws.Calculate();
				Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)ws.Cells["B1"].Value).Type);
			}
		}

		[TestMethod]
		public void LargeWithGeneralStringInputsReturnsPoundValue()
		{
			var function = new Large();
			var result = function.Execute(FunctionsHelper.CreateArgs("string", "string"), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void LargeWithDatesFromDateFunctionInputsReturnsPoundNum()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["B1"].Formula = "LARGE(DATE(2017,6,15), DATE(2017,6,20))";
				ws.Calculate();
				Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)ws.Cells["B1"].Value).Type);
			}
		}

		[TestMethod]
		public void LargeWithSecondInputAsDateFromDateFunctionReturnsPoundNum()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["A2"].Value = 56;
				ws.Cells["A3"].Value = 34;
				ws.Cells["A4"].Value = 3;
				ws.Cells["B1"].Formula = "LARGE(A2:A4, DATE(2017,6,15))";
				ws.Calculate();
				Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)ws.Cells["B1"].Value).Type);
			}
		}

		[TestMethod]
		public void LargeWithNoInputsReturnsPoundValue()
		{
			var function = new Large();
			var result = function.Execute(FunctionsHelper.CreateArgs(), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void LargeWithDatesAsStringsInputsReturnsPoundNum()
		{
			var function = new Large();
			var result = function.Execute(FunctionsHelper.CreateArgs("5/15/2017", 6), this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void LargeWithNumericStringInputReturnsPoundNum()
		{
			var function = new Large();
			var result = function.Execute(FunctionsHelper.CreateArgs("3", "56"), this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}


		[TestMethod]
		public void LargeShouldReturnTheLargestNumberIf1()
		{
			var func = new Large();
			var args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(1, 2, 3), 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(3d, result.Result);
		}

		[TestMethod]
		public void LargeShouldReturnTheSecondLargestNumberIf2()
		{
			var func = new Large();
			var args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(4, 1, 2, 3), 2);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(3d, result.Result);
		}

		[TestMethod]
		public void LargeShouldPoundNumIfIndexOutOfBounds()
		{
			var func = new Large();
			var args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(4, 1, 2, 3), 6);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(OfficeOpenXml.FormulaParsing.ExpressionGraph.DataType.ExcelError, result.DataType);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)(result.Result)).Type);
		}

		[TestMethod]
		public void LargeWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Large();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}
		#endregion
	}
}
