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
	public class LogTests : MathFunctionsTestBase
	{
		#region Log Function (Execute) Tests
		[TestMethod]
		public void LogWithTwoPositiveIntegersReturnsCorrectValue()
		{
			var function = new Log();
			var result = function.Execute(FunctionsHelper.CreateArgs(4, 2), this.ParsingContext);
			Assert.AreEqual(2d, result.Result);
		}

		[TestMethod]
		public void LogWithNegativeIntegerAsSecondArgumentReturnsPoundNum()
		{
			var function = new Log();
			var result = function.Execute(FunctionsHelper.CreateArgs(4, -2), this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void LogWithNegativeIntegerAsFirstArugmentReturnsPoundNum()
		{
			var function = new Log();
			var result = function.Execute(FunctionsHelper.CreateArgs(-4, 2), this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void LogWithNegativeIntegersReturnsPoundNum()
		{
			var function = new Log();
			var result = function.Execute(FunctionsHelper.CreateArgs(-4, -2), this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void LogWithDoublesReturnsCorrectValue()
		{
			var function = new Log();
			var result = function.Execute(FunctionsHelper.CreateArgs(6.5, 1.3), this.ParsingContext);
			Assert.AreEqual(7.134364052d, (double)result.Result, 0.000001);
		}

		[TestMethod]
		public void LogWithFractionsInExcelWorksheetReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["B1"].Formula = "LOG((1/15),(1/5))";
				ws.Calculate();
				Assert.AreEqual(1.682606194d, (double)ws.Cells["B1"].Value, 0.000001);
			}
		}

		[TestMethod]
		public void LogWithDateFunctionArgumentsReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["B1"].Formula = "LOG(DATE(2017,5,1), DATE(2017,6,15))";
				ws.Calculate();
				Assert.AreEqual(0.999901611d, (double)ws.Cells["B1"].Value, 0.000001);
			}
		}

		[TestMethod]
		public void LogWithDateAsStringInputReturnsCorrectValue()
		{
			var function = new Log();
			var result = function.Execute(FunctionsHelper.CreateArgs("5/1/2017", "6/15/2017"), this.ParsingContext);
			Assert.AreEqual(0.999901611d, (double)result.Result, 0.0000001);
		}

		[TestMethod]
		public void LogWithNumericStringsReturnsCorrectValue()
		{
			var function = new Log();
			var result = function.Execute(FunctionsHelper.CreateArgs("4", "2"), this.ParsingContext);
			Assert.AreEqual(2d, result.Result);
		}

		[TestMethod]
		public void LogWithGeneralStringAsArgumentReturnsCorrectValue()
		{
			var function = new Log();
			var resultWithStringFirstArg = function.Execute(FunctionsHelper.CreateArgs("string", 4), this.ParsingContext);
			var resultWithStringSecondArg = function.Execute(FunctionsHelper.CreateArgs(4, "string"), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)resultWithStringFirstArg.Result).Type);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)resultWithStringSecondArg.Result).Type);
		}

		[TestMethod]
		public void LogWithNoBaseArgumentReturnsCorrectValue()
		{
			var function = new Log();
			var result = function.Execute(FunctionsHelper.CreateArgs(10), this.ParsingContext);
			Assert.AreEqual(1d, result.Result);
		}

		[TestMethod]
		public void LogWithNoArgumentsReturnsPoundValue()
		{
			var function = new Log();
			var result = function.Execute(FunctionsHelper.CreateArgs(), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void LogWithNullFirstArgumentRetursnPoundNum()
		{
			var function = new Log();
			var result = function.Execute(FunctionsHelper.CreateArgs(null, 4), this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void LogWithNullSecondArgumentReturnsPoundNum()
		{
			var function = new Log();
			var result = function.Execute(FunctionsHelper.CreateArgs(10, null), this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void LogShouldReturnCorrectResult()
		{
			var func = new Log();
			var args = FunctionsHelper.CreateArgs(2);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(0.301029996d, (double)result.Result, 0.000001);
		}

		[TestMethod]
		public void LogShouldReturnCorrectResultWithBase()
		{
			var func = new Log();
			var args = FunctionsHelper.CreateArgs(2, 2);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(1d, result.Result);
		}

		[TestMethod]
		public void Log10ShouldReturnCorrectResult()
		{
			var func = new Log10();
			var args = FunctionsHelper.CreateArgs(2);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(0.301029996d, (double)result.Result, 0.000001);
		}

		[TestMethod]
		public void LogWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Log();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}
		#endregion

	}
}
