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
using EPPlusTest.Excel.Functions.DateTimeFunctions;
using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using System;
using System.Globalization;
using System.Threading;

namespace EPPlusTest.FormulaParsing.Excel.Functions.DateTimeFunctions
{
	[TestClass]
	public class MonthTests : DateTimeFunctionsTestBase
	{
		#region Month Function (Execute) Tests
		[TestMethod]
		public void MonthWithDateAsStringReturnsMonthOfYear()
		{
			var func = new Month();
			var result = func.Execute(FunctionsHelper.CreateArgs("2012-03-12"), this.ParsingContext);
			Assert.AreEqual(3, result.Result);
		}

		[TestMethod]
		public void MonthWithLongDateAsStringReturnsMonthOfYear()
		{
			var func = new Month();
			var result = func.Execute(FunctionsHelper.CreateArgs("March 12, 2012"), this.ParsingContext);
			Assert.AreEqual(3, result.Result);
		}

		[TestMethod]
		public void MonthWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Month();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void MonthFunctionWorksInDifferentCultureDateFormats()
		{
			var currentCulture = CultureInfo.CurrentCulture;
			try
			{
				var us = CultureInfo.CreateSpecificCulture("en-US");
				Thread.CurrentThread.CurrentCulture = us;
				using (var package = new ExcelPackage())
				{
					var ws = package.Workbook.Worksheets.Add("Sheet1");
					ws.Cells[2, 2].Value = "1/15/2014";
					ws.Cells[4, 3].Formula = "MONTH(B2)";
					ws.Calculate();
					Assert.AreEqual(1, ws.Cells[4, 3].Value);
				}
				var gb = CultureInfo.CreateSpecificCulture("en-GB");
				Thread.CurrentThread.CurrentCulture = gb;
				using (var package = new ExcelPackage())
				{
					var ws = package.Workbook.Worksheets.Add("Sheet1");
					ws.Cells[2, 2].Value = "15/1/2014";
					ws.Cells[4, 3].Formula = "MONTH(B2)";
					ws.Calculate();
					Assert.AreEqual(1, ws.Cells[4, 3].Value);
				}
			}
			finally
			{
				Thread.CurrentThread.CurrentCulture = currentCulture;
			}
		}

		[TestMethod]
		public void MonthWithDateAsIntegerReturnsMonthInYear()
		{
			// Note that 42874 is the OADate for May 19, 2017.
			var func = new Month();
			var args = FunctionsHelper.CreateArgs(42874);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(5, result.Result);
		}

		[TestMethod]
		public void MonthWithDateAsDoubleReturnsMonthInYear()
		{
			// Note that 42874.34114 is the OADate representation of
			// some time on May 19, 2017.
			var func = new Month();
			var args = FunctionsHelper.CreateArgs(42874.34114);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(5, result.Result);
		}

		[TestMethod]
		public void MonthWithZeroAsInputReturnsOneAsDateInMonth()
		{
			var func = new Month();
			var args = FunctionsHelper.CreateArgs(0);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(1, result.Result);
		}

		[TestMethod]
		public void MonthWithNegativeIntegerAsInputReturnsPoundNum()
		{
			var func = new Month();
			var args = FunctionsHelper.CreateArgs(-1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void MonthWithNegativeDoubleAsInputReturnsPoundNum()
		{
			var func = new Month();
			var args = FunctionsHelper.CreateArgs(-1.5);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void MonthWithNonDateStringAsInputReturnsPoundValue()
		{
			var func = new Month();
			var args = FunctionsHelper.CreateArgs("word");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void MonthWithEmptyStringAsInputReturnsPoundValue()
		{
			var func = new Month();
			var args = FunctionsHelper.CreateArgs(string.Empty);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void MonthWithIntegerInStringAsInputReturnsMonthInYear()
		{
			// Note that 42874 is the OADate representation of May 19, 2017.
			var func = new Month();
			var args = FunctionsHelper.CreateArgs("42874");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(5, result.Result);
		}

		[TestMethod]
		public void MonthWithDoubleInStringAsInputReturnsMonthInYear()
		{
			// Note that 42874.34114 is the OADate representation of some time
			// on May 19, 2017.
			var func = new Month();
			var args = FunctionsHelper.CreateArgs("42874.43114");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(5, result.Result);
		}

		[TestMethod]
		public void MonthWithNegativeIntegerInStringAsInputReturnsPoundNum()
		{
			var func = new Month();
			var args = FunctionsHelper.CreateArgs("-1");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void MonthWithNegativeDoubleInStringAsInputReturnsPoundNum()
		{
			var func = new Month();
			var args = FunctionsHelper.CreateArgs("-1.5");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void MonthWithDateTimeObjectReturnsCorrectResult()
		{
			var date = new DateTime(2017, 5, 22);
			var func = new Month();
			var args = FunctionsHelper.CreateArgs(date);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(5, result.Result);
		}

		[TestMethod]
		public void MonthWithDateTimeObjectForOADateLessThan61ReturnsCorrectResult()
		{
			var date = new DateTime(1900, 2, 28);
			var func = new Month();
			var args = FunctionsHelper.CreateArgs(date);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(2, result.Result);
		}

		[TestMethod]
		public void MonthWithOADate1ReturnsCorrectResult()
		{
			// OADate 1 corresponds to 1/1/1900 in Excel. Test the case where Excel's epoch date
			// is used as the input.
			var func = new Month();
			var args = FunctionsHelper.CreateArgs(1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(1, result.Result);
		}

		[TestMethod]
		public void MonthWithOADateBetween0And1ReturnsZero()
		{
			// Test the case where a time value is given for Excel's
			// "zeroeth" day, which Excel treats as a special day.
			var func = new Month();
			var args = FunctionsHelper.CreateArgs(0.5);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(1, result.Result);
		}

		[TestMethod]
		public void MonthWithNegativeOADateBetweenNegative1And0ReturnsPoundNum()
		{
			// Test the case where a negative time value is given for Excel's
			// "zeroeth" day, which Excel treats as a special day.
			var func = new Month();
			var args = FunctionsHelper.CreateArgs(-0.5);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}
		#endregion
	}
}
