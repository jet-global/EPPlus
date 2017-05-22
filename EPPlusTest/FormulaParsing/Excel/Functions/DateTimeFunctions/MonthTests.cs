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
using System.Globalization;
using System.Threading;
using EPPlusTest.Excel.Functions.DateTimeFunctions;
using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;

namespace EPPlusTest.FormulaParsing.Excel.Functions.DateTimeFunctions
{
	[TestClass]
	public class MonthTests : DateTimeFunctionsTestBase
	{
		#region Month Function (Execute) Tests
		[TestMethod]
		public void MonthShouldReturnMonthOfYear()
		{
			var date = new DateTime(2012, 3, 12);
			var func = new Month();
			var result = func.Execute(FunctionsHelper.CreateArgs(date.ToOADate()), this.ParsingContext);
			Assert.AreEqual(3, result.Result);
		}

		[TestMethod]
		public void MonthShouldReturnMonthOfYearWithStringParam()
		{
			// Test the case where the date is entered in date format
			// within a string.
			var func = new Month();
			var result = func.Execute(FunctionsHelper.CreateArgs("2012-03-12"), this.ParsingContext);
			Assert.AreEqual(3, result.Result);
		}

		[TestMethod]
		public void MonthWithLongDateAsStringReturnsMonthOfYear()
		{
			// Test the case where the date is entered in 
			// long date format within a string.
			var func = new Month();
			var result = func.Execute(FunctionsHelper.CreateArgs("March 12, 2012"), this.ParsingContext);
			Assert.AreEqual(3, result.Result);
		}

		[TestMethod]
		public void MonthWithInvalidArgumentReturnsPoundValue()
		{

			// Test the case where the nothing is entered in
			// the MONTH function.
			var func = new Month();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void MonthFunctionWorksInDifferentCultureDateFormats()
		{
			// Test the case where the dates are entered into
			// the MONTH function under different cultures.
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
			// Test the case where the date is entered as an integer.
			// Note that 42874 is the OADate for May 19, 2017.
			var func = new Month();
			var args = FunctionsHelper.CreateArgs(42874);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(5, result.Result);
		}

		[TestMethod]
		public void MonthWithDateAsDoubleReturnsMonthInYear()
		{
			// Test the case where the date is entered as a double.
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
			// Test the case where zero is the input.
			var func = new Month();
			var args = FunctionsHelper.CreateArgs(0);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(1, result.Result);
		}

		[TestMethod]
		public void MonthWithNegativeIntegerAsInputReturnsPoundNum()
		{
			// Test the case where a negative integer is the input.
			var func = new Month();
			var args = FunctionsHelper.CreateArgs(-1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void MonthWithNegativeDoubleAsInputReturnsPoundNum()
		{
			// Test the case where a negative double is the input
			var func = new Month();
			var args = FunctionsHelper.CreateArgs(-1.5);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void MonthWithNonDateStringAsInputReturnsPoundValue()
		{
			// Test the case where a non-date string is the input.
			var func = new Month();
			var args = FunctionsHelper.CreateArgs("word");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void MonthWithEmptyStringAsInputReturnsPoundValue()
		{
			// Test the case where an empty string is the input.
			var func = new Month();
			var args = FunctionsHelper.CreateArgs("");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void MonthWithIntegerInStringAsInputReturnsMonthInYear()
		{
			// Test the case where the input is an integer expressed as a string.
			// Note that 42874 is the OADate representation of May 19, 2017.
			var func = new Month();
			var args = FunctionsHelper.CreateArgs("42874");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(5, result.Result);
		}

		[TestMethod]
		public void MonthWithDoubleInStringAsInputReturnsMonthInYear()
		{
			// Test the case where the input is a double expressed as a string.
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
			// Test the case where the input is a negative integer expressed
			// as a string.
			var func = new Month();
			var args = FunctionsHelper.CreateArgs("-1");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void MonthWithNegativeDoubleInStringAsInputReturnsPoundNum()
		{
			// Test the case where the input is a negative double expressed
			// as a string.
			var func = new Month();
			var args = FunctionsHelper.CreateArgs("-1.5");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}
		#endregion
	}
}
