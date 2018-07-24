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
	public class WorkdayIntlTests : DateTimeFunctionsTestBase
	{
		#region Workday Function (Execute) Tests
		// The below Test Cases have no Holiday parameter supplied to them.
		// The below Test Cases have no negative second parameters.
		[TestMethod]
		public void WorkdayIntlWithOADateParameterReturnsCorrectResult()
		{
			var inputDate = new DateTime(2014, 1, 1).ToOADate();
			var expectedDate = new DateTime(2014, 1, 29).ToOADate();
			var function = new WorkdayIntl();
			var args = FunctionsHelper.CreateArgs(inputDate, 20);
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate, result.Result);
		}

		[TestMethod]
		public void WorkdayIntlShouldReturnCorrectResultWithFourDaysSupplied()
		{
			var inputDate = new DateTime(2014, 1, 1).ToOADate();
			var expectedDate = new DateTime(2014, 1, 7).ToOADate();
			var function = new WorkdayIntl();
			var args = FunctionsHelper.CreateArgs(inputDate, 4);
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate, result.Result);
		}

		[TestMethod]
		public void WorkdayIntlWithInvalidArgumentReturnsPoundValue()
		{
			var function = new WorkdayIntl();
			var args = FunctionsHelper.CreateArgs();
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WorkdayIntlWithStringInputReturnsPoundValue()
		{
			var function = new WorkdayIntl();
			var input1 = "testString";
			var input2 = "";
			var result1 = function.Execute(FunctionsHelper.CreateArgs(input1, 10), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(input2, 10), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result1.Result).Type);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result1.Result).Type);
		}

		[TestMethod]
		public void WorkdayIntlWithIntegerInputReturnsCorrectValue()
		{
			var function = new WorkdayIntl();
			var input = 10;
			var result = function.Execute(FunctionsHelper.CreateArgs(input, 10), this.ParsingContext);
			Assert.AreEqual(24.00, result.Result);
		}

		[TestMethod]
		public void WorkdayIntlWithDateAsStringReturnsCorrectValue()
		{
			var function = new WorkdayIntl();
			var input = "1/2/2017";
			var result = function.Execute(FunctionsHelper.CreateArgs(input, 10), this.ParsingContext);
			Assert.AreEqual(42751.00, result.Result);
		}

		[TestMethod]
		public void WorkdayIntlWithDATEFunctionAsInputReturnsCorrectValue()
		{
			var function = new WorkdayIntl();
			var input = new DateTime(2017, 1, 2);
			var result = function.Execute(FunctionsHelper.CreateArgs(input, 10), this.ParsingContext);
			Assert.AreEqual(42751.00, result.Result);
		}

		[TestMethod]
		public void WorkdayIntlWithDateUsingPeriodsInsteadOfSlashesReturnsCorrectValue()
		{
			// Test case where the first parameter is a date written as a string but with '.' in place of the '/'.
			// This functionality is different than that of Excel's. Excel normally returns a #VALUE! when this 
			// is entered into the function but here the date is parsed normally. 
			var function = new WorkdayIntl();
			var input = "1.2.2017";
			var result = function.Execute(FunctionsHelper.CreateArgs(input, 10), this.ParsingContext);
			Assert.AreEqual(42751.00, result.Result);
		}

		[TestMethod]
		public void WorkdayIntlWithNegativeFirstParamReturnsPoundNum()
		{
			var function = new WorkdayIntl();
			var input = -1;
			var result = function.Execute(FunctionsHelper.CreateArgs(input, 10), this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WorkdayIntlWithZeroInputReturnsCorrectValue()
		{
			var function = new WorkdayIntl();
			var input = 0;
			var result = function.Execute(FunctionsHelper.CreateArgs(input, 15), this.ParsingContext);
			Assert.AreEqual(20.00, result.Result);
		}

		[TestMethod]
		public void WorkdayIntlWithFractionInputReturnsCorrectValue()
		{
			var function = new WorkdayIntl();
			var input = 0.5;
			var result = function.Execute(FunctionsHelper.CreateArgs(input, 15), this.ParsingContext);
			Assert.AreEqual(20.00, result.Result);
		}

		[TestMethod]
		public void WorkdayIntlWithZeroAsStringInputReturnsCorrectValue()
		{
			var function = new WorkdayIntl();
			var input = "0";
			var result = function.Execute(FunctionsHelper.CreateArgs(input, 15), this.ParsingContext);
			Assert.AreEqual(20.00, result.Result);
		}

		[TestMethod]
		public void WorkdayIntlWithFractionAsStringInputReturnsCorrectValue()
		{
			var function = new WorkdayIntl();
			var input = "0.5";
			var result = function.Execute(FunctionsHelper.CreateArgs(input, 15), this.ParsingContext);
			Assert.AreEqual(20.00, result.Result);
		}

		[TestMethod]
		public void WorkdayIntlWithNullFirstParamReturnsPoundNum()
		{
			var function = new WorkdayIntl();
			var result = function.Execute(FunctionsHelper.CreateArgs(null, 10), this.ParsingContext);
			Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WorkdayIntlWithDateUsingDashesInsteadOfSlashesReturnsCorrectResult()
		{
			var function = new WorkdayIntl();
			var input = "1-2-2017";
			var result = function.Execute(FunctionsHelper.CreateArgs(input, 10), this.ParsingContext);
			Assert.AreEqual(42751.00, result.Result);
		}

		// The below Test Cases have no Holiday parameter supplied to them.
		// The below Test Cases have negative second parameters. 
		[TestMethod]
		public void WorkdayIntlShouldReturnCorrectResultWithNegativeArg()
		{
			var inputDate = new DateTime(2016, 6, 15).ToOADate();
			var expectedDate = new DateTime(2016, 5, 4).ToOADate();
			var func = new WorkdayIntl();
			var args = FunctionsHelper.CreateArgs(inputDate, -30);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(DateTime.FromOADate(expectedDate), DateTime.FromOADate((double)result.Result));
		}

		[TestMethod]
		public void WorkdayIntlWithDATEFunctionAndNegativeDayInputReturnsCorrectResult()
		{
			var function = new WorkdayIntl();
			var inputDate = new DateTime(2017, 1, 2);
			var result = function.Execute(FunctionsHelper.CreateArgs(inputDate, -10), this.ParsingContext);
			Assert.AreEqual(42723.00, result.Result);
		}

		[TestMethod]
		public void WorkdayIntlWithDateAsStringAndNegativeDayInputReturnsCorrectValue()
		{
			var function = new WorkdayIntl();
			var inputDate = "1/2/2017";
			var result = function.Execute(FunctionsHelper.CreateArgs(inputDate, -10), this.ParsingContext);
			Assert.AreEqual(42723.00, result.Result);
		}

		[TestMethod]
		public void WorkdayIntlWithGenericStringAndNegativeDayInputReturnsPoundValue()
		{
			var function = new WorkdayIntl();
			var inputDate1 = "testString";
			var inputDate2 = "";
			var result1 = function.Execute(FunctionsHelper.CreateArgs(inputDate1, -10), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(inputDate2, -10), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result1.Result).Type);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result2.Result).Type);
		}

		[TestMethod]
		public void WorkdayIntlWithEmptyFirstParameterAndNegativeDateInputReturnsPoundNA()
		{
			var function = new WorkdayIntl();
			var result = function.Execute(FunctionsHelper.CreateArgs(null, -10), this.ParsingContext);
			Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WorkdayIntlWithIntFirstInputNegativeDayInputReturnsPoundNum()
		{
			var function = new WorkdayIntl();
			var inputDate = 10;
			var result = function.Execute(FunctionsHelper.CreateArgs(inputDate, -10), this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WorkdayIntlWithDateUsingPeriodsInsteadOfSlashesAndNegativeDayInputReturnsCorrectValue()
		{
			// Test case where the second argument is negative and the first argument is the date as a string
			// with '.' instead of '/'.
			// This functionality is different than Excel's. Excel normally returns a #VALUE! when the date is written
			// this way, but here we parse the date normally (as if it were written as "1/2/2017", for example). 
			var function = new WorkdayIntl();
			var inputDate = "1.2.2017";
			var result = function.Execute(FunctionsHelper.CreateArgs(inputDate, -10), this.ParsingContext);
			Assert.AreEqual(42723.00, result.Result);
		}

		[TestMethod]
		public void WorkdayIntlWithNegativeFirstInputAndNegativeDateInputReturnsPoundNum()
		{
			var function = new WorkdayIntl();
			var inputDate = -1;
			var result = function.Execute(FunctionsHelper.CreateArgs(inputDate, -10), this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WorkdayIntlWithZeroFirstInputNegativeDateInputReturnsPoundNum()
		{
			var function = new WorkdayIntl();
			var result = function.Execute(FunctionsHelper.CreateArgs(0, -10), this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WorkdayIntlWithNullSecondParameterReturnsPoundNA()
		{
			var function = new WorkdayIntl();
			var result = function.Execute(FunctionsHelper.CreateArgs("1/2/2017", null), this.ParsingContext);
			Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WorkdayIntlDateWithSlashesFirstInputAndNegativeDateInputReturnsCorrectInput()
		{
			var function = new WorkdayIntl();
			var result = function.Execute(FunctionsHelper.CreateArgs("1-2-2017", -10), this.ParsingContext);
			Assert.AreEqual(42723.00, result.Result);
		}

		// The below Test Cases involve changes only in the 'Days' Parameter.

		[TestMethod]
		public void WorkdayIntlWithDayParameterAsDATEFunctionReturnsCorrectValue()
		{
			var function = new WorkdayIntl();
			var daysInput = new DateTime(2017, 1, 13);
			var startDate = new DateTime(2017, 1, 1);
			var result = function.Execute(FunctionsHelper.CreateArgs(startDate, daysInput), this.ParsingContext);
			Assert.AreEqual(102582.00, result.Result);
		}

		[TestMethod]
		public void WorkdayIntlWithDayParameterTypedWithQuotesReturnsCorrectResult()
		{
			var function = new WorkdayIntl();
			var daysInput = "1/13/2017";
			var startDate = new DateTime(2017, 1, 1);
			var result = function.Execute(FunctionsHelper.CreateArgs(startDate, daysInput), this.ParsingContext);
			Assert.AreEqual(102582.00, result.Result);
		}

		[TestMethod]
		public void WorkdayIntlWithDayParameterAsGenericStringReturnsPoundValue()
		{
			var function = new WorkdayIntl();
			var daysInput1 = "testString";
			var daysInput2 = "";
			var startDate = new DateTime(2017, 1, 1);
			var result = function.Execute(FunctionsHelper.CreateArgs(startDate, daysInput1), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(startDate, daysInput2), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result2.Result).Type);
		}

		[TestMethod]
		public void WorkdayIntlWithNullDayParameterReturnsPoundNA()
		{
			var function = new WorkdayIntl();
			var startDate = new DateTime(2017, 1, 1);
			var result = function.Execute(FunctionsHelper.CreateArgs(startDate, null), this.ParsingContext);
			Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WorkdayIntlWithDayParameterAsOADateReturnCorrectValue()
		{
			var function = new WorkdayIntl();
			var startDate = new DateTime(2017, 1, 1);
			var result = function.Execute(FunctionsHelper.CreateArgs(startDate, 42748), this.ParsingContext);
			Assert.AreEqual(102582.00, result.Result);
		}

		[TestMethod]
		public void WorkdayIntlWithDayParameterWithDotInsteadOfSlashForDateReturnsCorrectValue()
		{
			// Test case where the second argument is the date as a string with '.' instead of '/'.
			// This functionality is different than Excel's. Excel normally returns a #VALUE! when the date is written
			// this way, but here we parse the date normally (as if it were written as "1/2/2017", for example). 
			var function = new WorkdayIntl();
			var startDate = new DateTime(2017, 1, 1);
			var dayInput = "1.13.2017";
			var result = function.Execute(FunctionsHelper.CreateArgs(startDate, dayInput), this.ParsingContext);
			Assert.AreEqual(102582.00, result.Result);
		}

		[TestMethod]
		public void WorkdayIntlWithDayParameterWithDashInsteadOfSlashForDateReturnsCorrectValue()
		{
			var function = new WorkdayIntl();
			var startDate = new DateTime(2017, 1, 1);
			var dayInput = "1-13-2017";
			var result = function.Execute(FunctionsHelper.CreateArgs(startDate, dayInput), this.ParsingContext);
			Assert.AreEqual(102582.00, result.Result);
		}

		[TestMethod]
		public void WorkdayIntlWithZeroAsDateParameterReuturnsCorrectValeue()
		{
			var function = new WorkdayIntl();
			var startDate = new DateTime(2017, 1, 1);
			var result = function.Execute(FunctionsHelper.CreateArgs(startDate, 0), this.ParsingContext);
			Assert.AreEqual(42736.00, result.Result);
		}

		// The below Test Cases only involve the Holiday parameter.

		[TestMethod]
		public void WorkdayIntlWithNegativeArgShouldReturnCorrectWhenArrayOfHolidayDatesIsSupplied()
		{
			var inputDate = new DateTime(2016, 7, 27).ToOADate();
			var holidayDate1 = new DateTime(2016, 7, 11).ToOADate();
			var holidayDate2 = new DateTime(2016, 7, 8).ToOADate();
			var expectedDate = new DateTime(2016, 6, 13).ToOADate();
			var func = new WorkdayIntl();
			var args = FunctionsHelper.CreateArgs(inputDate, -30, 1, FunctionsHelper.CreateArgs(holidayDate1, holidayDate2));
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate, result.Result);
		}

		[TestMethod]
		public void WorkdayIntlWithNegativeArgShouldReturnCorrectWhenRangeWithHolidayDatesIsSupplied()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("test");
				ws.Cells["A1"].Value = new DateTime(2016, 7, 27).ToOADate();
				ws.Cells["B1"].Value = new DateTime(2016, 7, 11).ToOADate();
				ws.Cells["B2"].Value = new DateTime(2016, 7, 8).ToOADate();
				ws.Cells["B3"].Formula = "WORKDAY.INTL(A1,-30, 1, B1:B2)";
				ws.Calculate();
				var expectedDate = new DateTime(2016, 6, 13).ToOADate();
				var actualDate = ws.Cells["B3"].Value;
				Assert.AreEqual(expectedDate, actualDate);
			}
		}

		[TestMethod]
		public void WorkdayIntlWithPositiveArgsAndNullHolidayDatesReturnsCorrectValue()
		{
			var function = new WorkdayIntl();
			var inputDate = new DateTime(2017, 1, 2);
			var result = function.Execute(FunctionsHelper.CreateArgs(inputDate, 10, null, null), this.ParsingContext);
			Assert.AreEqual(42751.00, result.Result);
		}

		[TestMethod]
		public void WorkdayIntlWithStringsAsHolidayInputReturnsPoundValue()
		{
			var function = new WorkdayIntl();
			var inputDate = new DateTime(2017, 1, 2);
			var result1 = function.Execute(FunctionsHelper.CreateArgs(inputDate, 10, 1, "testString"), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(inputDate, 10, 1, ""), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result1.Result).Type);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result2.Result).Type);
		}

		[TestMethod]
		public void WorkdayIntlWithNegativeIntAsHolidayInputReturnsPoundNum()
		{
			var function = new WorkdayIntl();
			var inputDate = new DateTime(2017, 1, 2);
			var result = function.Execute(FunctionsHelper.CreateArgs(inputDate, 10, 1, -1), this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WorkdayIntlWithHolidaysAsStringsReturnsCorrectValue()
		{
			var function = new WorkdayIntl();
			var result = function.Execute(FunctionsHelper.CreateArgs("1/2/2017", 41, null, "1/25/2017"), this.ParsingContext);
			Assert.AreEqual(42795.00, result.Result);
		}

		[TestMethod]
		public void WorkdayIntlWithHolidayWithDashesInsteadOfSlashesReturnsCorrectValue()
		{ 
			var function = new WorkdayIntl();
			var inputDate = new DateTime(2017, 1, 2);
			var result = function.Execute(FunctionsHelper.CreateArgs(inputDate, 41, 1, "1-25-2017"), this.ParsingContext);
			Assert.AreEqual(42795.00, result.Result);
		}

		[TestMethod]
		public void WorkdayIntlWithHolidayWithDotsInsteadOfSlashesReturnsCorrectValue()
		{
			// Test case where the third parameter is a date as a string with '.' instead of '/'.
			// This functionality is different than Excel's. Excel normally returns a #VALUE! when the date is written
			// this way, but here we parse the date normally (as if it were written as "1/2/2017", for example). 
			var function = new WorkdayIntl();
			var inputDate = new DateTime(2017, 1, 2);
			var result = function.Execute(FunctionsHelper.CreateArgs(inputDate, 500, 1, "3.30.2017"), this.ParsingContext);
			Assert.AreEqual(43438.00, result.Result);
		}

		[TestMethod]
		public void WorkdayIntlWithZeroIntegerAsHolidayInputReturnsCorrectValue()
		{
			var function = new WorkdayIntl();
			var inputDate = new DateTime(2017, 1, 2);
			var result = function.Execute(FunctionsHelper.CreateArgs(inputDate, 40, 1, 0), this.ParsingContext);
			Assert.AreEqual(42793.00, result.Result);
		}

		[TestMethod]
		public void WorkdayIntlWithNonZeroIntegerAsHolidayInputReturnsCorrectValue()
		{
			var function = new WorkdayIntl();
			var inputDate = new DateTime(2017, 1, 2);
			var result = function.Execute(FunctionsHelper.CreateArgs(inputDate, 5, 1, 1), this.ParsingContext);
			Assert.AreEqual(42744.00, result.Result);
		}

		[TestMethod]
		public void WorkdayIntlWithDateFunctionHolidayInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["A1"].Value = new DateTime(2017, 1, 2);
				ws.Cells["B1"].Value = new DateTime(2017, 1, 20);
				ws.Cells["B2"].Value = new DateTime(2017, 1, 25);
				ws.Cells["B3"].Formula = "WORKDAY.INTL(A1,40,1,B1:B2)";
				ws.Calculate();
				var actualDate = ws.Cells["B3"].Value;
				Assert.AreEqual(42795.00, actualDate);
			}
		}

		[TestMethod]
		public void WorkdayIntlWithLargeNumberOfHolidaysReturnsCorrectInput()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("test");
				ws.Cells["A1"].Value = "1/2/2017";
				ws.Cells["B1"].Value = "1/7/2017";
				ws.Cells["B2"].Value = "1/8/2017";
				ws.Cells["B3"].Value = "1/9/2017";
				ws.Cells["b4"].Value = "1/10/2017";
				ws.Cells["B5"].Value = "1/11/2017";
				ws.Cells["B6"].Value = "1/12/2017";
				ws.Cells["B7"].Value = "1/13/2017";
				ws.Cells["B8"].Value = "1/14/2017";
				ws.Cells["B9"].Value = "1/15/2017";
				ws.Cells["B10"].Value = "1/16/2017";
				ws.Cells["B11"].Value = "1/17/2017";
				ws.Cells["B12"].Value = "1/18/2017";
				ws.Cells["B13"].Value = "1/19/2017";
				ws.Cells["B14"].Value = "1/20/2017";
				ws.Cells["B15"].Value = "1/21/2017";
				ws.Cells["B16"].Value = "1/22/2017";
				ws.Cells["B17"].Value = "1/23/2017";
				ws.Cells["B18"].Value = "1/24/2017";
				ws.Cells["B19"].Value = "1/25/2017";
				ws.Cells["B20"].Value = "1/26/2017";
				ws.Cells["B21"].Value = "1/27/2017";
				ws.Cells["B22"].Value = "1/28/2017";
				ws.Cells["B23"].Value = "1/29/2017";
				ws.Cells["B24"].Value = "1/30/2017";
				ws.Cells["B25"].Value = "1/31/2017";
				ws.Cells["B26"].Value = "2/1/2017";
				ws.Cells["B27"].Value = "2/2/2017";
				ws.Cells["B28"].Value = "2/3/2017";
				ws.Cells["B29"].Value = "2/4/2017";
				ws.Cells["B30"].Value = "2/5/2017";
				ws.Cells["B31"].Value = "2/6/2017";
				ws.Cells["B32"].Value = "2/7/2017";
				ws.Cells["B33"].Value = "2/8/2017";
				ws.Cells["B34"].Value = "2/9/2017";
				ws.Cells["B35"].Value = "2/10/2017";
				ws.Cells["B36"].Value = "2/11/2017";
				ws.Cells["B37"].Value = "2/12/2017";
				ws.Cells["B38"].Value = "2/13/2017";
				ws.Cells["B39"].Value = "3/15/2017";
				ws.Cells["B40"].Value = "4/1/2017";
				ws.Cells["B41"].Value = "5/4/2017";
				ws.Cells["C1"].Formula = "WORKDAY.INTL(A1, 150, 1, B1:B41)";
				ws.Calculate();
				Assert.AreEqual(42985.00, ws.Cells["C1"].Value);
			}
		}

		[TestMethod]
		public void WorkdayIntlWithLargeNumberOfHolidaysAsOADatesReturnsCorrectInput()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("test");
				ws.Cells["A1"].Value = new DateTime(2017, 1, 2).ToOADate();
				ws.Cells["B1"].Value = new DateTime(2017, 1, 7).ToOADate();
				ws.Cells["B2"].Value = new DateTime(2017, 1, 8).ToOADate();
				ws.Cells["B3"].Value = new DateTime(2017, 1, 9).ToOADate();
				ws.Cells["b4"].Value = new DateTime(2017, 1, 10).ToOADate();
				ws.Cells["B5"].Value = new DateTime(2017, 1, 11).ToOADate();
				ws.Cells["B6"].Value = new DateTime(2017, 1, 12).ToOADate();
				ws.Cells["B7"].Value = new DateTime(2017, 1, 13).ToOADate();
				ws.Cells["B8"].Value = new DateTime(2017, 1, 14).ToOADate();
				ws.Cells["B9"].Value = new DateTime(2017, 1, 15).ToOADate();
				ws.Cells["B10"].Value = new DateTime(2017, 1, 16).ToOADate();
				ws.Cells["B11"].Value = new DateTime(2017, 1, 17).ToOADate();
				ws.Cells["B12"].Value = new DateTime(2017, 1, 18).ToOADate();
				ws.Cells["B13"].Value = new DateTime(2017, 1, 19).ToOADate();
				ws.Cells["B14"].Value = new DateTime(2017, 1, 20).ToOADate();
				ws.Cells["B15"].Value = new DateTime(2017, 1, 21).ToOADate();
				ws.Cells["B16"].Value = new DateTime(2017, 1, 22).ToOADate();
				ws.Cells["B17"].Value = new DateTime(2017, 1, 23).ToOADate();
				ws.Cells["B18"].Value = new DateTime(2017, 1, 24).ToOADate();
				ws.Cells["B19"].Value = new DateTime(2017, 1, 25).ToOADate();
				ws.Cells["B20"].Value = new DateTime(2017, 1, 26).ToOADate();
				ws.Cells["B21"].Value = new DateTime(2017, 1, 27).ToOADate();
				ws.Cells["B22"].Value = new DateTime(2017, 1, 28).ToOADate();
				ws.Cells["B23"].Value = new DateTime(2017, 1, 29).ToOADate();
				ws.Cells["B24"].Value = new DateTime(2017, 1, 30).ToOADate();
				ws.Cells["B25"].Value = new DateTime(2017, 1, 31).ToOADate();
				ws.Cells["B26"].Value = new DateTime(2017, 2, 1).ToOADate();
				ws.Cells["B27"].Value = new DateTime(2017, 2, 2).ToOADate();
				ws.Cells["B28"].Value = new DateTime(2017, 2, 3).ToOADate();
				ws.Cells["B29"].Value = new DateTime(2017, 2, 4).ToOADate();
				ws.Cells["B30"].Value = new DateTime(2017, 2, 5).ToOADate();
				ws.Cells["B31"].Value = new DateTime(2017, 2, 6).ToOADate();
				ws.Cells["B32"].Value = new DateTime(2017, 2, 7).ToOADate();
				ws.Cells["B33"].Value = new DateTime(2017, 2, 8).ToOADate();
				ws.Cells["B34"].Value = new DateTime(2017, 2, 9).ToOADate();
				ws.Cells["B35"].Value = new DateTime(2017, 2, 10).ToOADate();
				ws.Cells["B36"].Value = new DateTime(2017, 2, 11).ToOADate();
				ws.Cells["B37"].Value = new DateTime(2017, 2, 12).ToOADate();
				ws.Cells["B38"].Value = new DateTime(2017, 2, 13).ToOADate();
				ws.Cells["B39"].Value = new DateTime(2017, 3, 15).ToOADate();
				ws.Cells["B40"].Value = new DateTime(2017, 4, 1).ToOADate();
				ws.Cells["B41"].Value = new DateTime(2017, 5, 4).ToOADate();
				ws.Cells["C1"].Formula = "WORKDAY.INTL(A1, 150, 1, B1:B41)";
				ws.Calculate();
				Assert.AreEqual(42985.00, ws.Cells["C1"].Value);
			}
		}

		[TestMethod]
		public void WorkdayIntlWithGermanDateAsStringWithPeriodReturnsCorrectResult()
		{
			var currentCulture = Thread.CurrentThread.CurrentCulture;
			try
			{
				Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("de-DE");
				var function = new WorkdayIntl();
				var args = FunctionsHelper.CreateArgs("2.1.2017", 40);
				var result = function.Execute(args, this.ParsingContext);
				Assert.AreEqual(42793.00, result.Result);
			}
			finally
			{
				Thread.CurrentThread.CurrentCulture = currentCulture;
			}
		}

		[TestMethod]
		public void WorkdayIntlFunctionWithErrorValuesAsInputReturnsTheInputErrorValue()
		{
			var func = new WorkdayIntl();
			var argNA = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.NA),2);
			var argNAME = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Name),2);
			var argVALUE = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Value),2);
			var argNUM = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Num),2);
			var argDIV0 = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Div0),2);
			var argREF = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Ref),2);
			var resultNA = func.Execute(argNA, this.ParsingContext);
			var resultNAME = func.Execute(argNAME, this.ParsingContext);
			var resultVALUE = func.Execute(argVALUE, this.ParsingContext);
			var resultNUM = func.Execute(argNUM, this.ParsingContext);
			var resultDIV0 = func.Execute(argDIV0, this.ParsingContext);
			var resultREF = func.Execute(argREF, this.ParsingContext);
			Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)resultNA.Result).Type);
			Assert.AreEqual(eErrorType.Name, ((ExcelErrorValue)resultNAME.Result).Type);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)resultVALUE.Result).Type);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)resultNUM.Result).Type);
			Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)resultDIV0.Result).Type);
			Assert.AreEqual(eErrorType.Ref, ((ExcelErrorValue)resultREF.Result).Type);
		}

		// The below Test Cases only involve the Weekend Parameter.

		[TestMethod]
		public void WorkdayIntlFunctionWithValidWeekendCodeReturnsCorrectResult()
		{
			var function = new WorkdayIntl();
			var inputDate = new DateTime(2018, 1, 1).ToOADate();
			var expectedDate = new DateTime(2018, 1, 25).ToOADate();
			var args = FunctionsHelper.CreateArgs(inputDate, 17, 3);
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate, result.Result);
		}

		[TestMethod]
		public void WorkdayIntlFunctionWithWeekendCodeZeroReturnsPoundNum()
		{
			var function = new WorkdayIntl();
			var inputDate = new DateTime(2012, 6, 1).ToOADate();
			var args = FunctionsHelper.CreateArgs(inputDate, 30, 0);
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WorkdayIntlFunctionWithNegativeWeekendCodeReturnsPoundNum()
		{
			var function = new WorkdayIntl();
			var inputDate = new DateTime(2007, 10, 20).ToOADate();
			var args = FunctionsHelper.CreateArgs(inputDate, 30, -1);
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WorkdayIntlFunctionWithOutOfRangeWeekendCodeReturnsPoundNum()
		{
			var function = new WorkdayIntl();
			var inputDate = new DateTime(2017, 9, 18).ToOADate();
			var args = FunctionsHelper.CreateArgs(inputDate, 30, 23);
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WorkdayIntlFunctionWithOmittedWeekendCodeReturnsCorrectResult()
		{
			var function = new WorkdayIntl();
			var inputDate = new DateTime(2018, 1, 1).ToOADate();
			var expectedDate = new DateTime(2018, 1, 18).ToOADate();
			var args = FunctionsHelper.CreateArgs(inputDate, 13, null);
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate, result.Result);
		}

		[TestMethod]
		public void WorkdayIntlWithNegativeDayArgReturndCorrectResultsWithWeekendSupplied()
		{
			var function = new WorkdayIntl();
			var inputDate = new DateTime(2017, 6, 25).ToOADate();
			var expectedDate = new DateTime(2017, 6, 14).ToOADate();
			var args = FunctionsHelper.CreateArgs(inputDate, -8, 1);
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate, result.Result);
		}

		[TestMethod]
		public void WorkdayIntlWithStringAsWeekendCodeReturnsPoundValue()
		{
			var function = new WorkdayIntl();
			var inputDate = new DateTime(2015, 4, 17).ToOADate();
			var args = FunctionsHelper.CreateArgs(inputDate, 6, "");
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WorkdayIntlWithDecimalAsWeekendCodeReturnsPoundNum()
		{
			var function = new WorkdayIntl();
			var inputDate = new DateTime(2013, 11, 23).ToOADate();
			var args = FunctionsHelper.CreateArgs(inputDate, 20, 0.1);
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WorkdayIntlWithValidStringWeekendCodeReturnsCorrectResult()
		{
			var function = new WorkdayIntl();
			var inputDate = new DateTime(2015, 12, 5).ToOADate();
			var expectedDate = new DateTime(2015, 12, 19).ToOADate();
			var args = FunctionsHelper.CreateArgs(inputDate, 10, "1100000");
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate, result.Result);
		}

		[TestMethod]
		public void WorkdayIntlWithInvalidStringWeekendCodeReturnsPoundValue()
		{
			var function = new WorkdayIntl();
			var inputDate = new DateTime(2019, 5, 7).ToOADate();
			var args = FunctionsHelper.CreateArgs(inputDate, 6, "1111111");
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WorkdayIntlWithIncorrectStringLengthWeekendCodeReturnsPoundValue()
		{
			var function = new WorkdayIntl();
			var inputDate = new DateTime(2019, 5, 7).ToOADate();
			var args = FunctionsHelper.CreateArgs(inputDate, 6, "10100");
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		// The below Test Cases involve the Weekend and Holiday Parameter

		[TestMethod]
		public void WorkdayIntlWithWeekendAndOneHolidayReturnsCorrectResults()
		{
			var function = new WorkdayIntl();
			var inputDate = new DateTime(2018, 7, 20).ToOADate();
			var expectedDate = new DateTime(2018, 8, 21).ToOADate();
			var holidayDate = new DateTime(2018, 7, 25).ToOADate();
			var args = FunctionsHelper.CreateArgs(inputDate, 23, 6, holidayDate);
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate, result.Result);
		}

		[TestMethod]
		public void WorkdayIntlWithWeekendAndMultipleHolidayReturnsCorrectResults()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("test");
				ws.Cells["A1"].Value = new DateTime(2018, 7, 20);
				ws.Cells["B1"].Value = new DateTime(2018, 7, 25);
				ws.Cells["B2"].Value = new DateTime(2018, 8, 5);
				ws.Cells["B3"].Value = new DateTime(2018, 8, 19);
				ws.Cells["C1"].Formula = "WORKDAY.INTL(A1, 23, 6, B1:B3)";
				ws.Calculate();
				Assert.AreEqual(43337.00, ws.Cells["C1"].Value);
			}
		}
		#endregion
	}
}
