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
using EPPlusTest.Excel.Functions.DateTimeFunctions;
using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;

namespace EPPlusTest.FormulaParsing.Excel.Functions.DateTimeFunctions
{
	[TestClass]
	public class WorkdayTests : DateTimeFunctionsTestBase
	{
		// The below Test Cases have no Holiday parameter supplied to them.
		// The below Test Cases have no negative second parameters.
		#region Workday Function (Execute) Tests
		[TestMethod]
		public void WorkdayWithOADateParameterReturnsCorrectResult()
		{
			// Test case where the first input is an OADate .
			var inputDate = new DateTime(2014, 1, 1).ToOADate();
			var expectedDate = new DateTime(2014, 1, 29).ToOADate();
			var function = new Workday();
			var args = FunctionsHelper.CreateArgs(inputDate, 20);
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate, result.Result);
		}

		[TestMethod]
		public void WorkdayShouldReturnCorrectResultWithFourDaysSupplied()
		{
			// Test case where the number of workdays is four.
			var inputDate = new DateTime(2014, 1, 1).ToOADate();
			var expectedDate = new DateTime(2014, 1, 7).ToOADate();
			var function = new Workday();
			var args = FunctionsHelper.CreateArgs(inputDate, 4);
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate, result.Result);
		}

		[TestMethod]
		public void WorkdayWithInvalidArgumentReturnsPoundValue()
		{
			// Test case where the first argument is empty (invalid).
			var function = new Workday();
			var args = FunctionsHelper.CreateArgs();
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WorkdayWithStringInputReturnsPoundValue()
		{
			// Test case where the first argument is a non-date string and an empty string.
			var function = new Workday();
			var input1 = "testString";
			var input2 = "";
			var result1 = function.Execute(FunctionsHelper.CreateArgs(input1, 10), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(input2, 10), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result1.Result).Type);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result1.Result).Type);
		}

		[TestMethod]
		public void WorkdayWithIntegerInputReturnsCorrectValue()
		{
			// Test case where the first argument is an integer.
			var function = new Workday();
			var input = 10;
			var result = function.Execute(FunctionsHelper.CreateArgs(input, 10), this.ParsingContext);
			Assert.AreEqual(24.00, result.Result);
		}

		[TestMethod]
		public void WorkdayWithDateAsStringReturnsCorrectValue()
		{
			// Test case where the first argument is a date as a string.
			var function = new Workday();
			var input = "1/2/2017";
			var result = function.Execute(FunctionsHelper.CreateArgs(input, 10), this.ParsingContext);
			Assert.AreEqual(42751.00, result.Result);
		}

		[TestMethod]
		public void WorkdayWithDateNotAsStringReturnsCorrectValue()
		{
			// Test case where the first argument is a date is not written as a string.
			var function = new Workday();
			var input = 1 / 2 / 2017;
			var result = function.Execute(FunctionsHelper.CreateArgs(input, 5), this.ParsingContext);
			Assert.AreEqual(6.00, result.Result);
		}

		[TestMethod]
		public void WorkdayWithDATEFunctionAsInputReturnsCorrectValue()
		{
			// Test case where the first argument is a result of the DATE function.
			var function = new Workday();
			var input = new DateTime(2017, 1, 2);
			var result = function.Execute(FunctionsHelper.CreateArgs(input, 10), this.ParsingContext);
			Assert.AreEqual(42751.00, result.Result);
		}

		[TestMethod]
		public void WorkdayWithDateUsingPeriodsIntseadOfSlashesReturnsCorrectValue()
		{
			// Test case where the first parameter is a date written as a string but with '.' in place of the '/'.
			// This functionality is different than that of Excell's. 
			var function = new Workday();
			var input = "1.2.2017";
			var result = function.Execute(FunctionsHelper.CreateArgs(input, 10), this.ParsingContext);
			Assert.AreEqual(42751.00, result.Result);
		}

		[TestMethod]
		public void WorkdayWithNegativeFirstParamReturnsPoundNum()
		{
			// Test case where the first parameter is a negative integer.
			var function = new Workday();
			var input = -1;
			var result = function.Execute(FunctionsHelper.CreateArgs(input, 10), this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WorkdayWithZeroInputReturnsCorrectValue()
		{
			// Test case where the first parameter is 0.
			var function = new Workday();
			var input = 0;
			var result = function.Execute(FunctionsHelper.CreateArgs(input, 15), this.ParsingContext);
			Assert.AreEqual(20.00, result.Result);
		}

		[TestMethod]
		public void WorkdayWithNullFirstParamReturnsPoundNum()
		{
			// Test case where the first parameter is null.
			var function = new Workday();
			var result = function.Execute(FunctionsHelper.CreateArgs(null, 10), this.ParsingContext);
			Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WorkdayWithDateUsingDashesInsteadOfSlashesReturnsCorrectResult()
		{
			// Test case where the first parameter is a date written with '-' instead of '/'.
			var function = new Workday();
			var input = "1-2-2017";
			var result = function.Execute(FunctionsHelper.CreateArgs(input, 10), this.ParsingContext);
			Assert.AreEqual(42751.00, result.Result);
		}

		// The below Test Cases have no Holiday parameter supplied to them.
		// The below Test Cases have negative second parameters. 
		[TestMethod]
		public void WorkdayShouldReturnCorrectResultWithNegativeArg()
		{
			// Test case where the second argument is negative and the first argument is an OADate.
			var inputDate = new DateTime(2016, 6, 15).ToOADate();
			var expectedDate = new DateTime(2016, 5, 4).ToOADate();
			var func = new Workday();
			var args = FunctionsHelper.CreateArgs(inputDate, -30);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(DateTime.FromOADate(expectedDate), DateTime.FromOADate((double)result.Result));
		}

		[TestMethod]
		public void WorkdayWithDATEFunctionAndNegativeDayInputReturnsCorrectResult()
		{
			// Test case where the second argument is negative and the first argument is a result of the DATE function.
			var function = new Workday();
			var inputDate = new DateTime(2017, 1, 2);
			var result = function.Execute(FunctionsHelper.CreateArgs(inputDate, -10), this.ParsingContext);
			Assert.AreEqual(42723.00, result.Result);
		}

		[TestMethod]
		public void WorkdayWithDateNotAsStringAndNegDayInputReturnsPoundNum()
		{
			// Test case where the second argument is negative and the first arugment is the date not as a string.
			var function = new Workday();
			var inputDate = 1 / 2 / 2017;
			var result = function.Execute(FunctionsHelper.CreateArgs(inputDate, -10), this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WorkdayWithDateAsStringAndNegDayInputReturnsCorrectValue()
		{
			// Test case where the second argument is negative and the first argument is the date written as a string.
			var function = new Workday();
			var inputDate = "1/2/2017";
			var result = function.Execute(FunctionsHelper.CreateArgs(inputDate, -10), this.ParsingContext);
			Assert.AreEqual(42723.00, result.Result);
		}

		[TestMethod]
		public void WorkdayWithGenericStringAndNegDayInputReturnsPoundValue()
		{
			// Test case where the second argument is negative and the first argument is a non-date string or empty string. 
			var function = new Workday();
			var inputDate1 = "testString";
			var inputDate2 = "";
			var result1 = function.Execute(FunctionsHelper.CreateArgs(inputDate1, -10), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(inputDate2, -10), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result1.Result).Type);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result2.Result).Type);
		}

		[TestMethod]
		public void WorkdayWithEmptyFirstParameterAndNegDateInputREturnsPoundNA()
		{
			// Test case where the second argument is negative and the first argument is null. 
			var function = new Workday();
			var result = function.Execute(FunctionsHelper.CreateArgs(null, -10), this.ParsingContext);
			Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WorkdayWithIntFirstInputNegativeDayInputReturnsPoundNum()
		{
			// Test case where the second argument is negative and the first argument is a non-zero integer.
			var function = new Workday();
			var inputDate = 10;
			var result = function.Execute(FunctionsHelper.CreateArgs(inputDate, -10), this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WorkdayWithDateUsingPeriodsIntseadOfSlashesAndNegDayInputReturnsCorrectValue()
		{
			// Test case where the second argument is negative and the first argument is the date as a string
			// with '.' instead of '/'.
			//This functionality is different than Excell's.
			var function = new Workday();
			var inputDate = "1.2.2017";
			var result = function.Execute(FunctionsHelper.CreateArgs(inputDate, -10), this.ParsingContext);
			Assert.AreEqual(42723.00, result.Result);
		}

		[TestMethod]
		public void WorkdayWithNegFirstInputAndNegDateInputReturnsPoundNum()
		{
			// Test case where the second argument is negative and the first argument is a negative integer.
			var function = new Workday();
			var inputDate = -1;
			var result = function.Execute(FunctionsHelper.CreateArgs(inputDate, -10), this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WorkdayWithZeroFirstInputNegDateInputReturnsPoundNum()
		{
			// Test case where the second argument is negative and the first argument is 0.
			var function = new Workday();
			var result = function.Execute(FunctionsHelper.CreateArgs(0, -10), this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WorkdayWithNullSecondParameterReturnsPoundNA()
		{
			// Test case where the second argument is null.
			var function = new Workday();
			var result = function.Execute(FunctionsHelper.CreateArgs("1/2/2017", null), this.ParsingContext);
			Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WorkdayDateWithSlashesFirstInputAndNegDateInputReturnsCorrectInput()
		{
			// Test case where the second argument is negative and the first argument is a date with '-' 
			// instead of '/'.
			var function = new Workday();
			var result = function.Execute(FunctionsHelper.CreateArgs("1-2-2017", -10), this.ParsingContext);
			Assert.AreEqual(42723.00, result.Result);
		}

		// The below Test Cases involve changes only in the 'Days' Parameter.

		[TestMethod]
		public void WorkdayWithDayParameterAsDATEFunctionReturnsCorrectValue()
		{
			// Test case where the second argument is a result of the date function.
			var function = new Workday();
			var daysInput = new DateTime(2017, 1, 13);
			var startDate = new DateTime(2017, 1, 1);
			var result = function.Execute(FunctionsHelper.CreateArgs(startDate, daysInput), this.ParsingContext);
			Assert.AreEqual(102582.00, result.Result);
		}

		[TestMethod]
		public void WorkdayWithDayParameterTypedWithQuotesReturnsCorrectResult()
		{
			// Test case where the second argument is a date as a string.
			var function = new Workday();
			var daysInput = "1/13/2017";
			var startDate = new DateTime(2017, 1, 1);
			var result = function.Execute(FunctionsHelper.CreateArgs(startDate, daysInput), this.ParsingContext);
			Assert.AreEqual(102582.00, result.Result);
		}

		[TestMethod]
		public void WorkdayWithDayParameterWithNoQuotesReturnsCorrectValue()
		{
			// Test case where the second argument is a date not as a string. 
			var function = new Workday();
			var daysInput = 1 / 13 / 2017;
			var startDate = new DateTime(2017, 1, 1);
			var result = function.Execute(FunctionsHelper.CreateArgs(startDate, daysInput), this.ParsingContext);
			Assert.AreEqual(42736.00, result.Result);
		}

		[TestMethod]
		public void WorkdayWithDayParameterAsGenericStringReturnsPoundValue()
		{
			// Test case where the second argument is either a non-date string or an empty string. 
			var function = new Workday();
			var daysInput1 = "testString";
			var daysInput2 = "";
			var startDate = new DateTime(2017, 1, 1);
			var result = function.Execute(FunctionsHelper.CreateArgs(startDate, daysInput1), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(startDate, daysInput2), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result2.Result).Type);
		}

		[TestMethod]
		public void WeekdayWithNullDayParameterReturnsPoundNA()
		{
			// Test case where the second argument is null.
			var function = new Workday();
			var startDate = new DateTime(2017, 1, 1);
			var result = function.Execute(FunctionsHelper.CreateArgs(startDate, null), this.ParsingContext);
			Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WeekdayWithDayParameterAsOADateReturnCorrectValue()
		{
			// Test case where the second argument is an OADate.
			var function = new Workday();
			var startDate = new DateTime(2017, 1, 1);
			var result = function.Execute(FunctionsHelper.CreateArgs(startDate, 42748), this.ParsingContext);
			Assert.AreEqual(102582.00, result.Result);
		}

		[TestMethod]
		public void WeekdayWithDayParameterWithDotInsteadOfSlashForDateReturnsCorrectValue()
		{
			// Test case where the second argument is a date as a string with '.' instead of '/'.
			var function = new Workday();
			var startDate = new DateTime(2017, 1, 1);
			var dayInput = "1.13.2017";
			var result = function.Execute(FunctionsHelper.CreateArgs(startDate, dayInput), this.ParsingContext);
			Assert.AreEqual(102582.00, result.Result);
		}

		[TestMethod]
		public void WeekdayWithDayParameterWithDashInsteadOfSlashForDateReturnsCorrectValue()
		{
			// Test case where the second argument is the date as a string with '-' instead of '/'.
			var function = new Workday();
			var startDate = new DateTime(2017, 1, 1);
			var dayInput = "1-13-2017";
			var result = function.Execute(FunctionsHelper.CreateArgs(startDate, dayInput), this.ParsingContext);
			Assert.AreEqual(102582.00, result.Result);
		}

		[TestMethod]
		public void WeekdayWithZeroAsDateParameterReuturnsCorrectValeue()
		{
			// Test case where the second argument is zero.
			var function = new Workday();
			var startDate = new DateTime(2017, 1, 1);
			var result = function.Execute(FunctionsHelper.CreateArgs(startDate, 0), this.ParsingContext);
			Assert.AreEqual(42736.00, result.Result);
		}

		// The below Test Cases only involve the Holiday parameter.

		[TestMethod]
		public void WorkdayWithNegativeArgShouldReturnCorrectWhenArrayOfHolidayDatesIsSupplied()
		{
			// Test case where the third parameter is an OADate as an array and the second parameter is a 
			// negative integer. 
			var inputDate = new DateTime(2016, 7, 27).ToOADate();
			var holidayDate1 = new DateTime(2016, 7, 11).ToOADate();
			var holidayDate2 = new DateTime(2016, 7, 8).ToOADate();
			var expectedDate = new DateTime(2016, 6, 13).ToOADate();
			var func = new Workday();
			var args = FunctionsHelper.CreateArgs(inputDate, -30, FunctionsHelper.CreateArgs(holidayDate1, holidayDate2));
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate, result.Result);
		}

		[TestMethod]
		public void WorkdayWithNegativeArgShouldReturnCorrectWhenRangeWithHolidayDatesIsSupplied()
		{
			// Test case where the third parameter is an OADate in an Excel worksheet and the second
			// parmeter is a negative integer.
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("test");
				ws.Cells["A1"].Value = new DateTime(2016, 7, 27).ToOADate();
				ws.Cells["B1"].Value = new DateTime(2016, 7, 11).ToOADate();
				ws.Cells["B2"].Value = new DateTime(2016, 7, 8).ToOADate();
				ws.Cells["B3"].Formula = "WORKDAY(A1,-30, B1:B2)";
				ws.Calculate();
				var expectedDate = new DateTime(2016, 6, 13).ToOADate();
				var actualDate = ws.Cells["B3"].Value;
				Assert.AreEqual(expectedDate, actualDate);
			}
		}

		[TestMethod]
		public void WorkdayWithPositiveArgsAndNullHolidayDatesReturnsCorrectValue()
		{
			// Test case where the third parameter is null.
			var function = new Workday();
			var inputDate = new DateTime(2017, 1, 2);
			var result = function.Execute(FunctionsHelper.CreateArgs(inputDate, 10, null), this.ParsingContext);
			Assert.AreEqual(42751.00, result.Result);
		}

		[TestMethod]
		public void WorkdayWithStringsAsHolidayInputReturnsPoundValue()
		{
			// Test case where the third parameter is a non-date string or an empty string. 
			var function = new Workday();
			var inputDate = new DateTime(2017, 1, 2);
			var result1 = function.Execute(FunctionsHelper.CreateArgs(inputDate, 10, "testString"), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(inputDate, 10, ""), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result1.Result).Type);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result2.Result).Type);
		}

		[TestMethod]
		public void WorkdayWithNegativeIntAsHolidayInputReturnsPoundNum()
		{
			// Test case where the third parameter is a negative integer. 
			var function = new Workday();
			var inputDate = new DateTime(2017, 1, 2);
			var result = function.Execute(FunctionsHelper.CreateArgs(inputDate, 10, -1), this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WorkdayWithHolidaysAsStringsReturnsCorrectValue()
		{
			// Test case where the third parameter is a date as a string.
			var function = new Workday();
			var result = function.Execute(FunctionsHelper.CreateArgs("1/2/2017", 41, "1/25/2017"), this.ParsingContext);
			Assert.AreEqual(42795.00, result.Result);
		}

		[TestMethod]
		public void WorkdayWithHolidayWithDashesInsteadOfSlashesReturnsCorrectValue()
		{
			// Test case where the third parameter is a date as a string with '-' instead of '/'. 
			var function = new Workday();
			var inputDate = new DateTime(2017, 1, 2);
			var result = function.Execute(FunctionsHelper.CreateArgs(inputDate, 41, "1-25-2017"), this.ParsingContext);
			Assert.AreEqual(42795.00, result.Result);
		}

		[TestMethod]
		public void WorkdayWithHolidayWithDotsInsteadOfSlashesReturnsCorrectValue()
		{
			// Test case where the third parameter is a date as a string with '.' instead of '/'.
			// This functionality is different than that of Excel's. Excel does not support the date being 
			// written with periods, however many European countries write their dates in this format, so
			// EPPlus is being changed to return the correct result when this format is used.
			var function = new Workday();
			var inputDate = new DateTime(2017, 1, 2);
			var result = function.Execute(FunctionsHelper.CreateArgs(inputDate, 500, "3.30.2017"), this.ParsingContext);
			Assert.AreEqual(43438.00, result.Result);
		}

		[TestMethod]
		public void WorkdayWithZeroIntegerAsHolidayInputReturnsCorrectValue()
		{
			// Test case where the third argument is zero.
			var function = new Workday();
			var inputDate = new DateTime(2017, 1, 2);
			var result = function.Execute(FunctionsHelper.CreateArgs(inputDate, 40, 0), this.ParsingContext);
			Assert.AreEqual(42793.00, result.Result);
		}

		[TestMethod]
		public void WorkdayWithNonZeroIntegerAsHolidayInputReturnsCorrectValue()
		{
			// Test case where the third argument is a non-zero integer.
			var function = new Workday();
			var inputDate = new DateTime(2017, 1, 2);
			var result = function.Execute(FunctionsHelper.CreateArgs(inputDate, 5, 1), this.ParsingContext);
			Assert.AreEqual(42744.00, result.Result);
		}

		[TestMethod]
		public void WeekdayWithDateFunctionHolidayInputReturnsCorrectValue()
		{
			// Test  case where the third argument is a result of the DATE Function. 
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["A1"].Value = new DateTime(2017, 1, 2);
				ws.Cells["B1"].Value = new DateTime(2017, 1, 20);
				ws.Cells["B2"].Value = new DateTime(2017, 1, 25);
				ws.Cells["B3"].Formula = "WORKDAY(A1,40, B1:B2)";
				ws.Calculate();
				var actualDate = ws.Cells["B3"].Value;
				Assert.AreEqual(42795.00, actualDate);
			}
		}

		[TestMethod]
		public void WeekdayWithHolidayDateNotAsStringReturnsCorrectInput()
		{
			// Test case where the third argument is a date not as a string. It is tested in the 
			// Excel environment as well as just regularly executing the function. 
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("test");
				ws.Cells["A1"].Value = "1/2/2017";
				ws.Cells["B1"].Value = 5 / 4 / 2017;
				ws.Cells["B2"].Value = 2 / 15 / 2017;
				ws.Cells["B3"].Formula = "WORKDAY(A1,40, B1:B2)";
				ws.Calculate();
				var actualDate = ws.Cells["B3"].Value;
				Assert.AreEqual(42793.00, actualDate);
			}
			var function = new Workday();
			var holiInput = 1 / 20 / 2017;
			var result = function.Execute(FunctionsHelper.CreateArgs("1/2/2017", 40, holiInput), this.ParsingContext);
			Assert.AreEqual(42793.00, result.Result);
		}

		[TestMethod]
		public void WeekdayWithLargeNumberOfHolidaysReturnsCorrectInput()
		{
			// Test case where 30 holiday dates are supplied as strings.
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
				ws.Cells["C1"].Formula = "WORKDAY(A1, 150, B1:B41)";
				ws.Calculate();
				Assert.AreEqual(42985.00, ws.Cells["C1"].Value);
			}
		}

		[TestMethod]
		public void WeekdayWithLargeNumberOfHolidaysAsOADatesReturnsCorrectInput()
		{
			// Test case where 30 holiday dates are supplied as OADates.
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
				ws.Cells["C1"].Formula = "WORKDAY(A1, 150, B1:B41)";
				ws.Calculate();
				Assert.AreEqual(42985.00, ws.Cells["C1"].Value);
			}
		}
		#endregion
	}
}
