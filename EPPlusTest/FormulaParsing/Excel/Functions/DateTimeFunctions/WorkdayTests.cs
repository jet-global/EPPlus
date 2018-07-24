using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
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
		[TestMethod]
		public void WorkdayWithOADateParameterReturnsCorrectResult()
		{
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
			var function = new Workday();
			var args = FunctionsHelper.CreateArgs();
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WorkdayWithStringInputReturnsPoundValue()
		{
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
			var function = new Workday();
			var input = 10;
			var result = function.Execute(FunctionsHelper.CreateArgs(input, 10), this.ParsingContext);
			Assert.AreEqual(24.00, result.Result);
		}

		[TestMethod]
		public void WorkdayWithDateAsStringReturnsCorrectValue()
		{
			var function = new Workday();
			var input = "1/2/2017";
			var result = function.Execute(FunctionsHelper.CreateArgs(input, 10), this.ParsingContext);
			Assert.AreEqual(42751.00, result.Result);
		}

		[TestMethod]
		public void WorkdayWithDATEFunctionAsInputReturnsCorrectValue()
		{
			var function = new Workday();
			var input = new DateTime(2017, 1, 2);
			var result = function.Execute(FunctionsHelper.CreateArgs(input, 10), this.ParsingContext);
			Assert.AreEqual(42751.00, result.Result);
		}

		[TestMethod]
		public void WorkdayWithDateUsingPeriodsInsteadOfSlashesReturnsCorrectValue()
		{
			// Test case where the first parameter is a date written as a string but with '.' in place of the '/'.
			// This functionality is different than that of Excel's. Excel normally returns a #VALUE! when this 
			// is entered into the function but here the date is parsed normally. 
			var function = new Workday();
			var input = "1.2.2017";
			var result = function.Execute(FunctionsHelper.CreateArgs(input, 10), this.ParsingContext);
			Assert.AreEqual(42751.00, result.Result);
		}

		[TestMethod]
		public void WorkdayWithNegativeFirstParamReturnsPoundNum()
		{
			var function = new Workday();
			var input = -1;
			var result = function.Execute(FunctionsHelper.CreateArgs(input, 10), this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WorkdayWithZeroInputReturnsCorrectValue()
		{
			var function = new Workday();
			var input = 0;
			var result = function.Execute(FunctionsHelper.CreateArgs(input, 15), this.ParsingContext);
			Assert.AreEqual(20.00, result.Result);
		}

		[TestMethod]
		public void WorkdayWithFractionInputReturnsCorrectValue()
		{
			var function = new Workday();
			var input = 0.5;
			var result = function.Execute(FunctionsHelper.CreateArgs(input, 15), this.ParsingContext);
			Assert.AreEqual(20.00, result.Result);
		}

		[TestMethod]
		public void WorkdayWithZeroAsStringInputReturnsCorrectValue()
		{
			var function = new Workday();
			var input = "0";
			var result = function.Execute(FunctionsHelper.CreateArgs(input, 15), this.ParsingContext);
			Assert.AreEqual(20.00, result.Result);
		}

		[TestMethod]
		public void WorkdayWithFractionAsStringInputReturnsCorrectValue()
		{
			var function = new Workday();
			var input = "0.5";
			var result = function.Execute(FunctionsHelper.CreateArgs(input, 15), this.ParsingContext);
			Assert.AreEqual(20.00, result.Result);
		}

		[TestMethod]
		public void WorkdayWithNullFirstParamReturnsPoundNum()
		{
			var function = new Workday();
			var result = function.Execute(FunctionsHelper.CreateArgs(null, 10), this.ParsingContext);
			Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WorkdayWithDateUsingDashesInsteadOfSlashesReturnsCorrectResult()
		{
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
			var function = new Workday();
			var inputDate = new DateTime(2017, 1, 2);
			var result = function.Execute(FunctionsHelper.CreateArgs(inputDate, -10), this.ParsingContext);
			Assert.AreEqual(42723.00, result.Result);
		}

		[TestMethod]
		public void WorkdayWithDateAsStringAndNegativeDayInputReturnsCorrectValue()
		{
			var function = new Workday();
			var inputDate = "1/2/2017";
			var result = function.Execute(FunctionsHelper.CreateArgs(inputDate, -10), this.ParsingContext);
			Assert.AreEqual(42723.00, result.Result);
		}

		[TestMethod]
		public void WorkdayWithGenericStringAndNegativeDayInputReturnsPoundValue()
		{
			var function = new Workday();
			var inputDate1 = "testString";
			var inputDate2 = "";
			var result1 = function.Execute(FunctionsHelper.CreateArgs(inputDate1, -10), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(inputDate2, -10), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result1.Result).Type);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result2.Result).Type);
		}

		[TestMethod]
		public void WorkdayWithEmptyFirstParameterAndNegativeDateInputReturnsPoundNA()
		{
			var function = new Workday();
			var result = function.Execute(FunctionsHelper.CreateArgs(null, -10), this.ParsingContext);
			Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WorkdayWithIntFirstInputNegativeDayInputReturnsPoundNum()
		{
			var function = new Workday();
			var inputDate = 10;
			var result = function.Execute(FunctionsHelper.CreateArgs(inputDate, -10), this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WorkdayWithDateUsingPeriodsInsteadOfSlashesAndNegativeDayInputReturnsCorrectValue()
		{
			// Test case where the second argument is negative and the first argument is the date as a string
			// with '.' instead of '/'.
			// This functionality is different than Excel's. Excel normally returns a #VALUE! when the date is written
			// this way, but here we parse the date normally (as if it were written as "1/2/2017", for example). 
			var function = new Workday();
			var inputDate = "1.2.2017";
			var result = function.Execute(FunctionsHelper.CreateArgs(inputDate, -10), this.ParsingContext);
			Assert.AreEqual(42723.00, result.Result);
		}

		[TestMethod]
		public void WorkdayWithNegativeFirstInputAndNegativeDateInputReturnsPoundNum()
		{
			var function = new Workday();
			var inputDate = -1;
			var result = function.Execute(FunctionsHelper.CreateArgs(inputDate, -10), this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WorkdayWithZeroFirstInputNegativeDateInputReturnsPoundNum()
		{
			var function = new Workday();
			var result = function.Execute(FunctionsHelper.CreateArgs(0, -10), this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WorkdayWithNullSecondParameterReturnsPoundNA()
		{
			var function = new Workday();
			var result = function.Execute(FunctionsHelper.CreateArgs("1/2/2017", null), this.ParsingContext);
			Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WorkdayDateWithSlashesFirstInputAndNegDateInputReturnsCorrectInput()
		{
			var function = new Workday();
			var result = function.Execute(FunctionsHelper.CreateArgs("1-2-2017", -10), this.ParsingContext);
			Assert.AreEqual(42723.00, result.Result);
		}

		// The below Test Cases involve changes only in the 'Days' Parameter.

		[TestMethod]
		public void WorkdayWithDayParameterAsDATEFunctionReturnsCorrectValue()
		{
			var function = new Workday();
			var daysInput = new DateTime(2017, 1, 13);
			var startDate = new DateTime(2017, 1, 1);
			var result = function.Execute(FunctionsHelper.CreateArgs(startDate, daysInput), this.ParsingContext);
			Assert.AreEqual(102582.00, result.Result);
		}

		[TestMethod]
		public void WorkdayWithDayParameterTypedWithQuotesReturnsCorrectResult()
		{
			var function = new Workday();
			var daysInput = "1/13/2017";
			var startDate = new DateTime(2017, 1, 1);
			var result = function.Execute(FunctionsHelper.CreateArgs(startDate, daysInput), this.ParsingContext);
			Assert.AreEqual(102582.00, result.Result);
		}

		[TestMethod]
		public void WorkdayWithDayParameterAsGenericStringReturnsPoundValue()
		{
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
		public void WorkdayWithNullDayParameterReturnsPoundNA()
		{
			var function = new Workday();
			var startDate = new DateTime(2017, 1, 1);
			var result = function.Execute(FunctionsHelper.CreateArgs(startDate, null), this.ParsingContext);
			Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WorkdayWithDayParameterAsOADateReturnCorrectValue()
		{
			var function = new Workday();
			var startDate = new DateTime(2017, 1, 1);
			var result = function.Execute(FunctionsHelper.CreateArgs(startDate, 42748), this.ParsingContext);
			Assert.AreEqual(102582.00, result.Result);
		}

		[TestMethod]
		public void WorkdayWithDayParameterWithDotInsteadOfSlashForDateReturnsCorrectValue()
		{
			// Test case where the second argument is the date as a string with '.' instead of '/'.
			// This functionality is different than Excel's. Excel normally returns a #VALUE! when the date is written
			// this way, but here we parse the date normally (as if it were written as "1/2/2017", for example). 
			var function = new Workday();
			var startDate = new DateTime(2017, 1, 1);
			var dayInput = "1.13.2017";
			var result = function.Execute(FunctionsHelper.CreateArgs(startDate, dayInput), this.ParsingContext);
			Assert.AreEqual(102582.00, result.Result);
		}

		[TestMethod]
		public void WorkdayWithDayParameterWithDashInsteadOfSlashForDateReturnsCorrectValue()
		{
			var function = new Workday();
			var startDate = new DateTime(2017, 1, 1);
			var dayInput = "1-13-2017";
			var result = function.Execute(FunctionsHelper.CreateArgs(startDate, dayInput), this.ParsingContext);
			Assert.AreEqual(102582.00, result.Result);
		}

		[TestMethod]
		public void WorkdayWithZeroAsDateParameterReuturnsCorrectValeue()
		{
			var function = new Workday();
			var startDate = new DateTime(2017, 1, 1);
			var result = function.Execute(FunctionsHelper.CreateArgs(startDate, 0), this.ParsingContext);
			Assert.AreEqual(42736.00, result.Result);
		}

		// The below Test Cases only involve the Holiday parameter.

		[TestMethod]
		public void WorkdayWithNegativeArgShouldReturnCorrectWhenArrayOfHolidayDatesIsSupplied()
		{
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
			var function = new Workday();
			var inputDate = new DateTime(2017, 1, 2);
			var result = function.Execute(FunctionsHelper.CreateArgs(inputDate, 10, null), this.ParsingContext);
			Assert.AreEqual(42751.00, result.Result);
		}

		[TestMethod]
		public void WorkdayWithStringsAsHolidayInputReturnsPoundValue()
		{
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
			var function = new Workday();
			var inputDate = new DateTime(2017, 1, 2);
			var result = function.Execute(FunctionsHelper.CreateArgs(inputDate, 10, -1), this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WorkdayWithHolidaysAsStringsReturnsCorrectValue()
		{
			var function = new Workday();
			var result = function.Execute(FunctionsHelper.CreateArgs("1/2/2017", 41, "1/25/2017"), this.ParsingContext);
			Assert.AreEqual(42795.00, result.Result);
		}

		[TestMethod]
		public void WorkdayWithHolidayWithDashesInsteadOfSlashesReturnsCorrectValue()
		{
			var function = new Workday();
			var inputDate = new DateTime(2017, 1, 2);
			var result = function.Execute(FunctionsHelper.CreateArgs(inputDate, 41, "1-25-2017"), this.ParsingContext);
			Assert.AreEqual(42795.00, result.Result);
		}

		[TestMethod]
		public void WorkdayWithHolidayWithDotsInsteadOfSlashesReturnsCorrectValue()
		{
			// Test case where the third parameter is a date as a string with '.' instead of '/'.
			// This functionality is different than Excel's. Excel normally returns a #VALUE! when the date is written
			// this way, but here we parse the date normally (as if it were written as "1/2/2017", for example). 
			var function = new Workday();
			var inputDate = new DateTime(2017, 1, 2);
			var result = function.Execute(FunctionsHelper.CreateArgs(inputDate, 500, "3.30.2017"), this.ParsingContext);
			Assert.AreEqual(43438.00, result.Result);
		}

		[TestMethod]
		public void WorkdayWithZeroIntegerAsHolidayInputReturnsCorrectValue()
		{
			var function = new Workday();
			var inputDate = new DateTime(2017, 1, 2);
			var result = function.Execute(FunctionsHelper.CreateArgs(inputDate, 40, 0), this.ParsingContext);
			Assert.AreEqual(42793.00, result.Result);
		}

		[TestMethod]
		public void WorkdayWithNonZeroIntegerAsHolidayInputReturnsCorrectValue()
		{
			var function = new Workday();
			var inputDate = new DateTime(2017, 1, 2);
			var result = function.Execute(FunctionsHelper.CreateArgs(inputDate, 5, 1), this.ParsingContext);
			Assert.AreEqual(42744.00, result.Result);
		}

		[TestMethod]
		public void WorkdayWithDateFunctionHolidayInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["A1"].Value = new DateTime(2017, 1, 2);
				ws.Cells["B1"].Value = new DateTime(2017, 1, 20);
				ws.Cells["B2"].Value = new DateTime(2017, 1, 25);
				ws.Cells["B3"].Formula = "WORKDAY(A1,40,B1:B2)";
				ws.Calculate();
				var actualDate = ws.Cells["B3"].Value;
				Assert.AreEqual(42795.00, actualDate);
			}
		}

		[TestMethod]
		public void WorkdayWithLargeNumberOfHolidaysReturnsCorrectInput()
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
				ws.Cells["C1"].Formula = "WORKDAY(A1, 150, B1:B41)";
				ws.Calculate();
				Assert.AreEqual(42985.00, ws.Cells["C1"].Value);
			}
		}

		[TestMethod]
		public void WorkdayWithLargeNumberOfHolidaysAsOADatesReturnsCorrectInput()
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
				ws.Cells["C1"].Formula = "WORKDAY(A1, 150, B1:B41)";
				ws.Calculate();
				Assert.AreEqual(42985.00, ws.Cells["C1"].Value);
			}
		}

		[TestMethod]
		public void WorkdayWithGermanDateAsStringWithPeriodReturnsCorrectResult()
		{
			var currentCulture = Thread.CurrentThread.CurrentCulture;
			try
			{
				Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("de-DE");
				var function = new Workday();
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
		public void WorkdayFunctionWithErrorValuesAsInputReturnsTheInputErrorValue()
		{
			var func = new Workday();
			var argNA = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.NA), 2);
			var argNAME = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Name), 2);
			var argVALUE = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Value), 2);
			var argNUM = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Num), 2);
			var argDIV0 = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Div0), 2);
			var argREF = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Ref), 2);
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
	}
}
