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
using OfficeOpenXml.Utils;

namespace EPPlusTest.FormulaParsing.Excel.Functions.DateTimeFunctions
{
	[TestClass]
	public class WeekdayTests : DateTimeFunctionsTestBase
	{
		#region Weekday Function (Execute) Tests
		[TestMethod]
		public void WeekdayWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Weekday();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WeekdayWithDateAsIntegerReturnsCorrectResult()
		{
			// Note that an omitted return_type is equivalent to using the
			// WEEKDAY function with a return_type of 1.
			var func = new Weekday();
			var args = FunctionsHelper.CreateArgs(8);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(1, result.Result);
		}

		[TestMethod]
		public void WeekdayWithDateAsDoubleReturnsCorrectResult()
		{
			var func = new Weekday();
			var args = FunctionsHelper.CreateArgs(8.5);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(1, result.Result);
		}

		[TestMethod]
		public void WeekdayWithDateAsIntegerInStringReturnsCorrectResult()
		{
			var func = new Weekday();
			var args = FunctionsHelper.CreateArgs("8");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(1, result.Result);
		}

		[TestMethod]
		public void WeekdayWithDateAsDoubleInStringReturnsCorrectResult()
		{
			var func = new Weekday();
			var args = FunctionsHelper.CreateArgs("8.5");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(1, result.Result);
		}

		[TestMethod]
		public void WeekdayWithDateAsNegativeIntegerReturnsPoundNum()
		{
			var func = new Weekday();
			var args = FunctionsHelper.CreateArgs(-1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WeekdayWithDateAsNegativeDoubleReturnsPoundNum()
		{
			var func = new Weekday();
			var args = FunctionsHelper.CreateArgs(-1.5);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WeekdayWithDateAsNonNumericStringReturnsPoundValue()
		{
			var func = new Weekday();
			var args = FunctionsHelper.CreateArgs("word");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WeekdayWithDateAsEmptyStringReturnsPoundValue()
		{
			var func = new Weekday();
			var args = FunctionsHelper.CreateArgs(string.Empty);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WeekdayWithDateAsStringReturnsCorrectResult()
		{
			var func = new Weekday();
			var args = FunctionsHelper.CreateArgs("5/17/2017");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(4, result.Result);
		}

		[TestMethod]
		public void WeekdayWithDateAsLongStringReturnsCorrectResult()
		{
			var func = new Weekday();
			var args = FunctionsHelper.CreateArgs("May 17, 2017");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(4, result.Result);
		}

		[TestMethod]
		public void WeekdayWithDateAsZeroReturnsCorrectResult()
		{
			// Note that Excel treats the OADate 0 as a special date with special output.
			var func = new Weekday();
			var args = FunctionsHelper.CreateArgs(0);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(7, result.Result);
		}

		[TestMethod]
		public void WeekdayWithDateAsFractionReturnsCorrectResult()
		{
			// Note that Excel treats the OADate 0 as a special date with special output.
			var func = new Weekday();
			var args = FunctionsHelper.CreateArgs(0.5);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(7, result.Result);
		}

		[TestMethod]
		public void WeekdayWithDateAsNegativeIntegerInStringReturnsPoundNum()
		{
			var func = new Weekday();
			var args = FunctionsHelper.CreateArgs("-1");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WeekdayWithDateAsNegativeDoubleInStringReturnsPoundNum()
		{
			var func = new Weekday();
			var args = FunctionsHelper.CreateArgs("-1.5");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WeekdayWithDateAsZeroInStringReturnsCorrectResult()
		{
			// Note that Excel treats the OADate 0 as a special date with special output.
			var func = new Weekday();
			var args = FunctionsHelper.CreateArgs("0");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(7, result.Result);
		}

		[TestMethod]
		public void WeekdayWithDateAsFractionInStringReturnsCorrectResult()
		{
			// Note that Excel treats the OADate 0 as a special date with special output.
			var func = new Weekday();
			var args = FunctionsHelper.CreateArgs("0.5");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(7, result.Result);
		}

		[TestMethod]
		public void WeekdayWithDateAs28February1900ReturnsCorrectResult()
		{
			// Note that System.DateTime dates before 3/1/1900 have their OADates off by one
			// from the Excel OADates due to Excel's inclusion of the non-existent day 2/29/1900
			// as a valid day.
			var func = new Weekday();
			var args = FunctionsHelper.CreateArgs("2/28/1900");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(3, result.Result);
		}

		[TestMethod]
		public void WeekdayWithDateAs1March1900ReturnsCorrectResult()
		{
			var func = new Weekday();
			var args = FunctionsHelper.CreateArgs("3/1/1900");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(5, result.Result);
		}

		[TestMethod]
		public void WeekdayShouldReturnCorrectResultForASundayWhenReturnTypeIs1()
		{
			// A return type of 1 requests the weekday as a 1-indexed value where
			// the week starts on Sunday.
			var func = new Weekday();
			var result = func.Execute(FunctionsHelper.CreateArgs(new DateTime(2012, 4, 1).ToOADate(), 1), this.ParsingContext);
			Assert.AreEqual(1, result.Result);
		}

		[TestMethod]
		public void WeekdayShouldReturnCorrectResultForASundayWhenReturnTypeIs2()
		{
			// A return type of 2 requests the weekday as a 1-indexed value where
			// the week starts on Monday.
			var func = new Weekday();
			var result = func.Execute(FunctionsHelper.CreateArgs(new DateTime(2012, 4, 1).ToOADate(), 2), this.ParsingContext);
			Assert.AreEqual(7, result.Result);
		}

		[TestMethod]
		public void WeekdayShouldReturnCorrectResultForASundayWhenReturnTypeIs3()
		{
			// A return type of 3 requests the weekday as a 0-indexed value where
			// the week starts on Monday.
			var func = new Weekday();
			var result = func.Execute(FunctionsHelper.CreateArgs(new DateTime(2012, 4, 1).ToOADate(), 3), this.ParsingContext);
			Assert.AreEqual(6, result.Result);
		}

		[TestMethod]
		public void WeekdayShouldReturnCorrectResultForASundayWhenReturnTypeIs11()
		{
			// A return type of 11 requests the weekday as a 1-indexed value where
			// the week starts on Monday.
			var func = new Weekday();
			var result = func.Execute(FunctionsHelper.CreateArgs(new DateTime(2012, 4, 1).ToOADate(), 11), this.ParsingContext);
			Assert.AreEqual(7, result.Result);
		}

		[TestMethod]
		public void WeekdayShouldReturnCorrectResultForASundayWhenReturnTypeIs12()
		{
			// A return type of 12 requests the weekday as a 1-indexed value where
			// the week starts on Tuesday.
			var func = new Weekday();
			var result = func.Execute(FunctionsHelper.CreateArgs(new DateTime(2012, 4, 1).ToOADate(), 12), this.ParsingContext);
			Assert.AreEqual(6, result.Result);
		}

		[TestMethod]
		public void WeekdayShouldReturnCorrectResultForASundayWhenReturnTypeIs13()
		{
			// A return type of 13 requests the weekday as a 1-indexed value where
			// the week starts on Wednesday.
			var func = new Weekday();
			var result = func.Execute(FunctionsHelper.CreateArgs(new DateTime(2012, 4, 1).ToOADate(), 13), this.ParsingContext);
			Assert.AreEqual(5, result.Result);
		}

		[TestMethod]
		public void WeekdayShouldReturnCorrectResultForASundayWhenReturnTypeIs14()
		{
			// A return type of 14 requests the weekday as a 1-indexed value where
			// the week starts on Thursday.
			var func = new Weekday();
			var result = func.Execute(FunctionsHelper.CreateArgs(new DateTime(2012, 4, 1).ToOADate(), 14), this.ParsingContext);
			Assert.AreEqual(4, result.Result);
		}

		[TestMethod]
		public void WeekdayShouldReturnCorrectResultForASundayWhenReturnTypeIs15()
		{
			// A return type of 15 requests the weekday as a 1-indexed value where
			// the week starts on Friday.
			var func = new Weekday();
			var result = func.Execute(FunctionsHelper.CreateArgs(new DateTime(2012, 4, 1).ToOADate(), 15), this.ParsingContext);
			Assert.AreEqual(3, result.Result);
		}

		[TestMethod]
		public void WeekdayShouldReturnCorrectResultForASundayWhenReturnTypeIs16()
		{
			// A return type of 16 requests the weekday as a 1-indexed value where
			// the week starts on Saturday.
			var func = new Weekday();
			var result = func.Execute(FunctionsHelper.CreateArgs(new DateTime(2012, 4, 1).ToOADate(), 16), this.ParsingContext);
			Assert.AreEqual(2, result.Result);
		}

		[TestMethod]
		public void WeekdayShouldReturnCorrectResultForASundayWhenReturnTypeIs17()
		{
			// A return type of 17 requests the weekday as a 1-indexed value where
			// the week starts on Sunday.
			var func = new Weekday();
			var result = func.Execute(FunctionsHelper.CreateArgs(new DateTime(2012, 4, 1).ToOADate(), 17), this.ParsingContext);
			Assert.AreEqual(1, result.Result);
		}

		[TestMethod]
		public void WeekdayWithReturnTypeAsIntegerInStringReturnsCorrectResult()
		{
			var func = new Weekday();
			var args = FunctionsHelper.CreateArgs(1, "1");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(1, result.Result);
		}

		[TestMethod]
		public void WeekdayWithReturnTypeAsDoubleReturnsCorrectResult()
		{
			var func = new Weekday();
			var args = FunctionsHelper.CreateArgs(1, 1.5);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(1, result.Result);
		}

		[TestMethod]
		public void WeekdayWithReturnTypeAsNonNumericStringReturnsPoundValue()
		{
			var func = new Weekday();
			var args = FunctionsHelper.CreateArgs(1, "word");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WeekdayWithReturnTypeAsEmptyStringReturnsPoundValue()
		{
			var func = new Weekday();
			var args = FunctionsHelper.CreateArgs(1, string.Empty);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WeekdayWithNonEnumeratedReturnTypeReturnsPoundNum()
		{
			// Note that the WEEKDAY function only accepts 1-3 and 11-17
			// as valid return_type values.
			var func = new Weekday();
			var args = FunctionsHelper.CreateArgs(new DateTime(2017, 5, 14), 4);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WeekdayWithNullReturnTypeReturnsPoundNum()
		{
			var func = new Weekday();
			var args = FunctionsHelper.CreateArgs(new DateTime(2017, 5, 14), null);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WeekdayWithNullDateReturnsCorrectResult()
		{
			var func = new Weekday();
			var args = FunctionsHelper.CreateArgs(null, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(7, result.Result);
		}

		[TestMethod]
		public void WeekdayWithBothParametersNullReturnsPoundNum()
		{
			var func = new Weekday();
			var args = FunctionsHelper.CreateArgs(null, null);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WeekdayWithNegativeDateAndNonNumericStringReturnTypeReturnsPoundValue()
		{
			var func = new Weekday();
			var args = FunctionsHelper.CreateArgs(-1, "word");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WeekdayNonNumericDateAndNegativeReturnTypeReturnsPoundValue()
		{
			var func = new Weekday();
			var args = FunctionsHelper.CreateArgs("word", -1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WeekdayWithDateAsStringWithPeriodsReturnsCorrectResult()
		{
			// Note that Excel does not consider "5.17.2017" as a valid date format under the
			// US culture, but System.DateTime does. In this regard, EPPlus is not completely replicating
			// Excel's handling of specifically formatted dates. Properly handling this specific case
			// is currently considered too much work for too little value.
			var func = new Weekday();
			var args = FunctionsHelper.CreateArgs("5.17.2017");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(4, result.Result);
		}

		[TestMethod]
		public void WeekdayWithDateAsStringWithCommasReturnsCorrectResult()
		{
			// Note that Excel does not consider "5,17,2017" as a valid date format under the
			// US culture, but System.DateTime does. In this regard, EPPlus is not completely replicating
			// Excel's handling of specifically formatted dates. Properly handling this specific case
			// is currently considered too much work for too little value.
			var func = new Weekday();
			var args = FunctionsHelper.CreateArgs("5,17,2017");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(4, result.Result);
		}

		[TestMethod]
		public void WeekdayWithGermanDateAsStringWithPeriodsReturnsCorrectResult()
		{
			var currentCulture = CultureInfo.CurrentCulture;
			try
			{
				var us = CultureInfo.CreateSpecificCulture("en-US");
				Thread.CurrentThread.CurrentCulture = us;
				using (var package = new ExcelPackage())
				{
					var ws = package.Workbook.Worksheets.Add("Sheet1");
					ws.Cells[2, 2].Value = "5/3/2017";
					ws.Cells[4, 3].Formula = "WEEKDAY(B2)";
					ws.Calculate();
					Assert.AreEqual(4, ws.Cells[4, 3].Value);
				}
				var gb = CultureInfo.CreateSpecificCulture("en-GB");
				Thread.CurrentThread.CurrentCulture = gb;
				using (var package = new ExcelPackage())
				{
					var ws = package.Workbook.Worksheets.Add("Sheet1");
					ws.Cells[2, 2].Value = "3/5/2017";
					ws.Cells[4, 3].Formula = "WEEKDAY(B2)";
					ws.Calculate();
					Assert.AreEqual(4, ws.Cells[4, 3].Value);
				}
				Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("de-DE");
				using (var package = new ExcelPackage())
				{
					var ws = package.Workbook.Worksheets.Add("Sheet1");
					ws.Cells[2, 2].Value = "3.5.2017";
					ws.Cells[4, 3].Formula = "WEEKDAY(B2)";
					ws.Calculate();
					Assert.AreEqual(4, ws.Cells[4, 3].Value);
				}
			}
			finally
			{
				Thread.CurrentThread.CurrentCulture = currentCulture;
			}
		}
		#endregion
	}
}
