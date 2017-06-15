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
	public class DayTests : DateTimeFunctionsTestBase
	{
		#region Day Function (Execute) Tests
		[TestMethod]
		public void DayWithDateAsStringReturnsDayOfMonth()
		{
			// Test the case where the date is entered in a date
			// format in a string.
			var func = new Day();
			var args = FunctionsHelper.CreateArgs("2012-03-12");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(12, result.Result);
		}

		[TestMethod]
		public void DayWithDateAsLongStringReturnsDayOfMonth()
		{
			// Test the case where the date is entered in a different
			// date format in a string.
			var func = new Day();
			var args = FunctionsHelper.CreateArgs("May 19, 2017");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(19, result.Result);
		}

		[TestMethod]
		public void DayWithInvalidArgumentReturnsPoundValue()
		{
			// Test the case where nothing is entered in the DAY function.
			var func = new Day();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void DayFunctionWorksInDifferentCultureDateFormats()
		{
			// Test the case where the date is represented under
			// different cultures.
			var currentCulture = CultureInfo.CurrentCulture;
			try
			{
				var us = CultureInfo.CreateSpecificCulture("en-US");
				Thread.CurrentThread.CurrentCulture = us;
				using (var package = new ExcelPackage())
				{
					var ws = package.Workbook.Worksheets.Add("Sheet1");
					ws.Cells[2, 2].Value = "1/15/2014";
					ws.Cells[5, 3].Formula = "DAY(B2)";
					ws.Calculate();
					Assert.AreEqual(15, ws.Cells[5, 3].Value);
				}
				var gb = CultureInfo.CreateSpecificCulture("en-GB");
				Thread.CurrentThread.CurrentCulture = gb;
				using (var package = new ExcelPackage())
				{
					var ws = package.Workbook.Worksheets.Add("Sheet1");
					ws.Cells[2, 2].Value = "15/1/2014";
					ws.Cells[5, 3].Formula = "DAY(B2)";
					ws.Calculate();
					Assert.AreEqual(15, ws.Cells[5, 3].Value);
				}
				var de = CultureInfo.CreateSpecificCulture("de-DE");
				Thread.CurrentThread.CurrentCulture = de;
				using (var package = new ExcelPackage())
				{
					var ws = package.Workbook.Worksheets.Add("Sheet1");
					ws.Cells[2, 2].Value = "15.1.2014";
					ws.Cells[5, 3].Formula = "DAY(B2)";
					ws.Calculate();
					Assert.AreEqual(15, ws.Cells[5, 3].Value);
				}
			}
			finally
			{
				Thread.CurrentThread.CurrentCulture = currentCulture;
			}
		}

		[TestMethod]
		public void DayWithDateAsIntegerReturnsDayInMonth()
		{
			// Test the case where the date is entered as an integer.
			// Note that 42874 is the OADate for May 19, 2017.
			var func = new Day();
			var args = FunctionsHelper.CreateArgs(42874);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(19, result.Result);
		}

		[TestMethod]
		public void DayWithDateAsDoubleReturnsDayInMonth()
		{
			// Test the case where the date is entered as a double.
			// Note that 42874.34114 is the OADate representation of
			// some time on May 19, 2017.
			var func = new Day();
			var args = FunctionsHelper.CreateArgs(42874.34114);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(19, result.Result);
		}

		[TestMethod]
		public void DayWithZeroAsInputReturnsZeroAsDateInMonth()
		{
			// Test the case where zero is the input.
			var func = new Day();
			var args = FunctionsHelper.CreateArgs(0);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(0, result.Result);
		}

		[TestMethod]
		public void DayWithNegativeIntegerAsInputReturnsPoundNum()
		{
			// Test the case where a negative integer is the input.
			var func = new Day();
			var args = FunctionsHelper.CreateArgs(-1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void DayWithNegativeDoubleAsInputReturnsPoundNum()
		{
			// Test the case where a negative double is the input
			var func = new Day();
			var args = FunctionsHelper.CreateArgs(-1.5);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void DayWithNonDateStringAsInputReturnsPoundValue()
		{
			// Test the case where a non-date string is the input.
			var func = new Day();
			var args = FunctionsHelper.CreateArgs("word");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void DayWithEmptyStringAsInputReturnsPoundValue()
		{
			// Test the case where an empty string is the input.
			var func = new Day();
			var args = FunctionsHelper.CreateArgs("");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void DayWithIntegerInStringAsInputReturnsDayInMonth()
		{
			// Test the case where the input is an integer expressed as a string.
			// Note that 42874 is the OADate representation of May 19, 2017.
			var func = new Day();
			var args = FunctionsHelper.CreateArgs("42874");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(19, result.Result);
		}

		[TestMethod]
		public void DayWithDoubleInStringAsInputReturnsDayInMonth()
		{
			// Test the case where the input is a double expressed as a string.
			// Note that 42874.34114 is the OADate representation of some time
			// on May 19, 2017.
			var func = new Day();
			var args = FunctionsHelper.CreateArgs("42874.43114");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(19, result.Result);
		}

		[TestMethod]
		public void DayWithNegativeIntegerInStringAsInputReturnsPoundNum()
		{
			// Test the case where the input is a negative integer expressed
			// as a string.
			var func = new Day();
			var args = FunctionsHelper.CreateArgs("-1");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void DayWithNegativeDoubleInStringAsInputReturnsPoundNum()
		{
			// Test the case where the input is a negative double expressed
			// as a string.
			var func = new Day();
			var args = FunctionsHelper.CreateArgs("-1.5");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void DayWithDateTimeObjectReturnsCorrectResult()
		{
			// Test the case where a DateTime object is given as the input.
			var date = new DateTime(2017, 5, 22);
			var func = new Day();
			var args = FunctionsHelper.CreateArgs(date);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(22, result.Result);
		}

		[TestMethod]
		public void DayWithDateTimeObjectForOADateLessThan61ReturnsCorrectResult()
		{
			// Test the case where a DateTime object representing a date before
			// March 1, 1900 is given as the input.
			var date = new DateTime(1900, 2, 28);
			var func = new Day();
			var args = FunctionsHelper.CreateArgs(date);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(28, result.Result);
		}

		[TestMethod]
		public void DayHandlesExcelOffByOneErrorFor27February1900()
		{
			// This test exists as a reminder that, for any date before March 1, 1900 (which has OADate == 61),
			// Excel and the C# Library System.DateTime have different OADates to represent that date.
			var date = new DateTime(1900, 2, 27); // The date being represented, 2/27/1900.
			Assert.AreEqual(59, date.ToOADate()); // The OADate from System.DateTime.
			var func = new Day();
			var args = FunctionsHelper.CreateArgs(58); // The OADate from Excel for 2/27/1900.
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(27, result.Result); // The day calculated using Excel's OADate, corresponding to 2/27/1900.
		}

		[TestMethod]
		public void DayHandlesExcelOffByOneErrorFor28February1900()
		{
			// This test exists as a reminder that, for any date before March 1, 1900 (which has OADate == 61),
			// Excel and the C# Library System.DateTime have different OADates to represent that date.
			var date = new DateTime(1900, 2, 28); // The date being represented, 2/28/1900.
			Assert.AreEqual(60, date.ToOADate()); // The OADate from System.DateTime.
			var func = new Day();
			var args = FunctionsHelper.CreateArgs(59); // The OADate from Excel for 2/28/1900.
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(28, result.Result); // The day calculated using Excel's OADate, corresponding to 2/28/1900.
		}

		[TestMethod]
		public void DayTreatsNonExistentDateOf29February1900As1March1900()
		{
			// This test exists as a reminder that, since the C# Library System.DateTime does not accept 2/29/1900 as a valid date
			// (because that day doesn't exist), the OADate for 2/29/1900 from Excel (which Excel considers a valid date)
			// is instead mapped to 3/1/1900 when using the Day function. So using the OADates 60 and 61 in the Day function
			// will both calculate the day using 3/1/1900 as the date.
			var func = new Day();
			var args = FunctionsHelper.CreateArgs(60); // The OADate from Excel for 2/29/1900, which doesnt exist in System.DateTime, so also the OADate for 3/1/1900.
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(1, result.Result); // The day calculated using Excel's OADate, corresponding to 3/1/1900.
		}

		[TestMethod]
		public void DayHandlesExcelOffByOneErrorFor1March1900()
		{
			// This test exists as a reminder that March 1, 1900 is the date where Excel
			// and the C# Library System.DateTime's OADates sync back up; 61 is the OADate for March 1, 1900 in
			// both Excel and EPPlus.
			var date = new DateTime(1900, 3, 1); // The date being represented, 3/1/1900.
			Assert.AreEqual(61, date.ToOADate()); // The OADate from System.DateTime.
			var func = new Day();
			var args = FunctionsHelper.CreateArgs(61); // The OADate from Excel for 3/1/1900.
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(1, result.Result); // The date calculated using Excel's OADate, corresponding to 3/1/1900.
		}

		[TestMethod]
		public void DayWithInputAsDoubleNearMarch1ReturnsCorrectResult()
		{
			// Test the case where a time and day close to 3/1/1900 returns
			// the correct result. Note that Excel represents 2/28/1900 as the
			// OADate 59.
			var func = new Day();
			var args = FunctionsHelper.CreateArgs(59.99999);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(28, result.Result);
		}

		[TestMethod]
		public void DayWithOADate1ReturnsCorrectResult()
		{
			// OADate 1 corresponds to 1/1/1900 in Excel. Test the case where Excel's epoch date
			// is used as the input.
			var func = new Day();
			var args = FunctionsHelper.CreateArgs(1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(1, result.Result);
		}

		[TestMethod]
		public void DayWithOADateBetween0And1ReturnsZero()
		{
			// Test the case where a time value is given for Excel's
			// "zeroeth" day, which Excel treats as a special day.
			var func = new Day();
			var args = FunctionsHelper.CreateArgs(0.5);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(0, result.Result);
		}

		[TestMethod]
		public void DayWithNegativeOADateBetweenNegative1And0ReturnsPoundNum()
		{
			// Test the case where a negative time value is given for Excel's
			// "zeroeth" day, which Excel treats as a special day.
			var func = new Day();
			var args = FunctionsHelper.CreateArgs(-0.5);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void DayWithDateAs0InStringReturnsCorrectResult()
		{
			var func = new Day();
			var args = FunctionsHelper.CreateArgs("0");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(0, result.Result);
		}

		[TestMethod]
		public void DayWithFractionalDateInStringReturnsCorrectResult()
		{
			var func = new Day();
			var args = FunctionsHelper.CreateArgs("0.5");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(0, result.Result);
		}

		[TestMethod]
		public void DayFunctionWithErrorValuesAsInputReturnsTheInputErrorValue()
		{
			var func = new Day();
			var argNA = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.NA));
			var argNAME = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Name));
			var argVALUE = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Value));
			var argNUM = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Num));
			var argDIV0 = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Div0));
			var argREF = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Ref));
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
		#endregion
	}
}
