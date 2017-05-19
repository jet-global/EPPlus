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
	public class WeeknumTests : DateTimeFunctionsTestBase
	{
		#region Weeknum Function (Execute) Tests

		//The below tests do not include a second parameter (return type)

		[TestMethod]
		public void WeekNumWtihNoInputReturnsPoundValue()
		{
			//Test case where there is no input into the weeknum function.
			var func = new Weeknum();

			var r1 = func.Execute(FunctionsHelper.CreateArgs(), this.ParsingContext);

			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)r1.Result).Type);
		}

		[TestMethod]
		public void WeekNumWithDateFunctionInputReturnsCorrectResult()
		{
			//Test case where the input is a DateTime from the DateTime function.
			var func = new Weeknum();

			var dt = new DateTime(2017, 1, 5);

			var r1 = func.Execute(FunctionsHelper.CreateArgs(dt), this.ParsingContext);

			Assert.AreEqual(1, r1.Result);
		}

		[TestMethod]
		public void WeekNumWithDateAsStringReturnsCorrectResult()
		{
			//Test case where the input is a date written as a string.
			var func = new Weeknum();

			var dt = "1/10/2017";

			var r1 = func.Execute(FunctionsHelper.CreateArgs(dt), this.ParsingContext);

			Assert.AreEqual(2, r1.Result);
		}

		[TestMethod]
		public void WeekNumWithDateAsStringWithDashesReturnsCorrectResult()
		{
			// Test case where the input is a date is written with '-' instead of '/'.
		
			var func = new Weeknum();

			var dt = "1-10-2017";

			var r1 = func.Execute(FunctionsHelper.CreateArgs(dt), this.ParsingContext);

			Assert.AreEqual(2, r1.Result);
		}

		[TestMethod]
		public void WeekNumWithStringArgumentReturnsPoundValue()
		{
			// Test case where the input is a general string and empty string.
			var func = new Weeknum();

			var dt1 = "testString";
			var dt2 = "";

			var r1 = func.Execute(FunctionsHelper.CreateArgs(dt1), this.ParsingContext);
			var r2 = func.Execute(FunctionsHelper.CreateArgs(dt2), this.ParsingContext);

			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)r1.Result).Type);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)r2.Result).Type);
		}

		[TestMethod]
		public void WeekNumWithOADateArgumentReturnsCorrectValue()
		{
			//Test case with the input as an OADate.
			var func = new Weeknum();

			var dt = new DateTime(2017, 1, 10).ToOADate();

			var r1 = func.Execute(FunctionsHelper.CreateArgs(dt), this.ParsingContext);

			Assert.AreEqual(2, r1.Result);
		}

		[TestMethod]
		public void WeekNumWithNonZeroIntegerArgumentReturnsCorrectValue()
		{
			//Test case where the input is a non-zero integer.
			var func = new Weeknum();

			var dt = 365;

			var r1 = func.Execute(FunctionsHelper.CreateArgs(dt), this.ParsingContext);

			Assert.AreEqual(53, r1.Result);
		}

		[TestMethod]
		public void WeekNumWithZeroIntegerReturnsZero()
		{
			//Test case where the input is zero.
			var func = new Weeknum();

			var dt = 0;

			var r1 = func.Execute(FunctionsHelper.CreateArgs(dt), this.ParsingContext);

			Assert.AreEqual(0, r1.Result);
		}

		[TestMethod]
		public void WeekNumWithNegativeIntReturnsPoundNum()
		{
			// Test case where the input is a negative integer.
			var func = new Weeknum();

			var dt = -5;

			var r1 = func.Execute(FunctionsHelper.CreateArgs(dt), this.ParsingContext);

			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)r1.Result).Type);
		}

		[TestMethod]
		public void WeekNumWithDateNotAsStringReturnsCorrectValue()
		{
			// Test case where the date is written not in string form.
			var func = new Weeknum();

			var dt = 1 / 5 / 2017;

			var r1 = func.Execute(FunctionsHelper.CreateArgs(dt), this.ParsingContext);

			Assert.AreEqual(0, r1.Result);
		}

		[TestMethod]
		public void WeekNumWithPeriodInsteadOfDashesOnUSCultureCorrecctValue()
		{
			// Test case where the input is the date written with '.' instead of '/'.
			//This functionality differs from Excel's. Excel normally returns a #VALUE! on the US 
			//Culture, however the Weeknum class in EPPlus returns the week number. 
			var func = new Weeknum();

			var dt = "1.5.2017";

			var r1 = func.Execute(FunctionsHelper.CreateArgs(dt), this.ParsingContext);

			Assert.AreEqual(1, r1.Result);
		}

		[TestMethod]
		public void WeekNumWithItalianhDateAsStringWithPeriodsReturnsCorrectResult()
		{
			// Test case where the culture has been changed to Italian to test the 
			// Italian date format.
			var currentCulture = Thread.CurrentThread.CurrentCulture;
			try
			{
				Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("it-IT");
				var func = new Weekday();
				var args = FunctionsHelper.CreateArgs("3.5.2017");
				var result = func.Execute(args, this.ParsingContext);
				Assert.AreEqual(4, result.Result);
			}
			finally
			{
				Thread.CurrentThread.CurrentCulture = currentCulture;
			}
		}

		//Below are the tests that include the second parameter (return type)

		[TestMethod]
		public void WeekNumWithReturnType1OrOmmittedReturnsCorrectValue()
		{
			// Test case with the a return type of 1 or an ommitted return type.
			var func = new Weeknum();

			var dt = new DateTime(2017, 1, 5);

			var r1 = func.Execute(FunctionsHelper.CreateArgs(dt, 1), this.ParsingContext);
			var r2 = func.Execute(FunctionsHelper.CreateArgs(dt), this.ParsingContext);

			Assert.AreEqual(1, r1.Result);
			Assert.AreEqual(1, r2.Result);
		}

		[TestMethod]
		public void WeekNumWithReturnType2ReturnsCorrectValue()
		{
			// Test case where the return type is 2.
			var func = new Weeknum();

			var dt = new DateTime(2017, 1, 5);

			var r1 = func.Execute(FunctionsHelper.CreateArgs(dt, 2), this.ParsingContext);

			Assert.AreEqual(2, r1.Result);
		}

		[TestMethod]
		public void WeekNumWithReturnType11ReturnsCorrectValue()
		{
			// Test case where the return type is 11.
			var func = new Weeknum();

			var dt = new DateTime(2017, 1, 5);

			var r1 = func.Execute(FunctionsHelper.CreateArgs(dt, 11), this.ParsingContext);

			Assert.AreEqual(2, r1.Result);
		}

		[TestMethod]
		public void WeekNumWithReturnType12ReturnsCorrectValue()
		{
			// Test case where the return type is 12.
			var func = new Weeknum();

			var dt = new DateTime(2017, 1, 5);

			var r1 = func.Execute(FunctionsHelper.CreateArgs(dt, 12), this.ParsingContext);

			Assert.AreEqual(2, r1.Result);
		}

		[TestMethod]
		public void WeekNumWithReturnType13ReturnsCorrectValue()
		{
			//Test case where the reutrn type is 13.
			var func = new Weeknum();

			var dt = new DateTime(2017, 1, 5);

			var r1 = func.Execute(FunctionsHelper.CreateArgs(dt, 13), this.ParsingContext);

			Assert.AreEqual(2, r1.Result);
		}

		[TestMethod]
		public void WeekNumWithReturnType14ReturnsCorrectValue()
		{
			// Test case where the return type is 14.
			var func = new Weeknum();

			var dt = new DateTime(2017, 1, 5);

			var r1 = func.Execute(FunctionsHelper.CreateArgs(dt, 14), this.ParsingContext);

			Assert.AreEqual(2, r1.Result);
		}

		[TestMethod]
		public void WeekNumWithReturnType15ReturnsCorrectValue()
		{
			// Test case where the return type is 15.
			var func = new Weeknum();

			var dt = new DateTime(2017, 1, 5);

			var r1 = func.Execute(FunctionsHelper.CreateArgs(dt, 15), this.ParsingContext);

			Assert.AreEqual(1, r1.Result);
		}

		[TestMethod]
		public void WeekNumWithReturnType16ReturnsCorrectValue()
		{
			// Test case where the return type is 16.
			var func = new Weeknum();

			var dt = new DateTime(2017, 1, 5);

			var r1 = func.Execute(FunctionsHelper.CreateArgs(dt, 16), this.ParsingContext);

			Assert.AreEqual(1, r1.Result);
		}

		[TestMethod]
		public void WeekNumWithReturnType17ReturnsCorrectValue()
		{
			// Test case where the return type is 17.
			var func = new Weeknum();

			var dt = new DateTime(2017, 1, 5);

			var r1 = func.Execute(FunctionsHelper.CreateArgs(dt, 17), this.ParsingContext);

			Assert.AreEqual(1, r1.Result);
		}

		[TestMethod]
		public void WeekNumWithReturnType21ReturnsCorrectValue()
		{
			// Test case where the return type is 21.
			var func = new Weeknum();

			var dt = new DateTime(2017, 1, 5);

			var r1 = func.Execute(FunctionsHelper.CreateArgs(dt, 21), this.ParsingContext);

			Assert.AreEqual(1, r1.Result);
		}

		[TestMethod]
		public void WeekNumWithInvalidReturnTypeReturnsPoundNum()
		{
			// Test case with an invalid return type.
			var func = new Weeknum();

			var dt = new DateTime(2017, 1, 5);

			var r1 = func.Execute(FunctionsHelper.CreateArgs(dt, 5), this.ParsingContext);

			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)r1.Result).Type);
		}

		[TestMethod]
		public void WeekNumWithNumericStringReturnTypeReturnsCorrectValue()
		{
			// Test case where the return type is in string format.
			var func = new Weeknum();

			var dt = new DateTime(2017, 1, 5);

			var r1 = func.Execute(FunctionsHelper.CreateArgs(dt, "1"), this.ParsingContext);

			Assert.AreEqual(1, r1.Result);
		}

		[TestMethod]
		public void WeekNumWithStringReturnTypeReturnsPoundValue()
		{
			// Test case where the return type is a general string
			var func = new Weeknum();

			var dt = new DateTime(2017, 1, 5);

			var r1 = func.Execute(FunctionsHelper.CreateArgs(dt, "testString"), this.ParsingContext);

			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)r1.Result).Type);
		}

		[TestMethod]
		public void WeekNumWithNoFirstParameterAndValidReturnTypeReturnsPoundNA()
		{
			// Test case where the first param is null with a valid return type.
			var func = new Weeknum();

			var r1 = func.Execute(FunctionsHelper.CreateArgs(null, 1), this.ParsingContext);

			Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)r1.Result).Type);
		}

		[TestMethod]
		public void WeekNumWithNoFirstParameterAndInvalidReturnTypeReturnsPoundNA()
		{
			// Test case where the first param is null with an invalid return type. 
			var func = new Weeknum();

			var r1 = func.Execute(FunctionsHelper.CreateArgs(null, 5), this.ParsingContext);

			Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)r1.Result).Type);
		}

		[TestMethod]
		public void WeekNumWithNegativeIntAndValidReturnTypeReutrnsPoundNum()
		{
			// Test case where the first param is a negative integer and has a valid
			// return type.
			var func = new Weeknum();

			var r1 = func.Execute(FunctionsHelper.CreateArgs(-1, 1), this.ParsingContext);

			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)r1.Result).Type);
		}

		[TestMethod]
		public void WeekNumWithNegativIntAndInvalidReturnTypeReturnsPoundNum()
		{
			// Test case where the first param is a negative integer and has an invalid
			// return type. 
			var func = new Weeknum();

			var r1 = func.Execute(FunctionsHelper.CreateArgs(-1, 5), this.ParsingContext);

			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)r1.Result).Type);
		}
		#endregion
	}
}
