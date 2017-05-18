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
	public class WeekdayTests : DateTimeFunctionsTestBase
	{
		#region Weekday Function (Execute) Tests
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
		public void WeekdayWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Weekday();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WeekdayWithStringArgumentReturnsPoundValue()
		{
			// Test the case where the serial_number input is a non-date String.
			var func = new Weekday();
			var args = FunctionsHelper.CreateArgs("word");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WeekdayWithNonzeroIntArgumentReturnsThatIntMod7()
		{
			// Test the case where the serial_number input is a non-zero integer.
			// Note that an omitted return_type is equivalent to using the
			// WEEKDAY function with a return_type of 1.
			var func = new Weekday();
			var args = FunctionsHelper.CreateArgs(8);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(1, result.Result);
		}

		[TestMethod]
		public void WeekdayWithArgumentAsTheNumberZeroReturnsTheNumber7()
		{
			// Test the case where the serial_number input is zero, which is
			// treated differently than other integers.
			var func = new Weekday();
			var args = FunctionsHelper.CreateArgs(0);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(7, result.Result);
		}

		[TestMethod]
		public void WeekdayWithDATEFunctionRepresentationOfSundayArgumentReturnsCorrentResult()
		{
			// Test the case where the serial_number input is provided by the
			// DATE function.
			var func = new Weekday();
			var args = FunctionsHelper.CreateArgs(new DateTime(2017, 5, 14));
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(1, result.Result);
		}

		[TestMethod]
		public void WeekdayWithInvalidReturnTypeArgumentReturnsPoundNum()
		{
			// Test the case where an invalid return_type is given.
			// Note that the WEEKDAY function only accepts 1-3 and 11-17
			// as valid return_type values.
			var func = new Weekday();
			var args = FunctionsHelper.CreateArgs(new DateTime(2017, 5, 14), 4);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WeekdayWithNegativeDateArgumentReturnsPoundNum()
		{
			// Test the case where a negative serial_number input is given.
			var func = new Weekday();
			var args = FunctionsHelper.CreateArgs(-1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WeekdayWithNonNumericStringAsSecondArgumentReturnsPoundValue()
		{
			// Test the case where a non-numeric String is given as the return_type input.
			var func = new Weekday();
			var args = FunctionsHelper.CreateArgs(new DateTime(2017, 5, 14),"word");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WeekdayWithNumericStringAsSecondArgumentReturnsCorrectResult()
		{
			// Test the case where a numeric String is given as the return_type input.
			var func = new Weekday();
			var args = FunctionsHelper.CreateArgs(1, "1");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(1, result.Result);
		}

		[TestMethod]
		public void WeekdayWithSecondArgumentNullReturnsPoundNum()
		{
			// Test the case where the serial_number and return_type parameters
			// are delineated, but the return_type is left empty.
			var func = new Weekday();
			var args = FunctionsHelper.CreateArgs(new DateTime(2017, 5, 14), null);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WeekdayWithFirstArgumentNullReturnsPoundNum()
		{
			// Test the case where the serial_number and return_type parameters
			// are delineated, but the serial_number is left empty.
			var func = new Weekday();
			var args = FunctionsHelper.CreateArgs(null, 3);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WeekdayWithDateAsStringWithSlashReturnsCorrectResult()
		{
			// Test the case where the serial_number input is expressed as
			// a date in a String.
			var func = new Weekday();
			var args = FunctionsHelper.CreateArgs("5/17/2017");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(4, result.Result);
		}

		[TestMethod]
		public void WeekdayWithDateAsStringWithDashesReturnsCorrectResult()
		{
			// Test the case where the serial_number input is expressed as
			// a date in a String.
			var func = new Weekday();
			var args = FunctionsHelper.CreateArgs("5-17-2017");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(4, result.Result);
		}

		[TestMethod]
		public void WeekdayWithDateAsStringWithPeriodsReturnsCorrectResult()
		{
			// Test the case where the serial_number input is expressed as
			// a date in a String.
			var func = new Weekday();
			var args = FunctionsHelper.CreateArgs("5.17.2017");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(4, result.Result);
		}
		#endregion
	}
}
