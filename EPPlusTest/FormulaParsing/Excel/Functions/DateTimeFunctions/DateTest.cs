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
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace EPPlusTest.FormulaParsing.Excel.Functions.DateTimeFunctions
{
	[TestClass]
	public class DateTest : DateTimeFunctionsTestBase
	{
		#region Date Function (Execute) Tests
		[TestMethod]
		public void DateFunctionShouldReturnADate()
		{
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2012, 4, 3);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(DataType.Date, result.DataType);
		}

		[TestMethod]
		public void DateFunctionShouldReturnACorrectDate()
		{
			var expectedDate = new DateTime(2012, 4, 3);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2012, 4, 3);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateFunctionShouldMonthFromPrevYearIfMonthIsNegative()
		{
			var expectedDate = new DateTime(2011, 11, 3);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2012, -1, 3);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateFunctionShouldMonthFromPrevYearIfMonthAndDayIsNegative()
		{
			var expectedDate = new DateTime(2011, 10, 30);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2012, -1, -1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Date();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		////////////////////////////

			/*
		[TestMethod]
		public void Date()
		{
			var func = new Date();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(, result.Result);
		}
		*/
		[TestMethod]
		public void DateFunctionWithNormalDateReturnsCorrectResult()
		{
			var expectedDate = new DateTime(2017, 5, 30);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2017, 5, 30);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateWith1March1900ReturnsCorrectResult()
		{
			// Note that 61.0 is the Excel OADate for 3/1/1900.
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(1900, 3, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(61.0, result.Result);
		}

		[TestMethod]
		public void DateWith28February1900ReturnsCorrectResult()
		{
			// Note that 59.0 is the Excel OADate for 2/28/1900.
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(1900, 2, 28);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(59.0, result.Result);
		}

		[TestMethod]
		public void DateWith29February1900ReturnsCorrectResult()
		{
			// Note that 61.0 is the Excel OADate for 3/1/1900; since 2/29/1900 is not accepted by
			// System.DateTime as a valid date, the Date method should push the returned date up to 3/1/1900.
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(1900, 2, 29);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(61.0, result.Result);
		}

		[TestMethod]
		public void DateWithYearLessThanExcelEpochYearReturnsCorrectResult()
		{
			var expectedDate = new DateTime(3799, 1, 1);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(1899, 1, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateWithYearAs1ReturnsCorrectResult()
		{
			var expectedDate = new DateTime(1901, 1, 1);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(1, 1, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateWithYearAs0ReturnsCorrectResult()
		{
			// Note that 1.0 is the Excel OADate for 1/1/1900.
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(0, 1, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(1.0, result.Result);
		}

		[TestMethod]
		public void DateWithNegativeYearReturnsPoundNum()
		{
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(-1, 1, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void DateWithYearAs10000ReturnsPoundNum()
		{
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(10000, 1, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void DateWithYear9999ReturnsCorrectResult()
		{
			var expectedDate = new DateTime(9999, 1, 1);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(9999, 1, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateWithYearAsOneDigitReturnsCorrectResult()
		{
			var expectedDate = new DateTime(1909, 1, 1);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(9, 1, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateWithYearAsTwoDigitsReturnsCorrectResult()
		{
			var expectedDate = new DateTime(1917, 1, 1);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(17, 1, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateWithYearAsThreeDigitsReturnsCorrectResult()
		{
			var expectedDate = new DateTime(2117, 1, 1);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(217, 1, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateWithYearAsDoubleReturnsCorrectResult()
		{
			// Note that 1.0 is the Excel OADate for 1/1/1900.
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(1900.5, 1, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(1.0, result.Result);
		}

		[TestMethod]
		public void DateWithYearAsIntegerInStringReturnsCorrectResult()
		{
			// Note that 1.0 is the Excel OADate for 1/1/1900.
			var func = new Date();
			var args = FunctionsHelper.CreateArgs("1900", 1, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(1.0, result.Result);
		}

		[TestMethod]
		public void DateWithYearAsDoubleInStringReturnsCorrectResult()
		{
			// Note that 1.0 is the Excel OADate for 1/1/1900.
			var func = new Date();
			var args = FunctionsHelper.CreateArgs("1900.5", 1, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(1.0, result.Result);
		}

		[TestMethod]
		public void DateWithYearAsNonNumericStringReturnsPoundValue()
		{
			var func = new Date();
			var args = FunctionsHelper.CreateArgs("word", 1, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void DateWithYearAsEmptyStringReturnsPoundValue()
		{
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(string.Empty);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void DateWithNullYearReturnsCorrectResult()
		{
			// Note that 1.0 is the Excel OADate for 1/1/1900.
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(null, 1, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(1.0, result.Result);
		}

		[TestMethod]
		public void DateWithYearAsNegativeIntegerInStringReturnsPoundNum()
		{
			var func = new Date();
			var args = FunctionsHelper.CreateArgs("-1", 1, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void DateWithMonthAs1ReturnsCorrectResult()
		{
			var expectedDate = new DateTime(2000, 1, 1);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2000, 1, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateWithMonthAs0ReturnsCorrectResult()
		{
			var expectedDate = new DateTime(1999, 12, 1);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2000, 0, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateWithMonthAsNegativeIntegerReturnsCorrectResult()
		{
			var expectedDate = new DateTime(1999, 12, 1);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2000, -1, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateWithMonthGreaterThan12ReturnsCorrectResult()
		{
			var expectedDate = new DateTime(2001, 1, 1);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2000, 13, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateWithMonthThatPushesDateBeforeMinYearReturnsPoundNum()
		{
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(0, -1, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void DateWithMonthThatPushesDateAfterMaxYearReturnsPoundNum()
		{
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(9999, 13, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void DateWithMonthThatPushesDateBeforeExcelEpochReturnsPoundNum()
		{
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(1900, -1, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void DateWithMonthThatPushesDateAfterExcelEpochDateReturnsCorrectResult()
		{
			var expectedDate = new DateTime(3800, 1, 1);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(1899, 13, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateWithMonthAsDoubleReturnsCorrectResult()
		{
			var expectedDate = new DateTime(2000, 1, 1);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2000, 1.5, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateWithMonthAsIntegerInStringReturnsCorrectResult()
		{
			var expectedDate = new DateTime(2000, 1, 1);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2000, "1", 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateWithMonthAsDoubleInStringReturnsCorrectResult()
		{
			var expectedDate = new DateTime(2000, 1, 1);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2000, "1.5", 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateWithMonthAsNonNumericStringReturnsPoundValue()
		{
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2000, "word", 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void DateWithMonthAsEmptyStringReturnsPoundValue()
		{
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2000, string.Empty, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void DateWithNullMonthReturnsCorrectResult()
		{
			var expectedDate = new DateTime(1999, 12, 1);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2000, null, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateWithDayAs0ReturnsCorrectResult()
		{
			var expectedDate = new DateTime(2000, 4, 30);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2000, 5, 0);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateWithDayAsNegativeIntegerReturnsCorrectResult()
		{
			var expectedDate = new DateTime(2000, 4, 29);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2000, 5, -1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateWithDayGreaterThan31ReturnsCorrectResult()
		{
			var expectedDate = new DateTime(2000, 6, 1);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2000, 5, 32);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateWithDayGreaterThanMonthMaxNumberOfDaysReturnsCorrectResult()
		{
			var expectedDate = new DateTime(2000, 5, 1);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2000, 4, 31);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateWithDayThatAffectsYearAndMonthReturnsCorrectResult()
		{
			var expectedDate = new DateTime(2001, 1, 1);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2000, 12, 32);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateWithDayThatPushesDateBefore1March1900ReturnsCorrectResult()
		{
			// Note that 59.0 is the Excel OADate for 2/28/1900.
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(1900, 3, -1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(59.0, result.Result);
		}

		[TestMethod]
		public void DateWithDayThatPushesDateBeforeExcelEpochReturnsPoundNum()
		{
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(1900, 1, -1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void DateWithDayThatPushesDateAboveMaxYearReturnsPoundNum()
		{
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(9999, 12, 32);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void DateWithDayThatPushesDateAfterExcelEpochDateReturnsCorrectResult()
		{
			var expectedDate = new DateTime(3800, 1, 1);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(1899, 12, 32);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateWithDayAsNegativeIntegerWithDateBefore1March1900ReturnsCorrectResult()
		{
			// Note that 30.0 is the Excel OADate for 1/30/1900.
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(1900, 2, -1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(30.0, result.Result);
		}

		[TestMethod]
		public void DateWithDayAsDoubleReturnsCorrectResult()
		{
			var expectedDate = new DateTime(2000, 1, 1);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2000, 1, 1.5);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateWithDayAsIntegerInStringReturnsCorrectResult()
		{
			var expectedDate = new DateTime(2000, 1, 1);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2000, 1, "1");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateWithDayAsDoubleInStringReturnsCorrectResult()
		{
			var expectedDate = new DateTime(2000, 1, 1);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2000, 1, "1.5");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateWithDayAsNonNumericStringReturnsPoundValue()
		{
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2000, 1, "word");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void DateWithDayAsEmptyStringReturnsPoundValue()
		{
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2000, 1, string.Empty);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void DateWithNullDayReturnsCorrectResult()
		{
			var expectedDate = new DateTime(1999, 12, 31);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2000, 1, null);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}


		#endregion
	}
}
