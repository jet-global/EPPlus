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
		public void DateFunctionReturnsADate()
		{
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2012, 4, 3);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(DataType.Date, result.DataType);
		}

		[TestMethod]
		public void DateFunctionWithNegativeMonthAndNegativeDayReturnsCorrectResult()
		{
			var expectedDate = new DateTime(2011, 10, 30);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2012, -1, -1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateFunctionWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Date();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

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
		public void DateFunctionWith1March1900ReturnsCorrectResult()
		{
			// Note that 61.0 is the Excel OADate for 3/1/1900.
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(1900, 3, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(61.0, result.Result);
		}

		[TestMethod]
		public void DateFunctionWith28February1900ReturnsCorrectResult()
		{
			// Note that 59.0 is the Excel OADate for 2/28/1900.
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(1900, 2, 28);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(59.0, result.Result);
		}

		[TestMethod]
		public void DateFunctionWith29February1900ReturnsCorrectResult()
		{
			// Note that 61.0 is the Excel OADate for 3/1/1900; since 2/29/1900 is not accepted by
			// System.DateTime as a valid date, the Date method should push the returned date up to 3/1/1900.
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(1900, 2, 29);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(61.0, result.Result);
		}

		[TestMethod]
		public void DateFunctionWithYearLessThanExcelEpochYearReturnsCorrectResult()
		{
			var expectedDate = new DateTime(3799, 1, 1);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(1899, 1, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateFunctionWithYearAs1ReturnsCorrectResult()
		{
			var expectedDate = new DateTime(1901, 1, 1);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(1, 1, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateFunctionWithYearAs0ReturnsCorrectResult()
		{
			// Note that 1.0 is the Excel OADate for 1/1/1900.
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(0, 1, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(1.0, result.Result);
		}

		[TestMethod]
		public void DateFunctionWithNegativeYearReturnsPoundNum()
		{
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(-1, 1, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void DateFunctionWithYearAs10000ReturnsPoundNum()
		{
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(10000, 1, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void DateFunctionWithYear9999ReturnsCorrectResult()
		{
			var expectedDate = new DateTime(9999, 1, 1);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(9999, 1, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateFunctionWithYearAsOneDigitReturnsCorrectResult()
		{
			var expectedDate = new DateTime(1909, 1, 1);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(9, 1, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateFunctionWithYearAsTwoDigitsReturnsCorrectResult()
		{
			var expectedDate = new DateTime(1917, 1, 1);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(17, 1, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateFunctionWithYearAsThreeDigitsReturnsCorrectResult()
		{
			var expectedDate = new DateTime(2117, 1, 1);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(217, 1, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateFunctionWithYearAsDoubleReturnsCorrectResult()
		{
			// Note that 1.0 is the Excel OADate for 1/1/1900.
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(1900.5, 1, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(1.0, result.Result);
		}

		[TestMethod]
		public void DateFunctionWithYearAsIntegerInStringReturnsCorrectResult()
		{
			// Note that 1.0 is the Excel OADate for 1/1/1900.
			var func = new Date();
			var args = FunctionsHelper.CreateArgs("1900", 1, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(1.0, result.Result);
		}

		[TestMethod]
		public void DateFunctionWithYearAsDoubleInStringReturnsCorrectResult()
		{
			// Note that 1.0 is the Excel OADate for 1/1/1900.
			var func = new Date();
			var args = FunctionsHelper.CreateArgs("1900.5", 1, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(1.0, result.Result);
		}

		[TestMethod]
		public void DateFunctionWithYearAsNonNumericStringReturnsPoundValue()
		{
			var func = new Date();
			var args = FunctionsHelper.CreateArgs("word", 1, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void DateFunctionWithYearAsEmptyStringReturnsPoundValue()
		{
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(string.Empty, 1, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void DateFunctionWithNullYearReturnsCorrectResult()
		{
			// Note that 1.0 is the Excel OADate for 1/1/1900.
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(null, 1, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(1.0, result.Result);
		}

		[TestMethod]
		public void DateFunctionWithYearAsNegativeIntegerInStringReturnsPoundNum()
		{
			var func = new Date();
			var args = FunctionsHelper.CreateArgs("-1", 1, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void DateFunctionWithMonthAs1ReturnsCorrectResult()
		{
			var expectedDate = new DateTime(2000, 1, 1);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2000, 1, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateFunctionWithMonthAs0ReturnsCorrectResult()
		{
			var expectedDate = new DateTime(1999, 12, 1);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2000, 0, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateFunctionWithMonthAsNegativeIntegerReturnsCorrectResult()
		{
			var expectedDate = new DateTime(1999, 11, 1);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2000, -1, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateFunctionWithMonthGreaterThan12ReturnsCorrectResult()
		{
			var expectedDate = new DateTime(2001, 1, 1);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2000, 13, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateFunctionWithMonthThatPushesDateBeforeMinYearReturnsPoundNum()
		{
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(0, -1, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void DateFunctionWithMonthThatPushesDateAfterMaxYearReturnsPoundNum()
		{
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(9999, 13, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void DateFunctionWithMonthThatPushesDateBeforeExcelEpochReturnsPoundNum()
		{
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(1900, -1, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void DateFunctionWithMonthThatPushesDateAfterExcelEpochDateReturnsCorrectResult()
		{
			var expectedDate = new DateTime(3800, 1, 1);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(1899, 13, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateFunctionWithMonthAsDoubleReturnsCorrectResult()
		{
			var expectedDate = new DateTime(2000, 1, 1);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2000, 1.5, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateFunctionWithMonthAsIntegerInStringReturnsCorrectResult()
		{
			var expectedDate = new DateTime(2000, 1, 1);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2000, "1", 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateFunctionWithMonthAsDoubleInStringReturnsCorrectResult()
		{
			var expectedDate = new DateTime(2000, 1, 1);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2000, "1.5", 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateFunctionWithMonthAsNonNumericStringReturnsPoundValue()
		{
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2000, "word", 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void DateFunctionWithMonthAsEmptyStringReturnsPoundValue()
		{
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2000, string.Empty, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void DateFunctionWithNullMonthReturnsCorrectResult()
		{
			var expectedDate = new DateTime(1999, 12, 1);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2000, null, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateFunctionWithDayAs0ReturnsCorrectResult()
		{
			var expectedDate = new DateTime(2000, 4, 30);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2000, 5, 0);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateFunctionWithDayAsNegativeIntegerReturnsCorrectResult()
		{
			var expectedDate = new DateTime(2000, 4, 29);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2000, 5, -1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateFunctionWithDayGreaterThan31ReturnsCorrectResult()
		{
			var expectedDate = new DateTime(2000, 6, 1);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2000, 5, 32);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateFunctionWithDayGreaterThanMonthMaxNumberOfDaysReturnsCorrectResult()
		{
			var expectedDate = new DateTime(2000, 5, 1);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2000, 4, 31);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateFunctionWithDayThatAffectsYearAndMonthReturnsCorrectResult()
		{
			var expectedDate = new DateTime(2001, 1, 1);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2000, 12, 32);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateFunctionWithDayThatPushesDateBefore1March1900ReturnsCorrectResult()
		{
			// Note that 59.0 is the Excel OADate for 2/28/1900.
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(1900, 3, -1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(59.0, result.Result);
		}

		[TestMethod]
		public void DateFunctionWithDayThatPushesDateDownIntoFebruary1900ReturnsCorrectResult()
		{
			// Note that 60.0 is the Excel OADate for 2/29/1900.
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(1900, 4, -31);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(60.0, result.Result);
		}

		[TestMethod]
		public void DateFunctionWithDayThatPushesDateUpIntoFebruary1900ReturnsCorrectResult()
		{
			// Note that 32.0 is the Excel OADate for 2/1/1900.
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(1900, 1, 32);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(32.0, result.Result);
		}

		[TestMethod]
		public void DateFunctionWithDayThatPushesDateDownPastFebruary1900ReturnsCorrectResult()
		{
			// Note that 30.0 is the Excel OADate for 1/30/1900.
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(1900, 3, -30);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(30.0, result.Result);
		}

		[TestMethod]
		public void DateFunctionWithDayThatPushesDateUpPastFebruary1900ReturnsCorrectResult()
		{
			// Note that 61.0 is the Excel OADate for 3/1/1900.
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(1900, 1, 61);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(61.0, result.Result);
		}

		[TestMethod]
		public void DateFunctionWithDayThatPushesDateBeforeExcelEpochReturnsPoundNum()
		{
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(1900, 1, -1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void DateFunctionWithDayThatPushesDateAboveMaxYearReturnsPoundNum()
		{
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(9999, 12, 32);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void DateFunctionWithDayThatPushesDateAfterExcelEpochDateReturnsCorrectResult()
		{
			var expectedDate = new DateTime(3800, 1, 1);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(1899, 12, 32);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateFunctionWithDayAsNegativeIntegerWithDateBefore1March1900ReturnsCorrectResult()
		{
			// Note that 30.0 is the Excel OADate for 1/30/1900.
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(1900, 2, -1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(30.0, result.Result);
		}

		[TestMethod]
		public void DateFunctionWithDayAsDoubleReturnsCorrectResult()
		{
			var expectedDate = new DateTime(2000, 1, 1);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2000, 1, 1.5);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateFunctionWithDayAsIntegerInStringReturnsCorrectResult()
		{
			var expectedDate = new DateTime(2000, 1, 1);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2000, 1, "1");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateFunctionWithDayAsDoubleInStringReturnsCorrectResult()
		{
			var expectedDate = new DateTime(2000, 1, 1);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2000, 1, "1.5");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateFunctionWithDayAsNonNumericStringReturnsPoundValue()
		{
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2000, 1, "word");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void DateFunctionWithDayAsEmptyStringReturnsPoundValue()
		{
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2000, 1, string.Empty);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void DateFunctionWithNullDayReturnsCorrectResult()
		{
			var expectedDate = new DateTime(1999, 12, 31);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2000, 1, null);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateFunctionWithLargeMonthAndDayValuesReturnsCorrectResult()
		{
			var expectedDate = new DateTime(2003, 1, 1);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2000, 25, 366);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateFunctionWithNegativeYearAndNonNumericStringMonthReturnsPoundValue()
		{
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(-1, "word", 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void DateFunctionWithNegativeYearAndNonNumericStringDayReturnsPoundValue()
		{
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(-1, 1, "word");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void DateFunctionWithNullYearAndNullMonthReturnsPoundNum()
		{
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(null, null, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void DateFunctionWithNullYearAndNullDayReturnsCorrectResult()
		{
			// Note that 0.0 is the Excel OADate for 1/0/1900, Excel's special 0-date.
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(null, 1, null);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(0.0, result.Result);
		}

		[TestMethod]
		public void DateFunctionWithNullMonthAndNullDayReturnsCorrectResult()
		{
			var expectedDate = new DateTime(1999, 11, 30);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2000, null, null);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateFunctionWithAllNullInputsReturnsPoundNum()
		{
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(null, null, null);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void DateFunctionWithErrorValuesAsInputReturnsTheInputErrorValue()
		{
			var func = new Date();
			var argNA = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.NA),1,1);
			var argNAME = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Name),1,1);
			var argVALUE = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Value),1,1);
			var argNUM = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Num),1,1);
			var argDIV0 = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Div0),1,1);
			var argREF = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Ref),1,1);
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
