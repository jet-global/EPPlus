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
	public class YearFracTests : DateTimeFunctionsTestBase
	{
		#region YearFrac Function (Execute) Tests
		[TestMethod]
		public void YearFracWithTooFewArgumentsReturnsPoundValue()
		{
			var func = new Yearfrac();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, (result.Result as ExcelErrorValue).Type);
		}

		[TestMethod]
		public void YearFracWithDateAsIntegerReturnsCorrectResult()
		{
			var date1 = new DateTime(2017, 1, 1);
			var date2 = new DateTime(2017, 5, 23);
			var func = new Yearfrac();
			var args = FunctionsHelper.CreateArgs((int)(date1.ToOADate()), (int)(date2.ToOADate()));
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(0.39444444444, result.Result);
		}

		[TestMethod]
		public void YearFracWithDateAsDoubleReturnsCorrectResult()
		{
			var date1 = new DateTime(2017, 1, 1);
			var date2 = new DateTime(2017, 5, 23);
			var func = new Yearfrac();
			var args = FunctionsHelper.CreateArgs(date1.ToOADate(), date2.ToOADate());
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(0.39444444444, result.Result);
		}

		[TestMethod]
		public void YearFracWithDateAsIntegerInStringReturnsCorrectResult()
		{
			// Note that 42736 and 42878 are the Excel OADates for 1/1/2017 and 5/23/2017 respectively.
			var func = new Yearfrac();
			var args = FunctionsHelper.CreateArgs("42736", "42878");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(0.39444444444, result.Result);
		}

		[TestMethod]
		public void YearFracWithDateAsDoubleInStringReturnsCorrectResult()
		{
			// Note that 42736.5 and 42878.5 are the Excel OADates for noon on 1/1/2017 and 5/23/2017 respectively.
			var func = new Yearfrac();
			var args = FunctionsHelper.CreateArgs("42736.5", "42878.5");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(0.39444444444, result.Result);
		}

		[TestMethod]
		public void YearFracWithDateAsStringReturnsCorrectResult()
		{
			var func = new Yearfrac();
			var args = FunctionsHelper.CreateArgs("1/1/2017", "5/23/2017");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(0.39444444444, result.Result);
		}

		[TestMethod]
		public void YearFracWithDateAsLongStringReturnsCorrectResult()
		{
			var func = new Yearfrac();
			var args = FunctionsHelper.CreateArgs("January 1, 2017", "May 23, 2017");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(0.39444444444, result.Result);
		}

		[TestMethod]
		public void YearFracWithNonNumericStringAsFirstParameterReturnsPoundValue()
		{
			// Note that 42878 is the Excel OADate for 5/23/2017.
			var func = new Yearfrac();
			var args = FunctionsHelper.CreateArgs("word", 42878);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void YearFracWithNonNumericStringAsSecondParameterReturnsPoundValue()
		{
			// Note that 42736 is the Excel OADate for 1/1/2017.
			var func = new Yearfrac();
			var args = FunctionsHelper.CreateArgs(42736, "word");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void YearFracWithEmptyStringAsFirstParameterReturnsPoundValue()
		{
			// Note that 42878 is the Excel OADate for 5/23/2017.
			var func = new Yearfrac();
			var args = FunctionsHelper.CreateArgs(string.Empty, 42878);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void YearFracWithEmptyStringAsSecondParameterReturnsPoundValue()
		{
			// Note that 42736 is the Excel OADate for 1/1/2017.
			var func = new Yearfrac();
			var args = FunctionsHelper.CreateArgs(42736, string.Empty);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void YearFracWithDateAsNegativeIntegerReturnsPoundNum()
		{
			// Note that 42878 is the Excel OADate for 5/23/2017.
			var func = new Yearfrac();
			var args = FunctionsHelper.CreateArgs(-1, 42878);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void YearFracWithDateAsNegativeDoubleReturnsPoundNum()
		{
			// Note that 42736 is the Excel OADate for 1/1/2017.
			var func = new Yearfrac();
			var args = FunctionsHelper.CreateArgs(42736, -1.5);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void YearFracWithDateAsNegativeIntegerInStringReturnsPoundNum()
		{
			// Note that 42878 is the Excel OADate for 5/23/2017.
			var func = new Yearfrac();
			var args = FunctionsHelper.CreateArgs("-1", 42878);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void YearFracWithDateAsNegativeDoubleInStringReturnsPoundNum()
		{
			// Note that 42736 is the Excel OADate for 1/1/2017.
			var func = new Yearfrac();
			var args = FunctionsHelper.CreateArgs(42736, "-1.5");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void YearFracWithFirstDateLaterThanSecondDateReturnsCorrectResult()
		{
			var date1 = new DateTime(2017, 5, 23);
			var date2 = new DateTime(2017, 1, 1);
			var func = new Yearfrac();
			var args = FunctionsHelper.CreateArgs(date1.ToOADate(), date2.ToOADate());
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(0.39444444444, result.Result);
		}

		[TestMethod]
		public void YearFracWithDatesMoreThanAYearApartReturnsCorrectResult()
		{
			var date1 = new DateTime(2017, 1, 1);
			var date2 = new DateTime(2019, 5, 23);
			var func = new Yearfrac();
			var args = FunctionsHelper.CreateArgs(date1.ToOADate(), date2.ToOADate());
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(2.39444444444, result.Result);
		}

		[TestMethod]
		public void YearFracWithNullFirstParameterReturnsPoundNA()
		{
			// Note that 42878 is the Excel OADate for 5/23/2017.
			var func = new Yearfrac();
			var args = FunctionsHelper.CreateArgs(null, 42878);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void YearFracWithNullSecondParameterReturnsPoundNA()
		{
			// Note that 42736 is Excel OADate for 1/1/2017.
			var func = new Yearfrac();
			var args = FunctionsHelper.CreateArgs(42736, null);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void YearFracWithNullBasisParameterReturnsCorrectResult()
		{
			var date1 = new DateTime(2017, 1, 1);
			var date2 = new DateTime(2017, 5, 23);
			var func = new Yearfrac();
			var args = FunctionsHelper.CreateArgs(date1.ToOADate(), date2.ToOADate(), null);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(0.39444444444, result.Result);
		}

		[TestMethod]
		public void YearFracWithBothDatesTheSameReturnsCorrectResult()
		{
			var date = new DateTime(2017, 1, 1);
			var func = new Yearfrac();
			var args = FunctionsHelper.CreateArgs(date.ToOADate(), date.ToOADate());
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(0.0, result.Result);
		}

		[TestMethod]
		public void YearFracWithDateAsZeroReturnsCorrectResult()
		{
			// Note that 42878 is the Excel OADate for 5/23/2017.
			// 0 is the Excel OADate for 1/0/1900, which Excel treats as a special date.
			var func = new Yearfrac();
			var args = FunctionsHelper.CreateArgs(0, 42878);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(117.39722222222, result.Result);
		}

		[TestMethod]
		public void YearFracWithDateAsOneReturnsCorrectResult()
		{
			// Note that 1 and 42878 are the Excel OADates for 1/1/1900 and 5/23/2017 respectively.
			var func = new Yearfrac();
			var args = FunctionsHelper.CreateArgs(1, 42878);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(117.39444444444, result.Result);
		}

		[TestMethod]
		public void YearFracWithNegativeFirstParameterAndNonNumericStringSecondParameterReturnsPoundNum()
		{
			var func = new Yearfrac();
			var args = FunctionsHelper.CreateArgs(-1, "word");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void YearFracWithNonNumericStringFirstParameterAndNegativeSecondParameterReturnsPoundValue()
		{
			var func = new Yearfrac();
			var args = FunctionsHelper.CreateArgs("word", -1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void YearFracWithDatesAfter1March1900WithBasis1ReturnsCorrectResult()
		{
			var date1 = new DateTime(2017, 1, 1);
			var date2 = new DateTime(2017, 5, 23);
			var func = new Yearfrac();
			var args = FunctionsHelper.CreateArgs(date1.ToOADate(), date2.ToOADate(),1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(0.38904109589, result.Result);
		}

		[TestMethod]
		public void YearFracWithDatesAfter1March1900WithBasis2ReturnsCorrectResult()
		{
			var date1 = new DateTime(2017, 1, 1);
			var date2 = new DateTime(2017, 5, 23);
			var func = new Yearfrac();
			var args = FunctionsHelper.CreateArgs(date1.ToOADate(), date2.ToOADate(), 2);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(0.39444444444, result.Result);
		}

		[TestMethod]
		public void YearFracWithDatesAfter1March1900WithBasis3ReturnsCorrectResult()
		{
			var date1 = new DateTime(2017, 1, 1);
			var date2 = new DateTime(2017, 5, 23);
			var func = new Yearfrac();
			var args = FunctionsHelper.CreateArgs(date1.ToOADate(), date2.ToOADate(), 3);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(0.38904109589, result.Result);
		}

		[TestMethod]
		public void YearFracWithDatesAfter1March1900WithBasis4ReturnsCorrectResult()
		{
			var date1 = new DateTime(2017, 1, 1);
			var date2 = new DateTime(2017, 5, 23);
			var func = new Yearfrac();
			var args = FunctionsHelper.CreateArgs(date1.ToOADate(), date2.ToOADate(), 4);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(0.39444444444, result.Result);
		}

		[TestMethod]
		public void YearFracWithDatesBefore1March1900WithBasis0ReturnsCorrectResult()
		{
			// Note that 31 and 59 are the Excel OADates for 1/31/1900 and 2/28/1900 respectively.
			var func = new Yearfrac();
			var args = FunctionsHelper.CreateArgs(31, 59, 0);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(0.07777777778, result.Result);
		}

		[TestMethod]
		public void YearFracWithDatesBefore1March1900WithBasis1ReturnsCorrectResult()
		{
			// Note that 31 and 59 are the Excel OADates for 1/31/1900 and 2/28/1900 respectively.
			var func = new Yearfrac();
			var args = FunctionsHelper.CreateArgs(31, 59, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(0.07671232877, result.Result);
		}

		[TestMethod]
		public void YearFracWithDatesBefore1March1900WithBasis2ReturnsCorrectResult()
		{
			// Note that 31 and 59 are the Excel OADates for 1/31/1900 and 2/28/1900 respectively.
			var func = new Yearfrac();
			var args = FunctionsHelper.CreateArgs(31, 59, 2);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(0.07777777778, result.Result);
		}

		[TestMethod]
		public void YearFracWithDatesBefore1March1900WithBasis3ReturnsCorrectResult()
		{
			// Note that 31 and 59 are the Excel OADates for 1/31/1900 and 2/28/1900 respectively.
			var func = new Yearfrac();
			var args = FunctionsHelper.CreateArgs(31, 59, 3);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(0.07671232877, result.Result);
		}

		[TestMethod]
		public void YearFracWithDatesBefore1March1900WithBasis4ReturnsCorrectResult()
		{
			// Note that 31 and 59 are the Excel OADates for 1/31/1900 and 2/28/1900 respectively.
			var func = new Yearfrac();
			var args = FunctionsHelper.CreateArgs(31, 59, 4);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(0.07777777778, result.Result);
		}

		[TestMethod]
		public void YearFracWithDatesWith1March1900BetweenWithBasis0ReturnsCorrectResult()
		{
			// Note that 31 and 42736 are the Excel OADates for 1/31/1900 and 1/1/2017 respectively.
			var func = new Yearfrac();
			var args = FunctionsHelper.CreateArgs(31, 42736, 0);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(116.91944444444, result.Result);
		}

		[TestMethod]
		public void YearFracWithDatesWith1March1900BetweenWithBasis1ReturnsCorrectResult()
		{
			// Note that 31 and 42736 are the Excel OADates for 1/31/1900 and 1/1/2017 respectively.
			var func = new Yearfrac();
			var args = FunctionsHelper.CreateArgs(31, 42736, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(116.92127427551, result.Result);
		}

		[TestMethod]
		public void YearFracWithDatesWith1March1900BetweenWithBasis2ReturnsCorrectResult()
		{
			// Note that 31 and 42736 are the Excel OADates for 1/31/1900 and 1/1/2017 respectively.
			var func = new Yearfrac();
			var args = FunctionsHelper.CreateArgs(31, 42736, 2);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(118.625, result.Result);
		}

		[TestMethod]
		public void YearFracWithDatesWith1March1900BetweenWithBasis3ReturnsCorrectResult()
		{
			// Note that 31 and 42736 are the Excel OADates for 1/31/1900 and 1/1/2017 respectively.
			var func = new Yearfrac();
			var args = FunctionsHelper.CreateArgs(31, 42736, 3);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(117.0, result.Result);
		}

		[TestMethod]
		public void YearFracWithDatesWith1March1900BetweenWithBasis4ReturnsCorrectResult()
		{
			// Note that 31 and 42736 are the Excel OADates for 1/31/1900 and 1/1/2017 respectively.
			var func = new Yearfrac();
			var args = FunctionsHelper.CreateArgs(31, 42736, 4);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(116.91944444444, result.Result);
		}

		[TestMethod]
		public void YearFracWithBasisAsDoubleReturnsCorrectResult()
		{
			var date1 = new DateTime(2017, 1, 1);
			var date2 = new DateTime(2017, 5, 23);
			var func = new Yearfrac();
			var args = FunctionsHelper.CreateArgs(date1.ToOADate(), date2.ToOADate(), 1.5);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(0.38904109589, result.Result);
		}

		[TestMethod]
		public void YearFracWithBasisIntegerInStringReturnsCorrectResult()
		{
			var date1 = new DateTime(2017, 1, 1);
			var date2 = new DateTime(2017, 5, 23);
			var func = new Yearfrac();
			var args = FunctionsHelper.CreateArgs(date1.ToOADate(), date2.ToOADate(), "1");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(0.38904109589, result.Result);
		}

		[TestMethod]
		public void YearFracWithBasisAsDoubleInStringReturnsCorrectResult()
		{
			var date1 = new DateTime(2017, 1, 1);
			var date2 = new DateTime(2017, 5, 23);
			var func = new Yearfrac();
			var args = FunctionsHelper.CreateArgs(date1.ToOADate(), date2.ToOADate(), "1.5");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(0.38904109589, result.Result);
		}

		[TestMethod]
		public void YearFracWithBasisAsNonEnumeratedNumberReturnsPoundNum()
		{
			// Note that the YEARFRAC function in Excel only accepts 0-4 as valid numbers for the basis parameter.
			var date1 = new DateTime(2017, 1, 1);
			var date2 = new DateTime(2017, 5, 23);
			var func = new Yearfrac();
			var args = FunctionsHelper.CreateArgs(date1.ToOADate(), date2.ToOADate(), 5);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void YearFracWithNegativeBasisNumberReturnsPoundNum()
		{
			var date1 = new DateTime(2017, 1, 1);
			var date2 = new DateTime(2017, 5, 23);
			var func = new Yearfrac();
			var args = FunctionsHelper.CreateArgs(date1.ToOADate(), date2.ToOADate(), -1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void YearFracWithBasisAsNonNumericStringReturnsPoundValue()
		{
			var date1 = new DateTime(2017, 1, 1);
			var date2 = new DateTime(2017, 5, 23);
			var func = new Yearfrac();
			var args = FunctionsHelper.CreateArgs(date1.ToOADate(), date2.ToOADate(), "word");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void YearFracWithBasisAsEmptyStringReturnsPoundValue()
		{
			var date1 = new DateTime(2017, 1, 1);
			var date2 = new DateTime(2017, 5, 23);
			var func = new Yearfrac();
			var args = FunctionsHelper.CreateArgs(date1.ToOADate(), date2.ToOADate(), string.Empty);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void YearFracWithNonNumericStringDateParametersAndNegativeBasisParameterReturnsPoundNum()
		{
			var func = new Yearfrac();
			var args = FunctionsHelper.CreateArgs("word", "word", -1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void YearFracWithNullDateNonNumericStringDateAndNegativeBasisReturnsPoundNum()
		{
			var func = new Yearfrac();
			var args = FunctionsHelper.CreateArgs(null, "word", -1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void YearFracWithDateTimeObjectInputReturnsCorrectResult()
		{
			var date1 = new DateTime(2017, 1, 1);
			var date2 = new DateTime(2017, 5, 23);
			var func = new Yearfrac();
			var args = FunctionsHelper.CreateArgs(date1, date2);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(0.39444444444, result.Result);
		}

		[TestMethod]
		public void YearFracWithDateTimeObjectInputBefore1March1900ReturnsCorrectResult()
		{
			var date1 = new DateTime(1900, 1, 31);
			var date2 = new DateTime(1900, 2, 28);
			var func = new Yearfrac();
			var args = FunctionsHelper.CreateArgs(date1, date2);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(0.07777777778, result.Result);
		}

		[TestMethod]
		public void YearFracWithDateTimeObjectInputWith1March1900BetweenReturnsCorrectResult()
		{
			var date1 = new DateTime(1900, 1, 31);
			var date2 = new DateTime(2017, 1, 1);
			var func = new Yearfrac();
			var args = FunctionsHelper.CreateArgs(date1, date2);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(116.91944444444, result.Result);
		}

		[TestMethod]
		public void YearFracFunctionWorksInDifferentCultureFormats()
		{
			var currentCulture = CultureInfo.CurrentCulture;
			try
			{
				var us = CultureInfo.CreateSpecificCulture("en-US");
				Thread.CurrentThread.CurrentCulture = us;
				using (var package = new ExcelPackage())
				{
					var ws = package.Workbook.Worksheets.Add("Sheet1");
					ws.Cells[2, 2].Value = "1/1/2017";
					ws.Cells[2, 3].Value = "5/23/2017";
					ws.Cells[4, 3].Formula = "YEARFRAC(B2, C2)";
					ws.Calculate();
					Assert.AreEqual(0.39444444444, ws.Cells[4, 3].Value);
				}
				var gb = CultureInfo.CreateSpecificCulture("en-GB");
				Thread.CurrentThread.CurrentCulture = gb;
				using (var package = new ExcelPackage())
				{
					var ws = package.Workbook.Worksheets.Add("Sheet1");
					ws.Cells[2, 2].Value = "1/1/2017";
					ws.Cells[2, 3].Value = "23/5/2017";
					ws.Cells[4, 3].Formula = "YEARFRAC(B2, C2)";
					ws.Calculate();
					Assert.AreEqual(0.39444444444, ws.Cells[4, 3].Value);
				}
				var de = CultureInfo.CreateSpecificCulture("de-DE");
				Thread.CurrentThread.CurrentCulture = de;
				using (var package = new ExcelPackage())
				{
					var ws = package.Workbook.Worksheets.Add("Sheet1");
					ws.Cells[2, 2].Value = "1.1.2017";
					ws.Cells[2, 3].Value = "23.5.2017";
					ws.Cells[4, 3].Formula = "YEARFRAC(B2, C2)";
					ws.Calculate();
					Assert.AreEqual(0.39444444444, ws.Cells[4, 3].Value);
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
