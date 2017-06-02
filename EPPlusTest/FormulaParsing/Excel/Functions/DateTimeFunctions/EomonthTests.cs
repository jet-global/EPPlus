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
	public class EomonthTests : DateTimeFunctionsTestBase
	{
		#region Eomonth Function (Execute) Tests
		[TestMethod]
		public void EomonthReturnsEndOfMonthWithPositiveOffset()
		{
			var function = new Eomonth();
			var date = new DateTime(2013, 2, 4).ToOADate();
			var result = function.Execute(FunctionsHelper.CreateArgs(date, 3), this.ParsingContext);
			Assert.AreEqual(41425d, result.Result);
			var resultDate = DateTime.FromOADate(result.ResultNumeric);
			Assert.AreEqual(5, resultDate.Month);
		}

		[TestMethod]
		public void EomonthWithNegativeOffsetReturnsCorrectValue()
		{
			var function = new Eomonth();
			var date = new DateTime(2013, 2, 4).ToOADate();
			var result = function.Execute(FunctionsHelper.CreateArgs(date, -3), this.ParsingContext);
			Assert.AreEqual(41243d, result.Result);
			var resultDate = DateTime.FromOADate(result.ResultNumeric);
			Assert.AreEqual(11, resultDate.Month);
			Assert.AreEqual(30, resultDate.Day);
			Assert.AreEqual(2012, resultDate.Year);
		}

		[TestMethod]
		public void EomonthReturnsEndOfMonthWithPositiveOffsetFromDateTime()
		{
			var function = new Eomonth();
			var date = new DateTime(2013, 2, 4);
			var result = function.Execute(FunctionsHelper.CreateArgs(date, 3), this.ParsingContext);
			Assert.AreEqual(41425d, result.Result);
			var resultDate = DateTime.FromOADate(result.ResultNumeric);
			Assert.AreEqual(5, resultDate.Month);
		}

		[TestMethod]
		public void EomonthReturnsEndOfMonthWitZeroOffsetFromDateTime()
		{
			var function = new Eomonth();
			var date = new DateTime(2013, 2, 4);
			var result = function.Execute(FunctionsHelper.CreateArgs(date, 0.0), this.ParsingContext);
			Assert.AreEqual(new DateTime(2013, 2, 28).ToOADate(), result.Result);
		}

		[TestMethod]
		public void EomonthWithNegativeOffsetFromDateTime()
		{
			var function = new Eomonth();
			var date = new DateTime(2013, 2, 4);
			var result = function.Execute(FunctionsHelper.CreateArgs(date, -3), this.ParsingContext);
			Assert.AreEqual(41243d, result.Result);
			var resultDate = DateTime.FromOADate(result.ResultNumeric);
			Assert.AreEqual(11, resultDate.Month);
			Assert.AreEqual(30, resultDate.Day);
			Assert.AreEqual(2012, resultDate.Year);
		}

		[TestMethod]
		public void EomonthWithStringOADateFirstArgument()
		{
			var function = new Eomonth();
			var date = new DateTime(2013, 2, 4);
			var dateString = $"{date.ToOADate()}";
			var result = function.Execute(FunctionsHelper.CreateArgs(dateString, 0), this.ParsingContext);
			var expected = new DateTime(2013, 2, 28);
			Assert.AreEqual(expected.ToOADate(), result.Result);
		}

		[TestMethod]
		public void EomonthWithStringDateFirstArgument()
		{
			var function = new Eomonth();
			var dateString = "4 FEB 2013";
			var result = function.Execute(FunctionsHelper.CreateArgs(dateString, 0), this.ParsingContext);
			var expected = new DateTime(2013, 2, 28);
			Assert.AreEqual(expected.ToOADate(), result.Result);
		}

		[TestMethod]
		public void EomonthWithStringOffset()
		{
			var function = new Eomonth();
			var date = new DateTime(2013, 2, 4);
			var result = function.Execute(FunctionsHelper.CreateArgs(date, "12"), this.ParsingContext);
			var expected = new DateTime(2014, 2, 28);
			Assert.AreEqual(expected.ToOADate(), result.Result);
		}

		[TestMethod]
		public void EomonthWithSmallStringDateAsOffset()
		{
			var function = new Eomonth();
			var date = new DateTime(2013, 2, 4);
			var result = function.Execute(FunctionsHelper.CreateArgs(date, "1/12/1900"), this.ParsingContext);
			var expected = new DateTime(2014, 2, 28);
			Assert.AreEqual(expected.ToOADate(), result.Result);
		}

		[TestMethod]
		public void EomonthWithSmallDateAsSecondArgument()
		{
			var function = new Eomonth();
			var date = new DateTime(2013, 2, 4);
			var offsetDate = new DateTime(1900, 1, 12);
			var result = function.Execute(FunctionsHelper.CreateArgs(date, offsetDate), this.ParsingContext);
			var expected = new DateTime(2014, 2, 28);
			Assert.AreEqual(expected.ToOADate(), result.Result);
		}

		[TestMethod]
		public void EomonthWithLargeStringDateAsOffset()
		{
			var function = new Eomonth();
			var date = new DateTime(2013, 2, 4);
			var result = function.Execute(FunctionsHelper.CreateArgs(date, "1/12/2017"), this.ParsingContext);
			var expected = new DateTime(5575, 5, 31);
			Assert.AreEqual(expected.ToOADate(), result.Result);
		}

		[TestMethod]
		public void EomonthWithZerosForInput()
		{
			var function = new Eomonth();
			var result = function.Execute(FunctionsHelper.CreateArgs(0, 0), this.ParsingContext);
			Assert.AreEqual(31d, result.Result);
		}
		// 60 returns 59 (special case)
		[TestMethod]
		public void EomonthOnFebruary291900ReturnsFebruary28()
		{
			var function = new Eomonth();
			// OADate 60 in Excel represents 2/29/1900, a day that never actually happened but existed in Lotus 1-2-3.
			var result = function.Execute(FunctionsHelper.CreateArgs(60, 0), this.ParsingContext);
			Assert.AreEqual(59d, result.Result);
		}

		[TestMethod]
		public void EomonthOnFeb291900AsDateReturnsFebruary28()
		{
			var function = new Eomonth();
			var result = function.Execute(FunctionsHelper.CreateArgs("2/29/1900", 1), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void EomonthDateWithFractionInputReturnsCorrectValue()
		{
			var function = new Eomonth();
			var result = function.Execute(FunctionsHelper.CreateArgs(0.3, 1), this.ParsingContext);
			Assert.AreEqual(59d, result.Result);
		}

		// string first arg, garbage
		[TestMethod]
		public void EomonthWithGarbageStringAsFirstArgumentReturnsPoundValue()
		{
			var function = new Eomonth();
			var result = function.Execute(FunctionsHelper.CreateArgs("garbage", 0), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		// string second arg, garbage
		[TestMethod]
		public void EomonthWithGarbageStringAsOffsetArgumentReturnsPoundValue()
		{
			var function = new Eomonth();
			var result = function.Execute(FunctionsHelper.CreateArgs(0, "garbage"), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void EomonthInvalidOADateReturnsPoundNum()
		{
			var function = new Eomonth();
			var result = function.Execute(FunctionsHelper.CreateArgs(-1, 0), this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		// OADate plus negative offset yields invalid date
		[TestMethod]
		public void EomonthOffsetPlusOADateGeneratesInvalidDateReturnsPoundNum()
		{
			var function = new Eomonth();
			var result = function.Execute(FunctionsHelper.CreateArgs(1, -1), this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}
		// 0, null -> #N/A

		[TestMethod]
		public void EomonthNullFirstArgumentReturnsPoundNA()
		{
			var function = new Eomonth();
			var result = function.Execute(FunctionsHelper.CreateArgs(null, 0), this.ParsingContext);
			Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void EomonthNullSecondArgumentReturnsPoundNA()
		{
			var function = new Eomonth();
			var result = function.Execute(FunctionsHelper.CreateArgs(0, null), this.ParsingContext);
			Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void EomonthWithoutArgumentsReturnsPoundValue()
		{
			var function = new Eomonth();
			var arguments = FunctionsHelper.CreateArgs();
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void EomonthWithDateAsStringReturnsCorrectValue()
		{
			var function = new Eomonth();
			var arguments = FunctionsHelper.CreateArgs("5/15/2017", 1);
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(42916d, result.Result);
		}

		[TestMethod]
		public void EomonthWithDateNotAsStringReturnsCorrectValue()
		{
			var function = new Eomonth();
			var arguments = FunctionsHelper.CreateArgs(5 / 15 / 2017, 1);
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(59d, result.Result);
		}

		[TestMethod]
		public void EomonthWithPositiveIntegerReturnsCorrectValue()
		{
			var function = new Eomonth();
			var arguments = FunctionsHelper.CreateArgs(20, 1);
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(59d, result.Result);
		}

		[TestMethod]
		public void EomonthWithDoulbeInputReturnsCorrectValue()
		{
			var function = new Eomonth();
			var arguments = FunctionsHelper.CreateArgs(150.5, 1);
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(182d, result.Result);
		}

		[TestMethod]
		public void EomonthWithZeroInputReturnsCorrectValue()
		{
			var function = new Eomonth();
			var arguments = FunctionsHelper.CreateArgs(0.0, 1);
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(59d, result.Result);
		}

		[TestMethod]
		public void EomonthWithDateWithDashInsteadOfSlashReturnsCorrectValue()
		{
			var function = new Eomonth();
			var arguments = FunctionsHelper.CreateArgs("5-15-2017", 1);
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(42916d, result.Result);
		}

		[TestMethod]
		public void EomonthWithDateWithPeriodsInsteadOfSlashesReturnsCorrectValue()
		{
			var function = new Eomonth();
			var arguments = FunctionsHelper.CreateArgs("5.15.2017", 1);
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(42916d, result.Result);
		}

		[TestMethod]
		public void EomonthWithMonthArgumentAsDateFunctionReturnsCorretValue()
		{
			var function = new Eomonth();
			var endDate = new DateTime(2017, 6, 25);
			var arguments = FunctionsHelper.CreateArgs("5/15/2017", endDate);
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(1348962d, result.Result);
		}

		[TestMethod]
		public void EomonthWithMonthArgumentAsDateNotAsStringReturnsCorretValue()
		{
			var function = new Eomonth();
			var arguments = FunctionsHelper.CreateArgs("5/15/2017", 6 / 25 / 2017);
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(42886d, result.Result);
		}

		[TestMethod]
		public void EomonthWithMonthArgumentAsDateWrittenOutReturnsCorrectValue()
		{
			var function = new Eomonth();
			var arguments = FunctionsHelper.CreateArgs("5/15/2017", "25 JUN 2017");
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(1348962d, result.Result);
		}

		[TestMethod]
		public void EomonthWithMonthArgumentAsNonZeroIntReturnsCorrectValue()
		{
			var function = new Eomonth();
			var arguments = FunctionsHelper.CreateArgs("5/15/2017", 20);
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(43496d, result.Result);
		}

		[TestMethod]
		public void EomonthWithMonthArgumentAsDoubleReturnsCorrectValue()
		{
			var function = new Eomonth();
			var arguments = FunctionsHelper.CreateArgs("5/15/2017", 20.6);
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(43496d, result.Result);
		}

		[TestMethod]
		public void EoMonthWithGermanCultureReturnCorrectValue()
		{
			var currentCulture = Thread.CurrentThread.CurrentCulture;
			try
			{
				Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("de-DE");
				var function = new Eomonth();
				var arguments = FunctionsHelper.CreateArgs("15.5.2017", 20);
				var result = function.Execute(arguments, this.ParsingContext);
				Assert.AreEqual(43496d, result.Result);
			}
			finally
			{
				Thread.CurrentThread.CurrentCulture = currentCulture;
			}
		}
		#endregion
	}
}
