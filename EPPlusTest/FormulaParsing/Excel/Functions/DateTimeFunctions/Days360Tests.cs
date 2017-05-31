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
	public class Days360Tests : DateTimeFunctionsTestBase
	{
		#region Days360 Function (Execute) Tests
		[TestMethod]
		public void Days360ShouldReturnCorrectResultWithNoMethodSpecified()
		{
			var function = new Days360();
			var dt1arg = new DateTime(2013, 1, 1).ToOADate();
			var dt2arg = new DateTime(2013, 3, 31).ToOADate();
			var result = function.Execute(FunctionsHelper.CreateArgs(dt1arg, dt2arg), this.ParsingContext);
			Assert.AreEqual(90, result.Result);
		}

		[TestMethod]
		public void Days360ShouldReturnCorrectResultWithNoMethodSpecifiedMiddleOfMonthDates()
		{
			var function = new Days360();
			var dt1arg = new DateTime(1982, 4, 25).ToOADate();
			var dt2arg = new DateTime(2016, 6, 12).ToOADate();
			var result = function.Execute(FunctionsHelper.CreateArgs(dt1arg, dt2arg), this.ParsingContext);
			Assert.AreEqual(12287, result.Result);
		}

		[TestMethod]
		public void Days360ShouldReturnCorrectResultWithEuroMethodSpecified()
		{
			var function = new Days360();
			var dt1arg = new DateTime(2013, 1, 1).ToOADate();
			var dt2arg = new DateTime(2013, 3, 31).ToOADate();
			var result = function.Execute(FunctionsHelper.CreateArgs(dt1arg, dt2arg, true), this.ParsingContext);
			Assert.AreEqual(89, result.Result);
		}

		[TestMethod]
		public void Days360ShouldHandleFebWithEuroMethodSpecified()
		{
			var function = new Days360();
			var dt1arg = new DateTime(2012, 2, 29).ToOADate();
			var dt2arg = new DateTime(2013, 2, 28).ToOADate();
			var result = function.Execute(FunctionsHelper.CreateArgs(dt1arg, dt2arg, true), this.ParsingContext);
			Assert.AreEqual(359, result.Result);
		}

		[TestMethod]
		public void Days360ShouldHandleFebWithUsMethodSpecified()
		{
			var function = new Days360();
			var dt1arg = new DateTime(2012, 2, 29).ToOADate();
			var dt2arg = new DateTime(2013, 2, 28).ToOADate();
			var result = function.Execute(FunctionsHelper.CreateArgs(dt1arg, dt2arg, false), this.ParsingContext);
			Assert.AreEqual(358, result.Result);
		}

		[TestMethod]
		public void Days360ShouldHandleFebWithUsMethodSpecifiedEndOfMonth()
		{
			var function = new Days360();
			var dt1arg = new DateTime(2013, 2, 28).ToOADate();
			var dt2arg = new DateTime(2013, 3, 31).ToOADate();
			var result = function.Execute(FunctionsHelper.CreateArgs(dt1arg, dt2arg, false), this.ParsingContext);
			Assert.AreEqual(30, result.Result);
		}

		[TestMethod]
		public void Days360ShouldHandleNullFirstDateArgument()
		{
			var function = new Days360();
			var dt2arg = new DateTime(2013, 3, 15).ToOADate();
			var result = function.Execute(FunctionsHelper.CreateArgs(null, dt2arg, false), this.ParsingContext);
			Assert.AreEqual(40755, result.Result);
		}

		[TestMethod]
		public void Days360ShouldHandleNullFirstDateArgumentEndOfMonth()
		{
			var function = new Days360();
			var dt2arg = new DateTime(2013, 3, 31).ToOADate();
			var result = function.Execute(FunctionsHelper.CreateArgs(null, dt2arg, false), this.ParsingContext);
			Assert.AreEqual(40771, result.Result);
		}

		[TestMethod]
		public void Days360ShouldHandleNullSecondDateArgument()
		{
			var function = new Days360();
			var dt1arg = new DateTime(1992, 2, 10).ToOADate();
			var result = function.Execute(FunctionsHelper.CreateArgs(dt1arg, null, false), this.ParsingContext);
			Assert.AreEqual(-33160, result.Result);
		}

		[TestMethod]
		public void Days360ShouldHandleNullSecondDateArgumentEndOfMonth()
		{
			var function = new Days360();
			var dt1arg = new DateTime(2013, 2, 28).ToOADate();
			var result = function.Execute(FunctionsHelper.CreateArgs(dt1arg, null, false), this.ParsingContext);
			Assert.AreEqual(-40740, result.Result);
		}

		[TestMethod]
		public void Days360WithInvalidArgumentReturnsPoundValue()
		{
			var function = new Days360();
			var args = FunctionsHelper.CreateArgs();
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void Days306WithInputAsResultOfDateFunctionReturnsCorrectValue()
		{
			var function = new Days360();
			var dateArg1 = new DateTime(2017, 5, 31);
			var dateArg2 = new DateTime(2017, 6, 30);
			var args = FunctionsHelper.CreateArgs(dateArg1, dateArg2);
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(30, result.Result);
		}

		[TestMethod]
		public void Days360WithInputAsDateStringsReturnsCorrectValue()
		{
			var function = new Days360();
			var args = FunctionsHelper.CreateArgs("5/31/2017", "6/30/2017");
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(30, result.Result);
		}

		[TestMethod]
		public void Days360WithIntegerInputReturnsCorrectValue()
		{
			var function = new Days360();
			var args = FunctionsHelper.CreateArgs(15, 20);
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(5, result.Result);
		}

		[TestMethod]
		public void Days360WithDatesNotAsStringReturnsZero()
		{
			var function = new Days360();
			var args = FunctionsHelper.CreateArgs(5 / 31 / 2017, 6 / 30 / 2017);
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(0, result.Result);
		}

		[TestMethod]
		public void Days360WithGenericStringReturnsPoundValue()
		{
			var function = new Days360();
			var args = FunctionsHelper.CreateArgs("string", "string");
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void Days360WithEmptyStringReturnsPoundValue()
		{
			var function = new Days360();
			var args = FunctionsHelper.CreateArgs("", "");
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void Days360WithStartDateAfterEndDateReturnsCorrectValue()
		{
			var function = new Days360();
			var args = FunctionsHelper.CreateArgs("6/30/2017", "5/31/2017");
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(-30, result.Result);
		}

		[TestMethod]
		public void Days360WithDatesWrittenOutAsStringReturnCorrectValue()
		{
			var function = new Days360();
			var args = FunctionsHelper.CreateArgs("31 May 2017", "30 Jun 2017");
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(30, result.Result);
		}

		[TestMethod]
		public void Days360WithDashesInsteadOfSlashesInStringReturnsCorrectValue()
		{
			var function = new Days360();
			var args = FunctionsHelper.CreateArgs("5-31-2017", "6-30-2017");
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(30, result.Result);
		}

		[TestMethod]
		public void Days360WithPeriodInsteadOfSlashesInStringReturnsCorrectValue()
		{
			// This functionality is different than that of Excel's. Excel does not support inputs of this format,
			// and instead returns a #VALUE!, however many European countries write their dates with periods instead
			//of slashes so EPPlus supports this format of entering dates.
			var function = new Days360();
			var args = FunctionsHelper.CreateArgs("5.31.2017", "6.30.2017");
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(30, result.Result);
		}

		//The following test cases have true in the second parameter for the European method 

		[TestMethod]
		public void Days360WithEuropeanDatesFromDateFunctionReturnsCorrectValue()
		{
			var function = new Days360();
			var dateArg1 = new DateTime(2017, 5, 31);
			var dateArg2 = new DateTime(2017, 6, 30);
			var args = FunctionsHelper.CreateArgs(dateArg1, dateArg2, true);
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(30, result.Result);
		}

		[TestMethod]
		public void Days360WithEuropeanDateArgumentsAsDateStringsReturnsCorrectValue()
		{
			var function = new Days360();
			var args = FunctionsHelper.CreateArgs("5/31/2017", "6/30/2017", true);
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(30, result.Result);
		}

		[TestMethod]
		public void Days360WithEuropeanDateArgumentAndIntegerArgumentsReturnsCorrectValue()
		{
			var function = new Days360();
			var args = FunctionsHelper.CreateArgs(15, 20, true);
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(5, result.Result);
		}

		[TestMethod]
		public void Days360WithEuropeanDateArgumentndDatesNotAStringsReturnsZero()
		{
			var function = new Days360();
			var args = FunctionsHelper.CreateArgs(5 / 31 / 2017, 6 / 30 / 2017, true);
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(0, result.Result);
		}

		[TestMethod]
		public void Days360WithEuropeanDateArgumentAndGeneralStringReturnsPoundValue()
		{
			var function = new Days360();
			var args = FunctionsHelper.CreateArgs("string", "string", true);
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void Days360WithEuropeanDateArgumentWithEmptyStringReturnsPoundValue()
		{
			var function = new Days360();
			var args = FunctionsHelper.CreateArgs("", "", true);
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void Days360WithEuropeanDateArgumentWithStartDateAfterEndDateReturnsCorrectValue()
		{
			var function = new Days360();
			var args = FunctionsHelper.CreateArgs("6/30/2017", "5/31/2017", true);
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(-30, result.Result);
		}

		[TestMethod]
		public void Days360WithEuropeanDateArgumentAndNullFirstDateReturnsCorrectValue()
		{
			var function = new Days360();
			var args = FunctionsHelper.CreateArgs(null, "6/30/2017", true);
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(42300, result.Result);
		}

		[TestMethod]
		public void Days360WithEuropeanDateArgumentAndNullSecondDateReturnsCorrectValue()
		{
			var function = new Days360();
			var args = FunctionsHelper.CreateArgs("5/30/2017", null, true);
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(-42270, result.Result);
		}

		[TestMethod]
		public void Days360WithEuropeanDateArgumentAndDatesWrittenOutAsStringsReturnsCorrectValue()
		{
			var function = new Days360();
			var args = FunctionsHelper.CreateArgs("31 May 2017", "30 Jun 2017", true);
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(30, result.Result);
		}

		[TestMethod]
		public void Days360WithEuropeanDateArgumentAndDatesWrittenWithDashesInsteadOfSlashesReturnsCorrectValue()
		{
			var function = new Days360();
			var args = FunctionsHelper.CreateArgs("5-31-2017", "6-30-2017", true);
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(30, result.Result);
		}

		[TestMethod]
		public void Days360WithEuropeanDateArgumentAndDatesWrittenWithPeriodsInsteadOfSlashesReturnsCorrectValue()
		{
			// This functionality is different than that of Excel's. Excel does not support inputs of this format,
			// and instead returns a #VALUE!, however many European countries write their dates with periods instead
			//of slashes so EPPlus supports this format of entering dates.
			var function = new Days360();
			var args = FunctionsHelper.CreateArgs("5.31.2017", "6.30.2017", true);
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(30, result.Result);
		}

		[TestMethod]
		public void Days360WithGenericStringAsMethodParameterReturnsPoundValue()
		{
			var function = new Days360();
			var args = FunctionsHelper.CreateArgs("5/31/2017", "6/31/2017", "string");
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void Days360WithTrueOrFalseAsStringReturnsCorrecValue()
		{
			var function = new Days360();
			var argsWithTrue = FunctionsHelper.CreateArgs("5/31/2017", "6/30/2017", "true");
			var argsWithFalse = FunctionsHelper.CreateArgs("5/31/2017", "6/30/2017", "false");
			var resultWithTrue = function.Execute(argsWithTrue, this.ParsingContext);
			var resultWithFalse = function.Execute(argsWithFalse, this.ParsingContext);
			Assert.AreEqual(30, resultWithTrue.Result);
			Assert.AreEqual(30, resultWithFalse.Result);
		}

		[TestMethod]
		public void Days360WithIntegerAsMethodParameterReturnsCorrectValue()
		{
			var function = new Days360();
			var args = FunctionsHelper.CreateArgs("5/1/2017", "5/31/2017", 1500);
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(29, result.Result);
		}

		[TestMethod]
		public void Days360WithZeroAsMethodParameterReturnsCorrectValue()
		{
			var function = new Days360();
			var args = FunctionsHelper.CreateArgs("5/1/2017", "5/31/2017", 0);
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(30, result.Result);
		}

		[TestMethod]
		public void Days360WithStringZeroAsMethodParameterReturnsCorrectValue()
		{
			var function = new Days360();
			var args = FunctionsHelper.CreateArgs("5/1/2017", "5/31/2017", "0");
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void Days360WithDateAsMethodParameterReurnsCorrectValue()
		{
			var function = new Days360();
			var dateArg = new DateTime(2017, 6, 3);
			var args = FunctionsHelper.CreateArgs("5/31/2017", "6/30/2017", dateArg);
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(30, result.Result);
		}

		[TestMethod]
		public void Days360WithGermanCultureReturnsCorrectValue()
		{
			var currentCulture = Thread.CurrentThread.CurrentCulture;
			try
			{
				Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("de-DE");
				var function = new Days360();
				var args = FunctionsHelper.CreateArgs("30.5.2017", "30.6.2017");
				var result = function.Execute(args, this.ParsingContext);
				Assert.AreEqual(30, result.Result);
			}
			finally
			{
				Thread.CurrentThread.CurrentCulture = currentCulture;
			}
		}

		#endregion
	}
}
