﻿/*******************************************************************************
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
	public class EdateTests : DateTimeFunctionsTestBase
	{
		#region EdateTests Function (Execute) Tests
		[TestMethod]
		public void EdateWithDateAsStringAsFirstParameterReturnsCorrectResult()
		{
			var date = new DateTime(2017, 5, 22);
			var func = new Edate();
			var args = FunctionsHelper.CreateArgs("5/22/2017", 0);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(date.ToOADate(), result.Result);
		}

		[TestMethod]
		public void EdateWithDateAsLongStringAsFirstParameterReturnsCorrectResult()
		{
			var date = new DateTime(2017, 5, 22);
			var func = new Edate();
			var args = FunctionsHelper.CreateArgs("May 22, 2017", 0);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(date.ToOADate(), result.Result);
		}

		[TestMethod]
		public void EdateWithNonDateStringAsFirstParameterReturnsPoundValue()
		{
			var func = new Edate();
			var args = FunctionsHelper.CreateArgs("word", 0);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void EdateWithEmptyStringAsFirstParameterReturnsPoundValue()
		{
			var func = new Edate();
			var args = FunctionsHelper.CreateArgs(string.Empty, 0);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void EdateWithIntegerAsStringAsFirstParameterReturnsCorrectResult()
		{
			// Note that 42877 is the OADate for 5/22/2017.
			var date = new DateTime(2017, 5, 22);
			var func = new Edate();
			var args = FunctionsHelper.CreateArgs("42877", 0);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(date.ToOADate(), result.Result);
		}

		[TestMethod]
		public void EdateWithDoubleAsStringAsFirstParameterReturnsCorrectResult()
		{
			// Note that 42877.5 is the OADate for some time on 5/22/2017.
			var date = new DateTime(2017, 5, 22);
			var func = new Edate();
			var args = FunctionsHelper.CreateArgs("42877.5", 0);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(date.ToOADate(), result.Result);
		}
		
		[TestMethod]
		public void EdateWithNegativeIntegerAsStringAsFirstParameterReturnsPoundNum()
		{
			var func = new Edate();
			var args = FunctionsHelper.CreateArgs("-1", 0);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void EdateWithNegativeDoubleAsStringAsFirstParameterReturnsPoundNum()
		{
			var func = new Edate();
			var args = FunctionsHelper.CreateArgs("-1.5", 0);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void EdateWithIntegerAsFirstParameterReturnsCorrectResult()
		{
			// Note that 42877 is the OADate for 5/22/2017
			var date = new DateTime(2017, 5, 22);
			var func = new Edate();
			var args = FunctionsHelper.CreateArgs(42877, 0);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(date.ToOADate(), result.Result);
		}

		[TestMethod]
		public void EdateWithDoubleAsFirstParameterReturnsCorrectResult()
		{
			// Note that 42877.5 is the OADate for some time on 5/22/2017.
			var date = new DateTime(2017, 5, 22);
			var func = new Edate();
			var args = FunctionsHelper.CreateArgs(42877.5, 0);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(date.ToOADate(), result.Result);
		}

		[TestMethod]
		public void EdateWithZeroAsFirstParameterReturnsZero()
		{
			// Zero is a special case and requires special output.
			var func = new Edate();
			var args = FunctionsHelper.CreateArgs(0,0);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(0.0, result.Result);
		}

		[TestMethod]
		public void EdateWithFractionAsStringAsFirstParameterReturnsZero()
		{
			// Zero is a special case and requires special output.
			var func = new Edate();
			var args = FunctionsHelper.CreateArgs("0.5", 0);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(0.0, result.Result);
		}

		[TestMethod]
		public void EdateWithOADateOfOneAsFirstParameterReturnsOne()
		{
			var func = new Edate();
			var args = FunctionsHelper.CreateArgs(1,0);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(1.0, result.Result);
		}

		[TestMethod]
		public void EdateWithNegativeIntegerAsFirstParameterReturnsPoundNum()
		{
			var func = new Edate();
			var args = FunctionsHelper.CreateArgs(-1,0);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void EdateWithNegativeDoubleAsFirstParameterReturnsPoundNum()
		{
			var func = new Edate();
			var args = FunctionsHelper.CreateArgs(-1.5,0);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void EdateWithNegativeFirstParamaterAndNonNumericStringSecondParamaterReturnsPoundNum()
		{
			var func = new Edate();
			var args = FunctionsHelper.CreateArgs(-1,"word");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void EdateWithNonNumericStringAsBothParametersReturnsPoundValue()
		{
			var func = new Edate();
			var args = FunctionsHelper.CreateArgs("word","word");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void EdateWithDate1March1900ReturnsCorrectResult()
		{
			// Note that for testing dates before 3/1/1900 (which has OADate 61), the OADate has to be written literally,
			// rather than as the result of calling ToOADate() on an expected DateTime object,
			// to ensure that the serial number being output is based on Excel's OADates and not
			// System.DateTime's OADates, which are all off by one for dates before 3/1/1900.
			var inputDate = new DateTime(1900, 3, 1);
			var func = new Edate();
			var args = FunctionsHelper.CreateArgs(inputDate, 0);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(61.0, result.Result);
		}

		[TestMethod]
		public void EdateWithDate28February1900ReturnsCorrectResult()
		{
			// Note that for testing dates before 3/1/1900 (which has OADate 61), the OADate has to be written literally,
			// rather than as the result of calling ToOADate() on an expected DateTime object,
			// to ensure that the serial number being output is based on Excel's OADates and not
			// System.DateTime's OADates, which are all off by one for dates before 3/1/1900.
			var inputDate = new DateTime(1900, 2, 28);
			var func = new Edate();
			var args = FunctionsHelper.CreateArgs(inputDate, 0);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(59.0, result.Result);
		}

		[TestMethod]
		public void EdateWithDate27February1900ReturnsCorrectResult()
		{
			// Note that for testing dates before 3/1/1900 (which has OADate 61), the OADate has to be written literally,
			// rather than as the result of calling ToOADate() on an expected DateTime object,
			// to ensure that the serial number being output is based on Excel's OADates and not
			// System.DateTime's OADates, which are all off by one for dates before 3/1/1900.
			var inputDate = new DateTime(1900, 2, 27);
			var func = new Edate();
			var args = FunctionsHelper.CreateArgs(inputDate, 0);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(58.0, result.Result);
		}

		[TestMethod]
		public void EdateWithNullFirstParamaterReturnsPoundNA()
		{
			var func = new Edate();
			var args = FunctionsHelper.CreateArgs(null, 0);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void EdateWithNullSecondParamaterReturnsPoundNA()
		{
			var inputDate = new DateTime(2017, 1, 1);
			var func = new Edate();
			var args = FunctionsHelper.CreateArgs(inputDate, null);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void EdateWithBothParamatersNullReturnsPoundNA()
		{
			var func = new Edate();
			var args = FunctionsHelper.CreateArgs(null, null);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void EdateWithDateTimeObjectAsFirstParameterReturnsCorrectResult()
		{
			var inputDate = new DateTime(2017, 1, 1);
			var expectedDate = new DateTime(2017, 2, 1);
			var func = new Edate();
			var args = FunctionsHelper.CreateArgs(inputDate, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void EdateWithIntegerAsSecondParameterReturnsCorrectResult()
		{
			var inputDate = new DateTime(2017, 1, 1);
			var expectedDate = new DateTime(2017, 3, 1);
			var func = new Edate();
			var args = FunctionsHelper.CreateArgs(inputDate, 2);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void EdateWithDoubleAsSecondParameterReturnsCorrectResult()
		{
			var inputDate = new DateTime(2017, 1, 1);
			var expectedDate = new DateTime(2017, 2, 1);
			var func = new Edate();
			var args = FunctionsHelper.CreateArgs(inputDate, 1.5);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void EdateWithNegativeIntegerAsSecondParamaterReturnsCorrectResult()
		{
			var inputDate = new DateTime(2017, 1, 1);
			var expectedDate = new DateTime(2016, 12, 1);
			var func = new Edate();
			var args = FunctionsHelper.CreateArgs(inputDate, -1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void EdateWithNegativeDoubleAsSecondParamaterReturnsCorrectResult()
		{
			var inputDate = new DateTime(2017, 1, 1);
			var expectedDate = new DateTime(2016, 12, 1);
			var func = new Edate();
			var args = FunctionsHelper.CreateArgs(inputDate, -1.5);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void EdateWithIntegerAsStringAsSecondParamaterReturnsCorrectResult()
		{
			var inputDate = new DateTime(2017, 1, 1);
			var expectedDate = new DateTime(2017, 2, 1);
			var func = new Edate();
			var args = FunctionsHelper.CreateArgs(inputDate, "1");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void EdateWithDoubleAsStringAsSecondParamaterReturnsCorrectResult()
		{
			var inputDate = new DateTime(2017, 1, 1);
			var expectedDate = new DateTime(2017, 2, 1);
			var func = new Edate();
			var args = FunctionsHelper.CreateArgs(inputDate, "1.5");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void EdateWithNonNumericStringAsSecondParamaterReturnsPoundValue()
		{
			var inputDate = new DateTime(2017, 1, 1);
			var func = new Edate();
			var args = FunctionsHelper.CreateArgs(inputDate, "word");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void EdateWithEmptyStringAsSecondParamaterReturnsPoundValue()
		{
			var inputDate = new DateTime(2017, 1, 1);
			var func = new Edate();
			var args = FunctionsHelper.CreateArgs(inputDate, string.Empty);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void EdateWithOriginalDateWithMoreDaysThanCalculatedDateReturnsCorrectResult()
		{
			var inputDate = new DateTime(2017, 3, 31);
			var expectedDate = new DateTime(2017, 4, 30);
			var func = new Edate();
			var args = FunctionsHelper.CreateArgs(inputDate, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void EdateWithZeroAsFirstParamaterAndIntegerAsSecondParamaterReturnsCorrectResult()
		{
			// Zero is a special case date and requires specific output.
			var func = new Edate();
			var args = FunctionsHelper.CreateArgs(0, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(31.0, result.Result);
		}

		[TestMethod]
		public void EdateWithZeroAsStringAsFirstParamaterAndIntegerAsSecondParamaterReturnsCorrectResult()
		{
			// Zero is a special case date and requires specific output.
			var func = new Edate();
			var args = FunctionsHelper.CreateArgs("0", 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(31.0, result.Result);
		}

		[TestMethod]
		public void EdateWhereCalculatedDateWouldBeBeforeExcelEpochDateReturnsPoundNum()
		{
			var inputDate = new DateTime(1900, 1, 1);
			var func = new Edate();
			var args = FunctionsHelper.CreateArgs(inputDate, -1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void EdateWithCalculatedDateBefore1March1900ReturnsCorrectOADate()
		{
			// Note that for testing dates before 3/1/1900 (which has OADate 61), the OADate has to be written literally,
			// rather than as the result of calling ToOADate() on an expected DateTime object,
			// to ensure that the serial number being output is based on Excel's OADates and not
			// System.DateTime's OADates, which are all off by one for dates before 3/1/1900.
			var inputDate = new DateTime(1900, 1, 1);
			var func = new Edate();
			var args = FunctionsHelper.CreateArgs(inputDate, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(32.0, result.Result);
		}

		[TestMethod]
		public void EdateWithMonthsAddedToReachNonExistentDateReturnsCorrectResult()
		{
			// Note that for testing dates before 3/1/1900 (which has OADate 61), the OADate has to be written literally,
			// rather than as the result of calling ToOADate() on an expected DateTime object,
			// to ensure that the serial number being output is based on Excel's OADates and not
			// System.DateTime's OADates, which are all off by one for dates before 3/1/1900.
			var inputDate = new DateTime(1900, 1, 29);
			var func = new Edate();
			var args = FunctionsHelper.CreateArgs(inputDate, 1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(59.0, result.Result);
		}

		[TestMethod]
		public void EdateWithMonthsSubtractedToReachTheEpochDateReturnsPoundNum()
		{
			var inputDate = new DateTime(1900, 1, 31);
			var func = new Edate();
			var args = FunctionsHelper.CreateArgs(inputDate, -1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void EdateWithMonthsSubtractedToCalculateRightBeforeTheEpochDateReturnsCorrectResult()
		{
			// Note that for testing dates before 3/1/1900 (which has OADate 61), the OADate has to be written literally,
			// rather than as the result of calling ToOADate() on an expected DateTime object,
			// to ensure that the serial number being output is based on Excel's OADates and not
			// System.DateTime's OADates, which are all off by one for dates before 3/1/1900.
			var inputDate = new DateTime(1900, 2, 1);
			var func = new Edate();
			var args = FunctionsHelper.CreateArgs(inputDate, -1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(1.0, result.Result);
		}

		[TestMethod]
		public void EdateWithMonthsAddedReturnsCorrectResult()
		{
			// Note that 196 is the OADate for 7/14/1900.
			var func = new Edate();
			var args = FunctionsHelper.CreateArgs("4/14/1900", 3);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(196.0, result.Result);
		}

		[TestMethod]
		public void EdateWithAYearOfMonthsAddedReturnsCorrectResult()
		{
			// Note that 470 is the OADate for 4/14/1901.
			var func = new Edate();
			var args = FunctionsHelper.CreateArgs("4/14/1900", 12);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(470.0, result.Result);
		}

		[TestMethod]
		public void EdateWithMonthsSubtractedReturnsCorrectResult()
		{
			// Note that 45 is the OADate for 2/14/1900.
			var func = new Edate();
			var args = FunctionsHelper.CreateArgs("4/14/1900", -2);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(45.0, result.Result);
		}

		[TestMethod]
		public void EdateFunctionWorksInDifferentCultureFormats()
		{
			// Note that 196.0 is the Excel OADate for July 14, 1900 (7/14/1900).
			var currentCulture = CultureInfo.CurrentCulture;
			try
			{
				var us = CultureInfo.CreateSpecificCulture("en-US");
				Thread.CurrentThread.CurrentCulture = us;
				using (var package = new ExcelPackage())
				{
					var ws = package.Workbook.Worksheets.Add("Sheet1");
					ws.Cells[2, 2].Value = "4/14/1900";
					ws.Cells[4, 3].Formula = "EDATE(B2,3)";
					ws.Calculate();
					Assert.AreEqual(196.0, ws.Cells[4, 3].Value);
				}
				var gb = CultureInfo.CreateSpecificCulture("en-GB");
				Thread.CurrentThread.CurrentCulture = gb;
				using (var package = new ExcelPackage())
				{
					var ws = package.Workbook.Worksheets.Add("Sheet1");
					ws.Cells[2, 2].Value = "14/4/1900";
					ws.Cells[4, 3].Formula = "EDATE(B2,3)";
					ws.Calculate();
					Assert.AreEqual(196.0, ws.Cells[4, 3].Value);
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
