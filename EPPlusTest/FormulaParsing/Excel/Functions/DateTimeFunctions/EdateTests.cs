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
	public class EdateTests : DateTimeFunctionsTestBase
	{
		#region EdateTests Function (Execute) Tests
		[TestMethod]
		public void EdateShouldReturnCorrectResult()
		{
			var func = new Edate();

			var dt1arg = new DateTime(2012, 1, 31).ToOADate();
			var dt2arg = new DateTime(2013, 1, 1).ToOADate();
			var dt3arg = new DateTime(2013, 2, 28).ToOADate();

			var r1 = func.Execute(FunctionsHelper.CreateArgs(dt1arg, 1), this.ParsingContext);
			var r2 = func.Execute(FunctionsHelper.CreateArgs(dt2arg, -1), this.ParsingContext);
			var r3 = func.Execute(FunctionsHelper.CreateArgs(dt3arg, 2), this.ParsingContext);
			var r4 = func.Execute(FunctionsHelper.CreateArgs(dt3arg), this.ParsingContext);
			var dt1 = DateTime.FromOADate((double)r1.Result);
			var dt2 = DateTime.FromOADate((double)r2.Result);
			var dt3 = DateTime.FromOADate((double)r3.Result);

			var exp1 = new DateTime(2012, 2, 29);
			var exp2 = new DateTime(2012, 12, 1);
			var exp3 = new DateTime(2013, 4, 28);

			Assert.AreEqual(exp1, dt1, "dt1 was not " + exp1.ToString("yyyy-MM-dd") + ", but " + dt1.ToString("yyyy-MM-dd"));
			Assert.AreEqual(exp2, dt2, "dt1 was not " + exp2.ToString("yyyy-MM-dd") + ", but " + dt2.ToString("yyyy-MM-dd"));
			Assert.AreEqual(exp3, dt3, "dt1 was not " + exp3.ToString("yyyy-MM-dd") + ", but " + dt3.ToString("yyyy-MM-dd"));
			Assert.AreEqual(eErrorType.Value, (r4.Result as ExcelErrorValue).Type);
		}

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
			Assert.AreEqual(0, result.Result);
		}

		[TestMethod]
		public void EdateWithOADateOfOneAsFirstParameterReturnsOne()
		{
			var func = new Edate();
			var args = FunctionsHelper.CreateArgs(1,0);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(1, result.Result);
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
			Assert.AreEqual(61, result.Result);
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
			Assert.AreEqual(59, result.Result);
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
			Assert.AreEqual(58, result.Result);
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
			Assert.AreEqual(31, result.Result);
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
			Assert.AreEqual(32, result.Result);
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
			Assert.AreEqual(59, result.Result);
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
			Assert.AreEqual(1, result.Result);
		}
		#endregion
	}
}
