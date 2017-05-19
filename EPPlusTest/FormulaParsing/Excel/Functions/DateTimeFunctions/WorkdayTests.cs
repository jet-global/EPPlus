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
	public class WorkdayTests : DateTimeFunctionsTestBase
	{
		// The below Test Cases have no Holiday parameter supplied to them.
		// The below Test Cases has no negative second parameters.
		#region Workday Function (Execute) Tests
		[TestMethod]
		public void WorkdayWithOADateParameterReturnsCorrectResult()
		{
			var inputDate = new DateTime(2014, 1, 1).ToOADate();
			var expectedDate = new DateTime(2014, 1, 29).ToOADate();

			var func = new Workday();
			var args = FunctionsHelper.CreateArgs(inputDate, 20);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate, result.Result);
		}

		[TestMethod]
		public void WorkdayShouldReturnCorrectResultWithFourDaysSupplied()
		{
			var inputDate = new DateTime(2014, 1, 1).ToOADate();
			var expectedDate = new DateTime(2014, 1, 7).ToOADate();

			var func = new Workday();
			var args = FunctionsHelper.CreateArgs(inputDate, 4);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate, result.Result);
		}

		[TestMethod]
		public void WorkdayWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Workday();

			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WorkdayWithStringInputReturnsPoundValue()
		{
			var function = new Workday();

			var input1 = "testString";
			var input2 = "";

			var result1 = function.Execute(FunctionsHelper.CreateArgs(input1, 10), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(input2, 10), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result1.Result).Type);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result1.Result).Type);
		}

		[TestMethod]
		public void WorkdayWithIntegerInputReturnsCorrectValue()
		{
			var function = new Workday();

			var input = 10;

			var result = function.Execute(FunctionsHelper.CreateArgs(input, 10), this.ParsingContext);
			Assert.AreEqual(24.00, result.Result);
		}

		[TestMethod]
		public void WorkdayWithDateAsStringReturnsCorrectValue()
		{
			var function = new Workday();

			var input = "1/2/2017";

			var result = function.Execute(FunctionsHelper.CreateArgs(input, 10), this.ParsingContext);
			Assert.AreEqual(42751.00, result.Result);
		}

		[TestMethod]
		public void WorkdayWithDateNotAsStringReturnsCorrectValue()
		{
			var function = new Workday();

			var input = 1 / 2 / 2017;

			var result = function.Execute(FunctionsHelper.CreateArgs(input, 10), this.ParsingContext);
			Assert.AreEqual(13, result.Result);
		}

		[TestMethod]
		public void WorkdayWithDATEFunctionAsInputReturnsCorrectValue()
		{
			var function = new Workday();

			var input = new DateTime(2017, 1, 2);

			var result = function.Execute(FunctionsHelper.CreateArgs(input, 10), this.ParsingContext);
			Assert.AreEqual(42751.00, result.Result);
		}

		[TestMethod]
		public void WorkdayWithDateUsingPeriodsIntseadOfSlashesReturnsCorrectValue()
		{
			var function = new Workday();

			var input = "1.2.2017";

			var result = function.Execute(FunctionsHelper.CreateArgs(input, 10), this.ParsingContext);
			Assert.AreEqual(42751.00, result.Result);
		}

		[TestMethod]
		public void WorkdayWithNegativeFirstParamReturnsPoundNum()
		{
			var function = new Workday();

			var input = -1;

			var result = function.Execute(FunctionsHelper.CreateArgs(input, 10), this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WorkdayWithZeroInputReturnsCorrectValue()
		{
			var function = new Workday();

			var input = 0;

			var result = function.Execute(FunctionsHelper.CreateArgs(input, 10), this.ParsingContext);
			Assert.AreEqual(13, result.Result);
		}

		[TestMethod]
		public void WorkdayWithNullFirstParamReturnsPoundNum()
		{
			var function = new Workday();

			var result = function.Execute(FunctionsHelper.CreateArgs(null, 10), this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WorkdayWithDateUsingDashesInsteadOfSlashesReturnsCorrectResult()
		{
			var function = new Workday();

			var input = "1-2-2017";

			var result = function.Execute(FunctionsHelper.CreateArgs(input, 10), this.ParsingContext);
			Assert.AreEqual(42751.00, result.Result);
		}

		// The below Test Cases have negative second parameters.
		[TestMethod]
		public void WorkdayShouldReturnCorrectResultWithNegativeArg()
		{
			var inputDate = new DateTime(2016, 6, 15).ToOADate();
			var expectedDate = new DateTime(2016, 5, 4).ToOADate();

			var func = new Workday();
			var args = FunctionsHelper.CreateArgs(inputDate, -30);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(DateTime.FromOADate(expectedDate), DateTime.FromOADate((double)result.Result));
		}

		








		


		// The below Test Cases involve the Holiday parameter.
		[TestMethod]
		public void WorkdayWithNegativeArgShouldReturnCorrectWhenArrayOfHolidayDatesIsSupplied()
		{
			var inputDate = new DateTime(2016, 7, 27).ToOADate();
			var holidayDate1 = new DateTime(2016, 7, 11).ToOADate();
			var holidayDate2 = new DateTime(2016, 7, 8).ToOADate();
			var expectedDate = new DateTime(2016, 6, 13).ToOADate();

			var func = new Workday();
			var args = FunctionsHelper.CreateArgs(inputDate, -30, FunctionsHelper.CreateArgs(holidayDate1, holidayDate2));
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate, result.Result);
		}

		[TestMethod]
		public void WorkdayWithNegativeArgShouldReturnCorrectWhenRangeWithHolidayDatesIsSupplied()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("test");
				ws.Cells["A1"].Value = new DateTime(2016, 7, 27).ToOADate();
				ws.Cells["B1"].Value = new DateTime(2016, 7, 11).ToOADate();
				ws.Cells["B2"].Value = new DateTime(2016, 7, 8).ToOADate();
				ws.Cells["B3"].Formula = "WORKDAY(A1,-30, B1:B2)";
				ws.Calculate();

				var expectedDate = new DateTime(2016, 6, 13).ToOADate();
				var actualDate = ws.Cells["B3"].Value;
				Assert.AreEqual(expectedDate, actualDate);
			}
		}
		#endregion
	}
}
