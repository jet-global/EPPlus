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
	/// <summary>
	/// Summary description for DayFunctionTests
	/// </summary>
	[TestClass]
	public class DayTests : DateTimeFunctionsTestBase
	{
		#region Day Function (Execute) Tests
		[TestMethod]
		public void DayWithDoubleInputReturnsDayInMonth()
		{
			var date = new DateTime(2012, 3, 12);
			var func = new Day();
			var args = FunctionsHelper.CreateArgs(date.ToOADate());
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(12, result.Result);
		}

		[TestMethod]
		public void DayWithDateAsStringReturnsDayOfMonth()
		{
			var func = new Day();
			var args = FunctionsHelper.CreateArgs("2012-03-12");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(12, result.Result);
		}

		[TestMethod]
		public void DayWithDateAsLongStringReturnsDayOfMonth()
		{
			var func = new Day();
			var args = FunctionsHelper.CreateArgs("May 19, 2017");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(19, result.Result);
		}

		[TestMethod]
		public void DayWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Day();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void DayFunctionWorksInDifferentCultureDateFormats()
		{
			// Test the case where the date is represented under
			// different cultures.
			var currentCulture = CultureInfo.CurrentCulture;
			try
			{
				var us = CultureInfo.CreateSpecificCulture("en-US");
				Thread.CurrentThread.CurrentCulture = us;
				using (var package = new ExcelPackage())
				{
					var ws = package.Workbook.Worksheets.Add("Sheet1");
					ws.Cells[2, 2].Value = "1/15/2014";
					ws.Cells[5, 3].Formula = "DAY(B2)";
					ws.Calculate();
					Assert.AreEqual(15, ws.Cells[5, 3].Value);
				}

				var gb = CultureInfo.CreateSpecificCulture("en-GB");
				Thread.CurrentThread.CurrentCulture = gb;
				using (var package = new ExcelPackage())
				{
					var ws = package.Workbook.Worksheets.Add("Sheet1");
					ws.Cells[2, 2].Value = "15/1/2014";
					ws.Cells[5, 3].Formula = "DAY(B2)";
					ws.Calculate();
					Assert.AreEqual(15, ws.Cells[5, 3].Value);
				}

				var de = CultureInfo.CreateSpecificCulture("de-DE");
				Thread.CurrentThread.CurrentCulture = de;
				using (var package = new ExcelPackage())
				{
					var ws = package.Workbook.Worksheets.Add("Sheet1");
					ws.Cells[2, 2].Value = "15.1.2014";
					ws.Cells[5, 3].Formula = "DAY(B2)";
					ws.Calculate();
					Assert.AreEqual(15, ws.Cells[5, 3].Value);
				}
			}
			finally
			{
				Thread.CurrentThread.CurrentCulture = currentCulture;
			}
		}

		[TestMethod]
		public void DayWithDateAsIntegerReturnsDayOfMonth()
		{
			// Test the case where the date is entered as an integer.
			// Note that 42874 is the OADate for May 19, 2017.
			var func = new Day();
			var args = FunctionsHelper.CreateArgs(42874);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(19, result.Result);
		}

		[TestMethod]
		public void DayWithZeroAsInputReturnsZeroAsDateOfMonth()
		{
			// Test the case where zero is the input.
			var func = new Day();
			var args = FunctionsHelper.CreateArgs(0);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(0, result.Result);
		}

		[TestMethod]
		public void DayWithNegativeIntegerAsInputReturnsPoundNum()
		{
			// Test the case where a negative integer is the input.
			var func = new Day();
			var args = FunctionsHelper.CreateArgs(-1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void DayWithNegativeDoubleAsInputReturnsPoundNum()
		{
			// Test the case where a negative double is the input
			var func = new Day();
			var args = FunctionsHelper.CreateArgs(-1.5);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void DayWithNonDateStringAsInputReturnsPoundValue()
		{
			// Test the case where a non-date string is the input.
			var func = new Day();
			var args = FunctionsHelper.CreateArgs("word");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void DayWithEmptyStringAsInputReturnsPoundValue()
		{
			// Test the case where an empty string is the input.
			var func = new Day();
			var args = FunctionsHelper.CreateArgs("");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void DayWithIntegerInStringAsInputReturnsDayOfMonth()
		{
			// Test the case where the input is an integer expressed as a string.
			// Note that 42874 is the OADate representation of May 19, 2017.
			var func = new Day();
			var args = FunctionsHelper.CreateArgs("42874");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(19, result.Result);
		}

		[TestMethod]
		public void DayWithDoubleInStringAsInputReturnsDayOfMonth()
		{
			// Test the case where the input is a double expressed as a string.
			// Note that 42874.34114 is the OADate representation of some time
			// on May 19, 2017.
			var func = new Day();
			var args = FunctionsHelper.CreateArgs("42874.43114");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(19, result.Result);
		}

		[TestMethod]
		public void DayWithNegativeIntegerInStringAsInputReturnsPoundNum()
		{
			// Test the case where the input is a negative integer expressed
			// as a string.
			var func = new Day();
			var args = FunctionsHelper.CreateArgs("-1");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void DayWithNegativeDoubleInStringAsInputReturnsPoundNum()
		{
			// Test the case where the input is a negative double expressed
			// as a string.
			var func = new Day();
			var args = FunctionsHelper.CreateArgs("-1.5");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}
		#endregion
	}
}
