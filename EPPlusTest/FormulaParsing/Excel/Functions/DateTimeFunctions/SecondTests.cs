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
	public class SecondTests : DateTimeFunctionsTestBase
	{
		#region Second Function (Execute) Tests
		[TestMethod]
		public void SecondWithTimeOnlyReturnsCorrectResult()
		{
			var func = new Second();
			var args = FunctionsHelper.CreateArgs(GetTime(9, 14, 14));
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(14, result.Result);
		}

		[TestMethod]
		public void SecondWithTimeOnlyAsStringReturnsCorrectResult()
		{
			var func = new Second();
			var args = FunctionsHelper.CreateArgs("6:28:48");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(48, result.Result);
		}

		[TestMethod]
		public void SecondWithDateAndTimeAsStringAsInputReturnsCorrectResult()
		{
			var func = new Second();
			var args = FunctionsHelper.CreateArgs("2012-03-27 10:11:12");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(12, result.Result);
		}

		[TestMethod]
		public void SecondWithDateAndTimeAsDifferentStringAsInputReturnsCorrectResult()
		{
			var func = new Second();
			var args = FunctionsHelper.CreateArgs("3/1/1900 8:47:32 PM");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(32, result.Result);
		}

		[TestMethod]
		public void SecondWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Second();

			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SecondWithDateTimeObjectAsInputReturnsCorrectResult()
		{
			var inputDateTime = new DateTime(1900, 3, 1, 8, 47, 32);
			var func = new Second();
			var args = FunctionsHelper.CreateArgs(inputDateTime);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(32, result.Result);
		}

		[TestMethod]
		public void SecondWithOADateAsInputReturnsCorrectResult()
		{
			var inputDateTime = new DateTime(1900, 3, 1, 6, 28, 48);
			var func = new Second();
			var args = FunctionsHelper.CreateArgs(inputDateTime.ToOADate());
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(48, result.Result);
		}

		[TestMethod]
		public void SecondWithMaxTimeValueOnOADateAsInputReturnsCorrectResult()
		{
			var inputDateTime = new DateTime(1900, 3, 1, 23, 59, 59);
			var func = new Second();
			var args = FunctionsHelper.CreateArgs(inputDateTime.ToOADate());
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(59, result.Result);
		}

		[TestMethod]
		public void SecondWithNegativeOADateAsInputReturnsPoundNum()
		{
			var func = new Second();
			var args = FunctionsHelper.CreateArgs(-1.25);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SecondWithNegativeIntegerAsInputReturnsPoundNum()
		{
			var func = new Second();
			var args = FunctionsHelper.CreateArgs(-1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SecondWithOADateAsStringAsInputReturnsCorrectResult()
		{
			// Note that 61.27 is the Excel OADate for 3/1/1900 at 6:28:48.
			var func = new Second();
			var args = FunctionsHelper.CreateArgs("61.27");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(48, result.Result);
		}

		[TestMethod]
		public void SecondWithNegativeOADateAsStringAsInputReturnsPoundNum()
		{
			var func = new Second();
			var args = FunctionsHelper.CreateArgs("-1.25");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SecondWithNonNumericStringAsInputReturnsPoundValue()
		{
			var func = new Second();
			var args = FunctionsHelper.CreateArgs("word");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SecondWithEmptyStringAsInputReturnsPoundValue()
		{
			var func = new Second();
			var args = FunctionsHelper.CreateArgs(string.Empty);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SecondWithZeroAsInputReturnsCorrectResult()
		{
			var func = new Second();
			var args = FunctionsHelper.CreateArgs(0.0);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(0, result.Result);
		}

		[TestMethod]
		public void SecondWithFractionAsInputReturnsCorrectResult()
		{
			// Note that 0.27 is the Excel OADate for 1/0/1900 (the special 0-date) at 6:28:48.
			var func = new Second();
			var args = FunctionsHelper.CreateArgs(0.27);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(48, result.Result);
		}

		[TestMethod]
		public void SecondWithSixthDecimalPlaceProperlyRoundsUp()
		{
			// Note that Excel's max time value only goes out to 5 decimal places;
			// The 6th decimal place is rounded up if greater than or equal to 5,
			// and rounded down if less than 5.
			// Without the 6th decimal place, 61.99999 is the date and time for 3/1/1900 23:59:59.
			var func = new Second();
			var args = FunctionsHelper.CreateArgs(61.999995);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(0, result.Result);
		}

		[TestMethod]
		public void SecondWithSixthDecimalPlaceProperlyRoundsDown()
		{
			// Note that Excel's max time value only goes out to 5 decimal places;
			// The 6th decimal place is rounded up if greater than or equal to 5,
			// and rounded down if less than 5.
			// Without the 6th decimal place, 61.99999 is the date and time for 3/1/1900 23:59:59.
			var func = new Second();
			var args = FunctionsHelper.CreateArgs(61.999994);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(59, result.Result);
		}

		[TestMethod]
		public void SecondAsStringWithSixthDecimalPlaceProperlyRoundsUp()
		{
			// Note that Excel's max time value only goes out to 5 decimal places;
			// The 6th decimal place is rounded up if greater than or equal to 5,
			// and rounded down if less than 5.
			// Without the 6th decimal place, 61.99999 is the date and time for 3/1/1900 23:59:59.
			var func = new Second();
			var args = FunctionsHelper.CreateArgs("61.999995");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(0, result.Result);
		}

		[TestMethod]
		public void SecondAsStringWithSixthDecimalPlaceProperlyRoundsDown()
		{
			// Note that Excel's max time value only goes out to 5 decimal places;
			// The 6th decimal place is rounded up if greater than or equal to 5,
			// and rounded down if less than 5.
			// Without the 6th decimal place, 61.99999 is the date and time for 3/1/1900 23:59:59.
			var func = new Second();
			var args = FunctionsHelper.CreateArgs("61.999994");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(59, result.Result);
		}

		[TestMethod]
		public void SecondWithExcelEpochOADateReturnsCorrectResult()
		{
			var func = new Second();
			var args = FunctionsHelper.CreateArgs(1.0);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(0, result.Result);
		}

		[TestMethod]
		public void SecondWithFractionAsInputProperlyRoundsUpToExcelEpochDate()
		{
			// Note that Excel's max time value only goes out to 5 decimal places;
			// The 6th decimal place is rounded up if greater than or equal to 5,
			// and rounded down if less than 5.
			// 0.99999 is the date and time for 1/0/1900 (the special 0-date) at 23:59:59.
			var func = new Second();
			var args = FunctionsHelper.CreateArgs(0.999995);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(0, result.Result);
		}

		[TestMethod]
		public void SecondAsStringWithFractionAsInputProperlyRoundsUpToExcelEpochDate()
		{
			// Note that Excel's max time value only goes out to 5 decimal places;
			// The 6th decimal place is rounded up if greater than or equal to 5,
			// and rounded down if less than 5.
			// 0.99999 is the date and time for 1/0/1900 (the special 0-date) at 23:59:59.
			var func = new Second();
			var args = FunctionsHelper.CreateArgs("0.999995");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(0, result.Result);
		}

		[TestMethod]
		public void SecondWithDateAndTimeExpressedAsStringWithUnspecifiedAMPMReturnsCorrectResult()
		{
			var func = new Second();
			var args = FunctionsHelper.CreateArgs("3/1/1900 8:47:32");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(32, result.Result);
		}

		[TestMethod]
		public void SecondFunctionWorksInDifferentCultureFormats()
		{
			var currentCulture = CultureInfo.CurrentCulture;
			try
			{
				var us = CultureInfo.CreateSpecificCulture("en-US");
				Thread.CurrentThread.CurrentCulture = us;
				using (var package = new ExcelPackage())
				{
					var ws = package.Workbook.Worksheets.Add("Sheet1");
					ws.Cells[2, 2].Value = "4/14/1900 6:28:48";
					ws.Cells[4, 3].Formula = "SECOND(B2)";
					ws.Calculate();
					Assert.AreEqual(48, ws.Cells[4, 3].Value);
				}
				var gb = CultureInfo.CreateSpecificCulture("en-GB");
				Thread.CurrentThread.CurrentCulture = gb;
				using (var package = new ExcelPackage())
				{
					var ws = package.Workbook.Worksheets.Add("Sheet1");
					ws.Cells[2, 2].Value = "14/4/1900 6:28:48";
					ws.Cells[4, 3].Formula = "SECOND(B2,3)";
					ws.Calculate();
					Assert.AreEqual(48, ws.Cells[4, 3].Value);
				}
				var de = CultureInfo.CreateSpecificCulture("de-DE");
				Thread.CurrentThread.CurrentCulture = de;
				using (var package = new ExcelPackage())
				{
					var ws = package.Workbook.Worksheets.Add("Sheet1");
					ws.Cells[2, 2].Value = "15.1.2014 6:28:48";
					ws.Cells[4, 3].Formula = "SECOND(B2)";
					ws.Cells[3, 2].Value = "15.1.2014";
					ws.Cells[5, 3].Formula = "SECOND(B3)";
					ws.Calculate();
					Assert.AreEqual(48, ws.Cells[4, 3].Value);
					Assert.AreEqual(0, ws.Cells[5, 3].Value);
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
