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
	public class HourTests : DateTimeFunctionsTestBase
	{
		#region Hour Function (Execute) Tests
		[TestMethod]
		public void HourWithTimeOnlyReturnsCorrectResult()
		{
			var func = new Hour();
			var args = FunctionsHelper.CreateArgs(this.GetTime(9, 14, 14));
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(9, result.Result);
		}

		[TestMethod]
		public void HourWithMaxTimeOnlyReturnsCorrectResult()
		{
			var func = new Hour();
			var args = FunctionsHelper.CreateArgs(this.GetTime(23, 59, 59));
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(23, result.Result);
		}

		[TestMethod]
		public void HourWithTimeOnlyAsStringReturnsCorrectResult()
		{
			var func = new Hour();
			var args = FunctionsHelper.CreateArgs("6:28:48");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(6, result.Result);
		}

		[TestMethod]
		public void HourWithDateAndTimeAsStringAsInputReturnsCorrectResult()
		{
			var func = new Hour();
			var args = FunctionsHelper.CreateArgs("2013-03-27 10:11:12");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(10, result.Result);
		}

		[TestMethod]
		public void HourWithDateAndTimeAsDifferentStringAsInputReturnsCorrectResult()
		{
			var func = new Hour();
			var args = FunctionsHelper.CreateArgs("3/1/1900 8:47:32 PM");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(20, result.Result);
		}

		[TestMethod]
		public void HourWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Hour();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void HourWithDateTimeObjectAsInputReturnsCorrectResult()
		{
			var inputDateTime = new DateTime(1900, 3, 1, 8, 47, 32);
			var func = new Hour();
			var args = FunctionsHelper.CreateArgs(inputDateTime);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(8, result.Result);
		}

		[TestMethod]
		public void HourWithMaxTimeValueOnOADateAsInputReturnsCorrectResult()
		{
			var inputDateTime = new DateTime(1900, 3, 1, 23, 59, 59);
			var func = new Hour();
			var args = FunctionsHelper.CreateArgs(inputDateTime.ToOADate());
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(23, result.Result);
		}

		[TestMethod]
		public void HourWithNegativeOADateAsInputReturnsPoundNum()
		{
			var func = new Hour();
			var args = FunctionsHelper.CreateArgs(-1.25);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void HourWithNegativeIntegerAsInputReturnsPoundNum()
		{
			var func = new Hour();
			var args = FunctionsHelper.CreateArgs(-1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void HourWithOADateAsStringAsInputReturnsCorrectResult()
		{
			// Note that 61.27 is the Excel OADate for 3/1/1900 at 6:28:48.
			var func = new Hour();
			var args = FunctionsHelper.CreateArgs("61.27");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(6, result.Result);
		}

		[TestMethod]
		public void HourWithNegativeOADateAsStringAsInputReturnsPoundNum()
		{
			var func = new Hour();
			var args = FunctionsHelper.CreateArgs("-1.25");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void HourWithNonNumericStringAsInputReturnsPoundValue()
		{
			var func = new Hour();
			var args = FunctionsHelper.CreateArgs("word");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void HourWithEmptyStringAsInputReturnsPoundValue()
		{
			var func = new Hour();
			var args = FunctionsHelper.CreateArgs(string.Empty);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void HourWithZeroAsInputReturnsCorrectResult()
		{
			var func = new Hour();
			var args = FunctionsHelper.CreateArgs(0.0);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(0, result.Result);
		}

		[TestMethod]
		public void HourWithFractionAsInputReturnsCorrectResult()
		{
			// Note that 0.27 is the Excel OADate for 1/0/1900 (the special 0-date) at 6:28:48.
			var func = new Hour();
			var args = FunctionsHelper.CreateArgs(0.27);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(6, result.Result);
		}

		[TestMethod]
		public void HourWithSixthDecimalPlaceProperlyRoundsUp()
		{
			// Note that Excel's max time value only goes out to 5 decimal places;
			// The 6th decimal place is rounded up if greater than or equal to 5,
			// and rounded down if less than 5.
			// Without the 6th decimal place, 61.99999 is the max date and time value for 3/1/1900 23:59:59.
			var func = new Hour();
			var args = FunctionsHelper.CreateArgs(61.999995);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(0, result.Result);
		}

		[TestMethod]
		public void HourWithSixthDecimalPlaceProperlyRoundsDown()
		{
			// Note that Excel's max time value only goes out to 5 decimal places;
			// The 6th decimal place is rounded up if greater than or equal to 5,
			// and rounded down if less than 5.
			// Without the 6th decimal place, 61.99999 is the max date and time value for 3/1/1900 23:59:59.
			var func = new Hour();
			var args = FunctionsHelper.CreateArgs(61.999994);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(23, result.Result);
		}

		[TestMethod]
		public void HourAsStringWithSixthDecimalPlaceProperlyRoundsUp()
		{
			// Note that Excel's max time value only goes out to 5 decimal places;
			// The 6th decimal place is rounded up if greater than or equal to 5,
			// and rounded down if less than 5.
			// Without the 6th decimal place, 61.99999 is the max date and time value for 3/1/1900 23:59:59.
			var func = new Hour();
			var args = FunctionsHelper.CreateArgs("61.999995");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(0, result.Result);
		}

		[TestMethod]
		public void HourAsStringWithSixthDecimalPlaceProperlyRoundsDown()
		{
			// Note that Excel's max time value only goes out to 5 decimal places;
			// The 6th decimal place is rounded up if greater than or equal to 5,
			// and rounded down if less than 5.
			// Without the 6th decimal place, 61.99999 is the max date and time value for 3/1/1900 23:59:59.
			var func = new Hour();
			var args = FunctionsHelper.CreateArgs("61.999994");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(23, result.Result);
		}

		[TestMethod]
		public void HourWithExcelEpochOADateReturnsCorrectResult()
		{
			var func = new Hour();
			var args = FunctionsHelper.CreateArgs(1.0);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(0, result.Result);
		}

		[TestMethod]
		public void HourWithFractionAsInputProperlyRoundsUpToExcelEpochDate()
		{
			// Note that Excel's max time value only goes out to 5 decimal places;
			// The 6th decimal place is rounded up if greater than or equal to 5,
			// and rounded down if less than 5.
			// 0.99999 is the max date and time value for 1/0/1900 (the special 0-date) at 23:59:59.
			var func = new Hour();
			var args = FunctionsHelper.CreateArgs(0.999995);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(0, result.Result);
		}

		[TestMethod]
		public void HourAsStringWithFractionAsInputProperlyRoundsUpToExcelEpochDate()
		{
			// Note that Excel's max time value only goes out to 5 decimal places;
			// The 6th decimal place is rounded up if greater than or equal to 5,
			// and rounded down if less than 5.
			// 0.99999 is the max date and time value for 1/0/1900 (the special 0-date) at 23:59:59.
			var func = new Hour();
			var args = FunctionsHelper.CreateArgs("0.999995");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(0, result.Result);
		}

		[TestMethod]
		public void HourWithDateAndTimeExpressedAsStringWithUnspecifiedAMPMReturnsCorrectResult()
		{
			var func = new Hour();
			var args = FunctionsHelper.CreateArgs("3/1/1900 8:47:32");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(8, result.Result);
		}

		[TestMethod]
		public void HourFunctionWorksInDifferentCultureFormats()
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
					ws.Cells[4, 3].Formula = "HOUR(B2)";
					ws.Calculate();
					Assert.AreEqual(6, ws.Cells[4, 3].Value);
				}
				var gb = CultureInfo.CreateSpecificCulture("en-GB");
				Thread.CurrentThread.CurrentCulture = gb;
				using (var package = new ExcelPackage())
				{
					var ws = package.Workbook.Worksheets.Add("Sheet1");
					ws.Cells[2, 2].Value = "14/4/1900 6:28:48";
					ws.Cells[4, 3].Formula = "HOUR(B2)";
					ws.Calculate();
					Assert.AreEqual(6, ws.Cells[4, 3].Value);
				}
				var de = CultureInfo.CreateSpecificCulture("de-DE");
				Thread.CurrentThread.CurrentCulture = de;
				using (var package = new ExcelPackage())
				{
					var ws = package.Workbook.Worksheets.Add("Sheet1");
					ws.Cells[2, 2].Value = "15.1.2014 6:28:48";
					ws.Cells[4, 3].Formula = "HOUR(B2)";
					ws.Cells[3, 2].Value = "15.1.2014";
					ws.Cells[5, 3].Formula = "HOUR(B3)";
					ws.Calculate();
					Assert.AreEqual(6, ws.Cells[4, 3].Value);
					Assert.AreEqual(0, ws.Cells[5, 3].Value);
				}
			}
			finally
			{
				Thread.CurrentThread.CurrentCulture = currentCulture;
			}
		}

		[TestMethod]
		public void HourFunctionWithErrorValuesAsInputReturnsTheInputErrorValue()
		{
			var func = new Hour();
			var argNA = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.NA));
			var argNAME = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Name));
			var argVALUE = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Value));
			var argNUM = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Num));
			var argDIV0 = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Div0));
			var argREF = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Ref));
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
