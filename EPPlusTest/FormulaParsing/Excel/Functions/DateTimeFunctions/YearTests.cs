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
	public class YearTests : DateTimeFunctionsTestBase
	{
		#region Year Function (Execute) Tests
		[TestMethod]
		public void YearWithDateAsStringAsInputReturnsCorrectYear()
		{
			var func = new Year();
			var args = FunctionsHelper.CreateArgs("2012-03-12");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(2012, result.Result);
		}

		[TestMethod]
		public void YearWithLongDateAsStringAsInputReturnsCorrectYear()
		{
			var func = new Year();
			var args = FunctionsHelper.CreateArgs("March 12, 2012");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(2012, result.Result);
		}

		[TestMethod]
		public void YearWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Year();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void YearFunctionWorksInDifferentCultureDateFormats()
		{
			var currentCulture = CultureInfo.CurrentCulture;
			try
			{
				var us = CultureInfo.CreateSpecificCulture("en-US");
				Thread.CurrentThread.CurrentCulture = us;
				using (var package = new ExcelPackage())
				{
					var ws = package.Workbook.Worksheets.Add("Sheet1");
					ws.Cells[2, 2].Value = "1/15/2014";
					ws.Cells[6, 3].Formula = "YEAR(B2)";
					ws.Calculate();
					Assert.AreEqual(2014, ws.Cells[6, 3].Value);
				}

				var gb = CultureInfo.CreateSpecificCulture("en-GB");
				Thread.CurrentThread.CurrentCulture = gb;
				using (var package = new ExcelPackage())
				{
					var ws = package.Workbook.Worksheets.Add("Sheet1");
					ws.Cells[2, 2].Value = "15/1/2014";
					ws.Cells[6, 3].Formula = "YEAR(B2)";
					ws.Calculate();
					Assert.AreEqual(2014, ws.Cells[6, 3].Value);
				}
				var de = CultureInfo.CreateSpecificCulture("de-DE");
				Thread.CurrentThread.CurrentCulture = de;
				using (var package = new ExcelPackage())
				{
					var ws = package.Workbook.Worksheets.Add("Sheet1");
					ws.Cells[2, 2].Value = "15.1.2014";
					ws.Cells[4, 3].Formula = "YEAR(B2)";
					ws.Calculate();
					Assert.AreEqual(2014, ws.Cells[4, 3].Value);
				}
			}
			finally
			{
				Thread.CurrentThread.CurrentCulture = currentCulture;
			}
		}

		[TestMethod]
		public void YearWithDateAsIntegerReturnsCorrectYear()
		{
			// Note that 42874 is the OADate for May 19, 2017.
			var func = new Year();
			var args = FunctionsHelper.CreateArgs(42874);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(2017, result.Result);
		}

		[TestMethod]
		public void YearWithDateAsDoubleReturnsCorrectYear()
		{
			// Note that 42874.34114 is the OADate representation of
			// some time on May 19, 2017.
			var func = new Year();
			var args = FunctionsHelper.CreateArgs(42874.34114);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(2017, result.Result);
		}

		[TestMethod]
		public void YearWithZeroAsInputReturns1900AsYear()
		{
			// Excel treats 0 as a special case OADate.
			var func = new Year();
			var args = FunctionsHelper.CreateArgs(0);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(1900, result.Result);
		}

		[TestMethod]
		public void YearWithNegativeIntegerAsInputReturnsPoundNum()
		{
			var func = new Year();
			var args = FunctionsHelper.CreateArgs(-1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void YearWithNegativeDoubleAsInputReturnsPoundNum()
		{
			var func = new Year();
			var args = FunctionsHelper.CreateArgs(-1.5);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void YearWithNonDateStringAsInputReturnsPoundValue()
		{
			var func = new Year();
			var args = FunctionsHelper.CreateArgs("word");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void YearWithEmptyStringAsInputReturnsPoundValue()
		{
			var func = new Year();
			var args = FunctionsHelper.CreateArgs(string.Empty);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void YearWithIntegerInStringAsInputReturnsCorrectYear()
		{
			// Note that 42874 is the OADate representation of May 19, 2017.
			var func = new Year();
			var args = FunctionsHelper.CreateArgs("42874");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(2017, result.Result);
		}

		[TestMethod]
		public void YearWithDoubleInStringAsInputReturnsCorrectYear()
		{
			// Note that 42874.34114 is the OADate representation of some time
			// on May 19, 2017.
			var date = new DateTime(2017, 5, 19);
			date.AddHours(6);
			date.AddMinutes(30);
			var func = new Year();
			var args = FunctionsHelper.CreateArgs(date.ToOADate());
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(2017, result.Result);
		}

		[TestMethod]
		public void YearWithNegativeIntegerInStringAsInputReturnsPoundNum()
		{
			var func = new Year();
			var args = FunctionsHelper.CreateArgs("-1");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void YearWithNegativeDoubleInStringAsInputReturnsPoundNum()
		{
			var func = new Year();
			var args = FunctionsHelper.CreateArgs("-1.5");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void YearWithDateTimeObjectReturnsCorrectYear()
		{
			var date = new DateTime(2017, 5, 22);
			var func = new Year();
			var args = FunctionsHelper.CreateArgs(date);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(2017, result.Result);
		}

		[TestMethod]
		public void YearWithDateTimeObjectForOADateLessThan61ReturnsCorrectYear()
		{
			var date = new DateTime(1900, 2, 28);
			var func = new Year();
			var args = FunctionsHelper.CreateArgs(date);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(1900, result.Result);
		}

		[TestMethod]
		public void YearWithOADate1ReturnsCorrectYear()
		{
			// OADate 1 corresponds to 1/1/1900 in Excel. Test the case where Excel's epoch date
			// is used as the input.
			var func = new Year();
			var args = FunctionsHelper.CreateArgs(1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(1900, result.Result);
		}

		[TestMethod]
		public void YearWithOADateBetween0And1ReturnsZero()
		{
			// Test the case where a time value is given for Excel's
			// "zeroeth" day, which Excel treats as a special day.
			var func = new Year();
			var args = FunctionsHelper.CreateArgs(0.5);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(1900, result.Result);
		}

		[TestMethod]
		public void YearWithNegativeOADateBetweenNegative1And0ReturnsPoundNum()
		{
			// Test the case where a negative time value is given for Excel's
			// "zeroeth" day, which Excel treats as a special day.
			var func = new Year();
			var args = FunctionsHelper.CreateArgs(-0.5);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void YearWithDateAs0InStringReturnsCorrectResult()
		{
			var func = new Year();
			var args = FunctionsHelper.CreateArgs("0");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(1900, result.Result);
		}

		[TestMethod]
		public void YearWithFractionalDateInStringReturnsCorrectResult()
		{
			var func = new Year();
			var args = FunctionsHelper.CreateArgs("0.5");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(1900, result.Result);
		}

		[TestMethod]
		public void YearFunctionWithErrorValuesAsInputReturnsTheInputErrorValue()
		{
			var func = new Year();
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
