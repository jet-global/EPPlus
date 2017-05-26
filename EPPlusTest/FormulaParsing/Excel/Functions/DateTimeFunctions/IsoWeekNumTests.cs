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
	public class IsoWeekNumTests : DateTimeFunctionsTestBase
	{
		#region IsoWeekNum Function (Execute) Tests
		[TestMethod]
		public void IsoWeekWithOADateInputReturnsCorrectValue()
		{
			var function = new IsoWeekNum();
			var args = new DateTime(2013, 1, 1).ToOADate();
			var result = function.Execute(FunctionsHelper.CreateArgs(args), this.ParsingContext);
			Assert.AreEqual(1, result.Result);
		}

		[TestMethod]
		public void IsoWeekNumWithInvalidArgumentReturnsPoundValue()
		{
			var function = new IsoWeekNum();
			var args = FunctionsHelper.CreateArgs();
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void IsoWeekNumWithDateFunctionInputReturnsCorrectValue()
		{
			var function = new IsoWeekNum();
			var date = new DateTime(2017, 5, 26);
			var args = FunctionsHelper.CreateArgs(date);
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(21, result.Result);
		}

		[TestMethod]
		public void IsoWeekNumWithDateAsStringReturnsCorrectValue()
		{
			var function = new IsoWeekNum();
			var args = FunctionsHelper.CreateArgs("5/26/2017");
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(21, result.Result);
		}

		[TestMethod]
		public void IsoWeekNumWithDateNotAsStringReturnsCorrectValue()
		{
			var function = new IsoWeekNum();
			var args = FunctionsHelper.CreateArgs(5 / 26 / 2017);
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(52, result.Result);
		}

		[TestMethod]
		public void IsoWeekNumWithPositiveIntInputReturnsCorrectValue()
		{
			var function = new IsoWeekNum();
			var args = FunctionsHelper.CreateArgs(5);
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(1, result.Result);

		}

		[TestMethod]
		public void IsoWeekNumWithNegativeIntInputReturnsPoundNum()
		{
			var function = new IsoWeekNum();
			var args = FunctionsHelper.CreateArgs(-10);
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void IsoWeekNumWithGeneralStringInputReturnsPoundValue()
		{
			var function = new IsoWeekNum();
			var args = FunctionsHelper.CreateArgs("string");
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void IsoWeekNumWithEmptyStringInputReturnsPoundValue()
		{
			var function = new IsoWeekNum();
			var args = FunctionsHelper.CreateArgs("");
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void IsoWeekNumWithZeroInputReturnsCorrectValue()
		{
			var function = new IsoWeekNum();
			var args = FunctionsHelper.CreateArgs(0);
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(52, result.Result);
		}

		[TestMethod]
		public void IsoWeekNumWithDotsInsteadOfSlashesReturnsCorrectValue()
		{
			//This functionality is different than that of Excel's. Excel reutrns a #VALUE! when the date is enterd this
			//way, however many European cultures write their dates with periods instead of slashes so EPPlus supports 
			//dates in this format. 
			var function = new IsoWeekNum();
			var args = FunctionsHelper.CreateArgs("5.26.2017");
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(21, result.Result);
		}

		[TestMethod]
		public void IsoWeekNumWithDashesInsteadOfSlashesReturnsCorrectValue()
		{
			var function = new IsoWeekNum();
			var args = FunctionsHelper.CreateArgs("5-26-2017");
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(21, result.Result);
		}

		[TestMethod]
		public void IsoWeekNumWithGermanDateAsStringWithPeriodReturnsCorrectResult()
		{
			//Test case for dates written in the form used by most European cultures.
			var currentCulture = Thread.CurrentThread.CurrentCulture;
			try
			{
				Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("de-DE");
				var function = new IsoWeekNum();
				var args = FunctionsHelper.CreateArgs("26.5.2017");
				var result = function.Execute(args, this.ParsingContext);
				Assert.AreEqual(21, result.Result);
			}
			finally
			{
				Thread.CurrentThread.CurrentCulture = currentCulture;
			}
		}

		[TestMethod]
		public void IsoWeekNumWithTooManyArgsReturnsPoundNA()
		{
			//This functionality is different than that of Excel's. Excel does not let you compute the formula if you have
			//too many arguments, however EPPlus lets you put more than one argument in. Now it returns a #NA! to indicate 
			//the user has put in too many arugments. 
			var function = new IsoWeekNum();
			var args = FunctionsHelper.CreateArgs("5-26-2017", "5-26-2017");
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)result.Result).Type);
		}
		#endregion
	}
}
