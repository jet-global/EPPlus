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
		#endregion
	}
}
