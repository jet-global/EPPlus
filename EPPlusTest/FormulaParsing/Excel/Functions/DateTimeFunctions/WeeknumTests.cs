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
	public class WeeknumTests : DateTimeFunctionsTestBase
	{
		#region Weeknum Function (Execute) Tests
		//The below tests do not include a second parameter (return type)
		[TestMethod]
		public void WeekNumWtihNoInputReturnsPoundValue()
		{
			var func = new Weeknum();

			var r1 = func.Execute(FunctionsHelper.CreateArgs(), this.ParsingContext);

			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)r1.Result).Type);
		}

		[TestMethod]
		public void WeekNumWithDateFunctionInputReturnsCorrectResult()
		{
			var func = new Weeknum();

			var dt = new DateTime(2017, 1, 5);

			var r1 = func.Execute(FunctionsHelper.CreateArgs(dt), this.ParsingContext);

			Assert.AreEqual(1, r1.Result);
		}

		[TestMethod]
		public void WeekNumWithDateAsStringReturnsCorrectResult()
		{
			var func = new Weeknum();

			var dt = "1/10/2017";

			var r1 = func.Execute(FunctionsHelper.CreateArgs(dt), this.ParsingContext);

			Assert.AreEqual(2, r1.Result);
		}

		[TestMethod]
		public void WeekNumWithDateAsStringWithDashesReturnsCorrectResult()
		{
			var func = new Weeknum();

			var dt = "1-10-2017";

			var r1 = func.Execute(FunctionsHelper.CreateArgs(dt), this.ParsingContext);

			Assert.AreEqual(2, r1.Result);
		}

		[TestMethod]
		public void WeekNumWithDateNotAsStringReturnsCorrectResult()
		{
			var func = new Weeknum();

			var dt = 1/10/2017;

			var r1 = func.Execute(FunctionsHelper.CreateArgs(dt), this.ParsingContext);

			Assert.AreEqual(0, r1.Result);
		}

		[TestMethod]
		public void WeekNumWithStringArgumentReturnsPoundValue()
		{
			var func = new Weeknum();

			var dt1 = "testString";
			var dt2 = "";

			var r1 = func.Execute(FunctionsHelper.CreateArgs(dt1), this.ParsingContext);
			var r2 = func.Execute(FunctionsHelper.CreateArgs(dt2), this.ParsingContext);

			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)r1.Result).Type);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)r2.Result).Type);
		}

		[TestMethod]
		public void WeekNumWithOADateArgumentReturnsCorrectValue()
		{
			var func = new Weeknum();

			var dt = new DateTime(2017, 1, 10).ToOADate();

			var r1 = func.Execute(FunctionsHelper.CreateArgs(dt), this.ParsingContext);

			Assert.AreEqual(2, r1.Result);
		}

		[TestMethod]
		public void WeekNumWithIntegerArgumentReturnsCorrectValue()
		{
			var func = new Weeknum();

			var dt = 9;

			var r1 = func.Execute(FunctionsHelper.CreateArgs(dt), this.ParsingContext);

			Assert.AreEqual(2, r1.Result);
		}

		[TestMethod]
		public void WeekNumWithNegativeIntReturnsPoundNum()
		{
			var func = new Weeknum();

			var dt = -5;

			var r1 = func.Execute(FunctionsHelper.CreateArgs(dt), this.ParsingContext);

			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)r1.Result).Type);
		}

		//Below are the tests that include the second parameter (return type)

		[TestMethod]
		public void WeekNumWithReturnType1OrOmmittedReturnsCorrectValue()
		{
			var func = new Weeknum();

			var dt = new DateTime(2017,1,5);

			var r1 = func.Execute(FunctionsHelper.CreateArgs(dt, 1), this.ParsingContext);
			var r2 = func.Execute(FunctionsHelper.CreateArgs(dt), this.ParsingContext);

			Assert.AreEqual(1, r1.Result);
			Assert.AreEqual(1, r2.Result);
		}

		[TestMethod]
		public void WeekNumWithReturnType2ReturnsCorrectValue()
		{
			var func = new Weeknum();

			var dt = new DateTime(2017, 1, 5);

			var r1 = func.Execute(FunctionsHelper.CreateArgs(dt, 2), this.ParsingContext);

			Assert.AreEqual(2, r1.Result);
		}

		[TestMethod]
		public void WeekNumWithReturnType11ReturnsCorrectValue()
		{
			var func = new Weeknum();

			var dt = new DateTime(2017, 1, 5);

			var r1 = func.Execute(FunctionsHelper.CreateArgs(dt, 11), this.ParsingContext);

			Assert.AreEqual(2, r1.Result);
		}

		[TestMethod]
		public void WeekNumWithReturnType12ReturnsCorrectValue()
		{
			var func = new Weeknum();

			var dt = new DateTime(2017, 1, 5);

			var r1 = func.Execute(FunctionsHelper.CreateArgs(dt, 12), this.ParsingContext);

			Assert.AreEqual(2, r1.Result);
		}

		[TestMethod]
		public void WeekNumWithReturnType13ReturnsCorrectValue()
		{
			var func = new Weeknum();

			var dt = new DateTime(2017, 1, 5);

			var r1 = func.Execute(FunctionsHelper.CreateArgs(dt, 13), this.ParsingContext);

			Assert.AreEqual(2, r1.Result);
		}

		[TestMethod]
		public void WeekNumWithReturnType14ReturnsCorrectValue()
		{
			var func = new Weeknum();

			var dt = new DateTime(2017, 1, 5);

			var r1 = func.Execute(FunctionsHelper.CreateArgs(dt, 14), this.ParsingContext);

			Assert.AreEqual(2, r1.Result);
		}

		[TestMethod]
		public void WeekNumWithReturnType15ReturnsCorrectValue()
		{
			var func = new Weeknum();

			var dt = new DateTime(2017, 1, 5);

			var r1 = func.Execute(FunctionsHelper.CreateArgs(dt, 15), this.ParsingContext);

			Assert.AreEqual(1, r1.Result);
		}

		[TestMethod]
		public void WeekNumWithReturnType16ReturnsCorrectValue()
		{
			var func = new Weeknum();

			var dt = new DateTime(2017, 1, 5);

			var r1 = func.Execute(FunctionsHelper.CreateArgs(dt, 16), this.ParsingContext);

			Assert.AreEqual(1, r1.Result);
		}

		[TestMethod]
		public void WeekNumWithReturnType17ReturnsCorrectValue()
		{
			var func = new Weeknum();

			var dt = new DateTime(2017, 1, 5);

			var r1 = func.Execute(FunctionsHelper.CreateArgs(dt, 17), this.ParsingContext);

			Assert.AreEqual(1, r1.Result);
		}

		[TestMethod]
		public void WeekNumWithReturnType21ReturnsCorrectValue()
		{
			var func = new Weeknum();

			var dt = new DateTime(2017, 1, 5);

			var r1 = func.Execute(FunctionsHelper.CreateArgs(dt, 21), this.ParsingContext);

			Assert.AreEqual(1, r1.Result);
		}

		[TestMethod]
		public void WeekNumWithInvalidReturnTypeReturnsPoundNum()
		{
			var func = new Weeknum();

			var dt = new DateTime(2017, 1, 5);

			var r1 = func.Execute(FunctionsHelper.CreateArgs(dt, 5), this.ParsingContext);

			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)r1.Result).Type);
		}

		[TestMethod]
		public void WeekNumWithNumericStringReturnTypeReturnsCorrectValue()
		{
			var func = new Weeknum();

			var dt = new DateTime(2017, 1, 5);

			var r1 = func.Execute(FunctionsHelper.CreateArgs(dt, "1"), this.ParsingContext);

			Assert.AreEqual(1, r1.Result);
		}

		[TestMethod]
		public void WeekNumWithStringReturnTypeReturnsPoundValue()
		{
			var func = new Weeknum();

			var dt = new DateTime(2017, 1, 5);

			var r1 = func.Execute(FunctionsHelper.CreateArgs(dt, "testString"), this.ParsingContext);

			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)r1.Result).Type);
		}

		[TestMethod]
		public void WeekNumWithNoFirstParameterAndValidReturnTypeReturnsPoundNA()
		{
			var func = new Weeknum();

			var r1 = func.Execute(FunctionsHelper.CreateArgs(null, 1), this.ParsingContext);

			Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)r1.Result).Type);
		}
		#endregion
	}
}
