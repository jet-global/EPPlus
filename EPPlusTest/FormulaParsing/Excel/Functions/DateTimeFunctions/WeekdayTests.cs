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
	public class WeekdayTests : DateTimeFunctionsTestBase
	{
		#region Weekday Function (Execute) Tests
		[TestMethod]
		public void WeekdayShouldReturnCorrectResultForASundayWhenReturnTypeIs1()
		{
			var func = new Weekday();
			var result = func.Execute(FunctionsHelper.CreateArgs(new DateTime(2012, 4, 1).ToOADate(), 1), this.ParsingContext);
			Assert.AreEqual(1, result.Result);
		}

		[TestMethod]
		public void WeekdayShouldReturnCorrectResultForASundayWhenReturnTypeIs2()
		{
			var func = new Weekday();
			var result = func.Execute(FunctionsHelper.CreateArgs(new DateTime(2012, 4, 1).ToOADate(), 2), this.ParsingContext);
			Assert.AreEqual(7, result.Result);
		}

		[TestMethod]
		public void WeekdayShouldReturnCorrectResultForASundayWhenReturnTypeIs3()
		{
			var func = new Weekday();
			var result = func.Execute(FunctionsHelper.CreateArgs(new DateTime(2012, 4, 1).ToOADate(), 3), this.ParsingContext);
			Assert.AreEqual(6, result.Result);
		}

		[TestMethod]
		public void WeekdayShouldReturnCorrectResultForASundayWhenReturnTypeIs11()
		{
			var func = new Weekday();
			var result = func.Execute(FunctionsHelper.CreateArgs(new DateTime(2012, 4, 1).ToOADate(), 11), this.ParsingContext);
			Assert.AreEqual(7, result.Result);
		}

		[TestMethod]
		public void WeekdayShouldReturnCorrectResultForASundayWhenReturnTypeIs12()
		{
			var func = new Weekday();
			var result = func.Execute(FunctionsHelper.CreateArgs(new DateTime(2012, 4, 1).ToOADate(), 12), this.ParsingContext);
			Assert.AreEqual(6, result.Result);
		}

		[TestMethod]
		public void WeekdayShouldReturnCorrectResultForASundayWhenReturnTypeIs13()
		{
			var func = new Weekday();
			var result = func.Execute(FunctionsHelper.CreateArgs(new DateTime(2012, 4, 1).ToOADate(), 13), this.ParsingContext);
			Assert.AreEqual(5, result.Result);
		}

		[TestMethod]
		public void WeekdayShouldReturnCorrectResultForASundayWhenReturnTypeIs14()
		{
			var func = new Weekday();
			var result = func.Execute(FunctionsHelper.CreateArgs(new DateTime(2012, 4, 1).ToOADate(), 14), this.ParsingContext);
			Assert.AreEqual(4, result.Result);
		}

		[TestMethod]
		public void WeekdayShouldReturnCorrectResultForASundayWhenReturnTypeIs15()
		{
			var func = new Weekday();
			var result = func.Execute(FunctionsHelper.CreateArgs(new DateTime(2012, 4, 1).ToOADate(), 15), this.ParsingContext);
			Assert.AreEqual(3, result.Result);
		}

		[TestMethod]
		public void WeekdayShouldReturnCorrectResultForASundayWhenReturnTypeIs16()
		{
			var func = new Weekday();
			var result = func.Execute(FunctionsHelper.CreateArgs(new DateTime(2012, 4, 1).ToOADate(), 16), this.ParsingContext);
			Assert.AreEqual(2, result.Result);
		}

		[TestMethod]
		public void WeekdayShouldReturnCorrectResultForASundayWhenReturnTypeIs17()
		{
			var func = new Weekday();
			var result = func.Execute(FunctionsHelper.CreateArgs(new DateTime(2012, 4, 1).ToOADate(), 17), this.ParsingContext);
			Assert.AreEqual(1, result.Result);
		}

		[TestMethod]
		public void WeekdayWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Weekday();

			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WeekdayWithStringArgumentReturnsPoundValue()
		{
			var func = new Weekday();

			var args = FunctionsHelper.CreateArgs("word");
			var result = func.Execute(args, this.ParsingContext);
			
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WeekdayWithNonzeroIntArgumentReturnsThatIntMod7()
		{
			var func = new Weekday();

			var args = FunctionsHelper.CreateArgs(8);
			var result = func.Execute(args, this.ParsingContext);

			Assert.AreEqual(1, result.Result);
		}

		[TestMethod]
		public void WeekdayWithArgumentAsTheNumberZeroReturnsTheNumber7()
		{
			var func = new Weekday();

			var args = FunctionsHelper.CreateArgs(0);
			var result = func.Execute(args, this.ParsingContext);

			Assert.AreEqual(7, result.Result);
		}

		[TestMethod]
		public void WeekdayWithDATEFunctionRepresentationOfSundayArgumentReturnsCorrentResult()
		{
			var func = new Weekday();

			var args = FunctionsHelper.CreateArgs(new DateTime(2017, 5, 14));
			var result = func.Execute(args, this.ParsingContext);

			Assert.AreEqual(1, result.Result);
		}

		[TestMethod]
		public void WeekdayWithInvalidReturnTypeArgumentReturnsPoundNum()
		{
			var func = new Weekday();

			var args = FunctionsHelper.CreateArgs(new DateTime(2017, 5, 14), 4);
			var result = func.Execute(args, this.ParsingContext);

			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WeekdayWithNegativeDateArgumentReturnsPoundNum()
		{
			var func = new Weekday();

			var args = FunctionsHelper.CreateArgs(-1);
			var result = func.Execute(args, this.ParsingContext);

			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WeekdayWithStringAsSecondArgumentReturnsPoundValue()
		{
			var func = new Weekday();

			var args = FunctionsHelper.CreateArgs(new DateTime(2017, 5, 14),"word");
			var result = func.Execute(args, this.ParsingContext);

			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WeekdayWithSecondArgumentNullReturnsPoundNum()
		{
			var func = new Weekday();

			var args = FunctionsHelper.CreateArgs(new DateTime(2017, 5, 14), null);
			var result = func.Execute(args, this.ParsingContext);

			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WeekdayWithFirstArgumentNullReturnsPoundNum()
		{
			var func = new Weekday();

			var args = FunctionsHelper.CreateArgs(null, 3);
			var result = func.Execute(args, this.ParsingContext);

			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void WeekdayWithDateAsStringWithSlashReturnsCorrectResult()
		{
			var func = new Weekday();

			var args = FunctionsHelper.CreateArgs("5/17/2017");
			var result = func.Execute(args, this.ParsingContext);

			Assert.AreEqual(4, result.Result);
		}

		[TestMethod]
		public void WeekdayWithDateNotAsStringReturnsCorrectResult()
		{
			var func = new Weekday();

			var args = FunctionsHelper.CreateArgs(5/17/2017);
			var result = func.Execute(args, this.ParsingContext);

			Assert.AreEqual(7, result.Result);
		}

		[TestMethod]
		public void WeekdayWithDateAsStringWithDashesReturnsCorrectResult()
		{
			var func = new Weekday();

			var args = FunctionsHelper.CreateArgs("5-17-2017");
			var result = func.Execute(args, this.ParsingContext);

			Assert.AreEqual(4, result.Result);
		}
		#endregion
	}
}
