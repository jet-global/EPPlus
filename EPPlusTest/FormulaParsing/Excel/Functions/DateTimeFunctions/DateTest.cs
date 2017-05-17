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
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace EPPlusTest.FormulaParsing.Excel.Functions.DateTimeFunctions
{
	[TestClass]
	public class DateTest : DateTimeFunctionsTestBase
	{
		#region Date Function (Execute) Tests
		[TestMethod]
		public void DateFunctionShouldReturnADate()
		{
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2012, 4, 3);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(DataType.Date, result.DataType);
		}

		[TestMethod]
		public void DateFunctionShouldReturnACorrectDate()
		{
			var expectedDate = new DateTime(2012, 4, 3);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2012, 4, 3);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateFunctionShouldMonthFromPrevYearIfMonthIsNegative()
		{
			var expectedDate = new DateTime(2011, 11, 3);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2012, -1, 3);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateFunctionShouldMonthFromPrevYearIfMonthAndDayIsNegative()
		{
			var expectedDate = new DateTime(2011, 10, 30);
			var func = new Date();
			var args = FunctionsHelper.CreateArgs(2012, -1, -1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(expectedDate.ToOADate(), result.Result);
		}

		[TestMethod]
		public void DateWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Date();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}
		#endregion
	}
}
