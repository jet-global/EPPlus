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
	public class EdateTests : DateTimeFunctionsTestBase
	{
		#region EdateTests Function (Execute) Tests
		[TestMethod]
		public void EdateShouldReturnCorrectResult()
		{
			var func = new Edate();

			var dt1arg = new DateTime(2012, 1, 31).ToOADate();
			var dt2arg = new DateTime(2013, 1, 1).ToOADate();
			var dt3arg = new DateTime(2013, 2, 28).ToOADate();

			var r1 = func.Execute(FunctionsHelper.CreateArgs(dt1arg, 1), this.ParsingContext);
			var r2 = func.Execute(FunctionsHelper.CreateArgs(dt2arg, -1), this.ParsingContext);
			var r3 = func.Execute(FunctionsHelper.CreateArgs(dt3arg, 2), this.ParsingContext);
			var r4 = func.Execute(FunctionsHelper.CreateArgs(dt3arg), this.ParsingContext);
			var dt1 = DateTime.FromOADate((double)r1.Result);
			var dt2 = DateTime.FromOADate((double)r2.Result);
			var dt3 = DateTime.FromOADate((double)r3.Result);

			var exp1 = new DateTime(2012, 2, 29);
			var exp2 = new DateTime(2012, 12, 1);
			var exp3 = new DateTime(2013, 4, 28);

			Assert.AreEqual(exp1, dt1, "dt1 was not " + exp1.ToString("yyyy-MM-dd") + ", but " + dt1.ToString("yyyy-MM-dd"));
			Assert.AreEqual(exp2, dt2, "dt1 was not " + exp2.ToString("yyyy-MM-dd") + ", but " + dt2.ToString("yyyy-MM-dd"));
			Assert.AreEqual(exp3, dt3, "dt1 was not " + exp3.ToString("yyyy-MM-dd") + ", but " + dt3.ToString("yyyy-MM-dd"));
			Assert.AreEqual(eErrorType.Value, (r4.Result as ExcelErrorValue).Type);
		}

		[TestMethod]
		public void EdateWithDateAsStringAsFirstParameterReturnsCorrectResult()
		{
			var date = new DateTime(2017, 5, 22);
			var func = new Edate();
			var args = FunctionsHelper.CreateArgs("5/22/2017", 0);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(date.ToOADate(), result.Result);
		}

		[TestMethod]
		public void EdateWithDateAsLongStringAsFirstParameterReturnsCorrectResult()
		{
			var date = new DateTime(2017, 5, 22);
			var func = new Edate();
			var args = FunctionsHelper.CreateArgs("May 22, 2017", 0);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(date.ToOADate(), result.Result);
		}

		[TestMethod]
		public void EdateWithNonDateStringAsFirstParameterReturnsPoundValue()
		{
			var func = new Edate();
			var args = FunctionsHelper.CreateArgs("word", 0);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}
		#endregion
	}
}
