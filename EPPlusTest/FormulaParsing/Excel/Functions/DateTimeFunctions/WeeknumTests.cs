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
		[TestMethod]
		public void WeekNumShouldReturnCorrectResult()
		{
			var func = new Weeknum();
			var dt1 = new DateTime(2012, 12, 31).ToOADate();
			var dt2 = new DateTime(2012, 1, 1).ToOADate();
			var dt3 = new DateTime(2013, 1, 20).ToOADate();

			var r1 = func.Execute(FunctionsHelper.CreateArgs(dt1), this.ParsingContext);
			var r2 = func.Execute(FunctionsHelper.CreateArgs(dt2), this.ParsingContext);
			var r3 = func.Execute(FunctionsHelper.CreateArgs(dt3, 2), this.ParsingContext);
			var r4 = func.Execute(FunctionsHelper.CreateArgs(), this.ParsingContext);
			Assert.AreEqual(53, r1.Result, "r1.Result was not 53, but " + r1.Result.ToString());
			Assert.AreEqual(1, r2.Result, "r2.Result was not 1, but " + r2.Result.ToString());
			Assert.AreEqual(3, r3.Result, "r3.Result was not 3, but " + r3.Result.ToString());
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)r4.Result).Type);
		}
		#endregion
	}
}
