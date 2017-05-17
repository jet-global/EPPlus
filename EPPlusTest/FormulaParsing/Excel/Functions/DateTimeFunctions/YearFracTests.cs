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
	public class YearFracTests : DateTimeFunctionsTestBase
	{
		#region YearFrac Function (Execute) Tests
		[TestMethod]
		public void YearFracShouldReturnCorrectResultWithUsBasis()
		{
			var func = new Yearfrac();
			var dt1arg = new DateTime(2013, 2, 28).ToOADate();
			var dt2arg = new DateTime(2013, 3, 31).ToOADate();

			var result = func.Execute(FunctionsHelper.CreateArgs(dt1arg, dt2arg), this.ParsingContext);

			var roundedResult = System.Math.Round((double)result.Result, 4);

			Assert.IsTrue(System.Math.Abs(0.0861 - roundedResult) < double.Epsilon);
		}

		[TestMethod]
		public void YearFracShouldReturnCorrectResultWithEuroBasis()
		{
			var func = new Yearfrac();
			var dt1arg = new DateTime(2013, 2, 28).ToOADate();
			var dt2arg = new DateTime(2013, 3, 31).ToOADate();

			var result = func.Execute(FunctionsHelper.CreateArgs(dt1arg, dt2arg, 4), this.ParsingContext);

			var roundedResult = System.Math.Round((double)result.Result, 4);

			Assert.IsTrue(System.Math.Abs(0.0889 - roundedResult) < double.Epsilon);
		}

		[TestMethod]
		public void YearFracActualActual()
		{
			var func = new Yearfrac();
			var dt1arg = new DateTime(2012, 2, 28).ToOADate();
			var dt2arg = new DateTime(2013, 3, 31).ToOADate();

			var result = func.Execute(FunctionsHelper.CreateArgs(dt1arg, dt2arg, 1), this.ParsingContext);

			var roundedResult = System.Math.Round((double)result.Result, 4);

			Assert.IsTrue(System.Math.Abs(1.0862 - roundedResult) < double.Epsilon);
		}

		[TestMethod]
		public void YearFracTooFewArgumentsReturnsPoundValue()
		{
			var func = new Yearfrac();
			var result = func.Execute(FunctionsHelper.CreateArgs(), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, (result.Result as ExcelErrorValue).Type);
		}
		#endregion
	}
}
