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
	class YearTests : DateTimeFunctionsTestBase
	{
		#region Year Function (Execute) Tests
		[TestMethod]
		public void YearShouldReturnCorrectYear()
		{
			var date = new DateTime(2012, 3, 12);
			var func = new Year();
			var result = func.Execute(FunctionsHelper.CreateArgs(date.ToOADate()), this.ParsingContext);
			Assert.AreEqual(2012, result.Result);
		}

		[TestMethod]
		public void YearShouldReturnCorrectYearWithStringParam()
		{
			var date = new DateTime(2012, 3, 12);
			var func = new Year();
			var result = func.Execute(FunctionsHelper.CreateArgs("2012-03-12"), this.ParsingContext);
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
			}
			finally
			{
				Thread.CurrentThread.CurrentCulture = currentCulture;
			}
		}
		#endregion
	}
}
