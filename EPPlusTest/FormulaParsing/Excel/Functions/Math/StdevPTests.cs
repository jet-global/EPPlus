﻿/*******************************************************************************
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
using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math
{

	[TestClass]
	public class StdevPTests : MathFunctionsTestBase
	{
		#region StdevP Function(Execute) Tests

		[TestMethod]
		public void StdevpIsGivenBooleanInputs()
		{
			var function = new StdevP();
			var boolInputTrue = true;
			var boolInputFalse = false;
			var result1 = function.Execute(FunctionsHelper.CreateArgs(boolInputTrue, boolInputTrue, boolInputFalse), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(boolInputTrue, boolInputTrue, boolInputTrue), this.ParsingContext);
			var result3 = function.Execute(FunctionsHelper.CreateArgs(boolInputTrue, boolInputFalse, boolInputFalse), this.ParsingContext);

			Assert.AreEqual(0.471404521, result1.ResultNumeric, .00001);
			Assert.AreEqual(0, result2.ResultNumeric, .00001);
			Assert.AreEqual(0.471404521, result3.ResultNumeric, .00001);
		}

		[TestMethod]
		public void StdevpIsGivenBooleanInputsFromCellRefrence()
		{
			var function = new StdevP();

			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = true;
				worksheet.Cells["B2"].Value = false;
				worksheet.Cells["B3"].Formula = "=stdev.p(B1,B1,B2)";
				worksheet.Cells["B4"].Formula = "=stdev.p(B1,B1,B1)";
				worksheet.Cells["B5"].Formula = "=stdev.p(B1,B2,B2)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B3"].Value).Type);
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B4"].Value).Type);
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B5"].Value).Type);
			}
		}

		[TestMethod]
		public void StdevpIsGivenADateSeperatedByABackslash()
		{
			var function = new StdevP();
			var input1 = "1/17/2011 2:00";
			var input2 = "6/17/2011 2:00";
			var input3 = "1/17/2012 2:00";
			var input4 = "1/17/2013 2:00";
			var result1 = function.Execute(FunctionsHelper.CreateArgs(input1, input1, input1), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(input1, input2, input1), this.ParsingContext);
			var result3 = function.Execute(FunctionsHelper.CreateArgs(input1, input3, input4), this.ParsingContext);

			Assert.AreEqual(0, result1.ResultNumeric, .00001);
			Assert.AreEqual(71.18208264, result2.ResultNumeric, .00001);
			Assert.AreEqual(298.4295934, result3.ResultNumeric, .00001);
		}

		[TestMethod]
		public void StdevpIsGivenADateSeperatedByABackslashInputFromCellRefrence()
		{
			var function = new StdevP();

			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "1/17/2011 2:00";
				worksheet.Cells["B2"].Value = "6/17/2011 2:00";
				worksheet.Cells["B3"].Value = "1/17/2012 2:00";
				worksheet.Cells["B4"].Value = "1/17/2013 2:00";
				worksheet.Cells["B5"].Formula = "=stdev.p(B1,B1,B1)";
				worksheet.Cells["B6"].Formula = "=stdev.p(B1,B2,B1)";
				worksheet.Cells["B7"].Formula = "=stdev.p(B1,B3,B4)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B5"].Value).Type);
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B6"].Value).Type);
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B7"].Value).Type);
			}
		}

		[TestMethod]
		public void StdevpIsGivenMaxAmountOfInputs254()
		{
			var function = new StdevP();
			var input1 = 1;

			var result1 = function.Execute(FunctionsHelper.CreateArgs(
				100, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1,
				input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1,
				input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1,
				input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1,
				input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1,
				input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1,
				input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1,
				input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1,
				input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1,
				input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1,
				input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1,
				input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1,
				input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1,
				input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1,
				input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1,
				input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1,
				input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1
				), this.ParsingContext);
			

			Assert.AreEqual(6.199572434, result1.ResultNumeric, .00001);
		}

		[TestMethod]
		public void StdevpIsGivenNumberInputFromCellRefrence()
		{
			var function = new StdevP();

			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 66;
				worksheet.Cells["B2"].Value = 52;
				worksheet.Cells["B3"].Value = 77;
				worksheet.Cells["B4"].Value = 71;
				worksheet.Cells["B5"].Value = 30;
				worksheet.Cells["B6"].Value = 90;
				worksheet.Cells["B7"].Value = 26;
				worksheet.Cells["B8"].Value = 56;
				worksheet.Cells["B9"].Value = 7;
				worksheet.Cells["A10"].Formula = "=stdev.p(B:B)";
				worksheet.Cells["A11"].Formula = "=stdev.p(B1,B3,B5,B6,B9)";
				worksheet.Cells["A12"].Formula = "=stdev.p(B1,B3,B5,B6)";
				worksheet.Calculate();

				Assert.AreEqual(25.43303966, (double)worksheet.Cells["A10"].Value, .00001);
				Assert.AreEqual(30.835045, (double)worksheet.Cells["A11"].Value, .00001);
				Assert.AreEqual(22.32011425, (double)worksheet.Cells["A12"].Value, .00001);
			}
		}

		[TestMethod]
		public void StdevpIsGivenNumbersAsInputs()
		{
			var function = new StdevP();
			var result1 = function.Execute(FunctionsHelper.CreateArgs(-1,0,1), this.ParsingContext);
			Assert.AreEqual(0.816496581, result1.ResultNumeric, .00001);
		}

		[TestMethod]
		public void StdevpIsGivenMilitaryTimesAsInputs()
		{
			var function = new StdevP();
			var result1 = function.Execute(FunctionsHelper.CreateArgs("00:00", "02:00", "13:00"), this.ParsingContext);
			Assert.AreEqual(0.238144836, result1.ResultNumeric, .00001);
		}

		[TestMethod]
		public void StdevpIsGivenMilitaryTimesAsInputsByCellRefrence()
		{
			var function = new StdevP();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "00:00";
				worksheet.Cells["B2"].Value = "02:00";
				worksheet.Cells["B3"].Value = "13:00";
				worksheet.Cells["B4"].Formula = "=stdev.p(B1,B2,B3)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B4"].Value).Type);
			}
		}

		[TestMethod]
		public void StdevpIsGiven12HourTimesAsInputs()
		{
			var function = new StdevP();
			var result1 = function.Execute(FunctionsHelper.CreateArgs("12:00 am", "02:00 am", "01:00 pm"), this.ParsingContext);
			Assert.AreEqual(0.238144836, result1.ResultNumeric, .00001);
		}

		[TestMethod]
		public void StdevpIsGiven12HourTimesAsInputsByCellRefrence()
		{
			var function = new StdevP();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "12:00 am";
				worksheet.Cells["B2"].Value = "02:00 am";
				worksheet.Cells["B3"].Value = "01:00 pm";
				worksheet.Cells["B4"].Formula = "=stdev.p(B1,B2,B3)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B4"].Value).Type);
			}
		}

		[TestMethod]
		public void StdevpIsGivenMonthDayYear12HourTimeAsInputs()
		{
			var function = new StdevP();
			var result1 = function.Execute(FunctionsHelper.CreateArgs("Jan 17, 2011 2:00 am", "June 5, 2017 11:00 pm", "June 15, 2017 11:00 pm"), this.ParsingContext);
			Assert.AreEqual(1101.621004, result1.ResultNumeric, .00001);
		}

		[TestMethod]
		public void StdevpIsGivenMonthDayYear12HourTimeAsInputsByCellRefrence()
		{
			var function = new StdevP();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "Jan 17, 2011 2:00 am";
				worksheet.Cells["B2"].Value = "June 5, 2017 11:00 pm";
				worksheet.Cells["B3"].Value = "June 15, 2017 11:00 pm";
				worksheet.Cells["B4"].Formula = "=stdev.p(B1,B2,B3)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B4"].Value).Type);
			}
		}

		[TestMethod]
		public void StdevpIsGivenDateTimeInputsSeperatedByADashAsInputs()
		{
			var function = new StdevP();
			var result1 = function.Execute(FunctionsHelper.CreateArgs("1-17-2017 2:00", "6-17-2017 2:00", "9-17-2017 2:00"), this.ParsingContext);
			Assert.AreEqual(100.1742926, result1.ResultNumeric, .00001);
		}

		[TestMethod]
		public void StdevpIsGivenDateTimeInputsSeperatedByADashAsInputsByCellRefrence()
		{
			var function = new StdevP();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "1-17-2017 2:00";
				worksheet.Cells["B2"].Value = "6-17-2017 2:00";
				worksheet.Cells["B3"].Value = "9-17-2017 2:00";
				worksheet.Cells["B4"].Formula = "=stdev.p(B1,B2,B3)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B4"].Value).Type);
			}
		}

		[TestMethod]
		public void StdevpIsGivenStringsAsInputs()
		{
			var function = new StdevP();
			var result1 = function.Execute(FunctionsHelper.CreateArgs("string", "another string", "a third string"), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result1.Result).Type);
		}

		[TestMethod]
		public void StdevpIsGivenStringsAsInputsByCellRefrence()
		{
			var function = new StdevP();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "string";
				worksheet.Cells["B2"].Value = "another string";
				worksheet.Cells["B3"].Value = "a third string";
				worksheet.Cells["B4"].Formula = "=stdev.p(B1,B2,B3)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B4"].Value).Type);
			}
		}

		[TestMethod]
		public void StdevpIsGivenStringNumbersAsInputs()
		{
			var function = new StdevP();
			var result1 = function.Execute(FunctionsHelper.CreateArgs("5.5", "6.6", "7.7"), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs("5", "6", "7"), this.ParsingContext);
			Assert.AreEqual(0.898146239, result1.ResultNumeric, .00001);
			Assert.AreEqual(0.816496581, result2.ResultNumeric, .00001);
		}

		[TestMethod]
		public void StdevpIsGivenStringNumbersAsInputsByCellRefrence()
		{
			var function = new StdevP();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["A1"].Value = "5";
				worksheet.Cells["A2"].Value = "6";
				worksheet.Cells["A3"].Value = "7";
				worksheet.Cells["B1"].Value = "5.5";
				worksheet.Cells["B2"].Value = "6.6";
				worksheet.Cells["B3"].Value = "7.7";
				worksheet.Cells["A4"].Formula = "=stdev.p(A1,A2,A3)";
				worksheet.Cells["B4"].Formula = "=stdev.p(B1,B2,B3)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["A4"].Value).Type);
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B4"].Value).Type);
			}
		}
		#endregion
	}
}
