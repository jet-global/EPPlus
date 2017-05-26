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
using EPPlusTest.Excel.Functions.DateTimeFunctions;
using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;

namespace EPPlusTest.FormulaParsing.Excel.Functions.DateTimeFunctions
{
	[TestClass]
	public class TimeTests : DateTimeFunctionsTestBase
	{
		#region Time Function (Execute) Tests
		[TestMethod]
		public void TimeShouldReturnACorrectSerialNumber()
		{
			var expectedResult = this.GetTime(10, 11, 12);
			var function = new Time();
			var result = function.Execute(FunctionsHelper.CreateArgs(10, 11, 12), this.ParsingContext);
			Assert.AreEqual(expectedResult, result.Result);
		}

		[TestMethod]
		public void TimeShouldParseStringCorrectly()
		{
			//Ask Matt about this 
			var expectedResult = this.GetTime(10, 11, 12);
			var function = new Time();
			var result = function.Execute(FunctionsHelper.CreateArgs("10:11:12"), this.ParsingContext);
			Assert.AreEqual(expectedResult, result.Result);
		}

		[TestMethod]
		public void TimeWithInvalidArgumentReturnsPoundValue()
		{
			var function = new Time();

			var args = FunctionsHelper.CreateArgs();
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void TimeWithLastArgNegativeReturnsCorrectValue()
		{
			var function = new Time();
			var args = FunctionsHelper.CreateArgs(10, 10, -10);
			var result = function.Execute(args, this.ParsingContext);
			var expectedResult = this.GetTime(10, 09, 50);
			Assert.AreEqual(expectedResult, result.Result);
		}

		[TestMethod]
		public void TimeWithSecondArgNegativeReturnsCorrectValue()
		{
			var function = new Time();
			var args = FunctionsHelper.CreateArgs(10, -10, 10);
			var result = function.Execute(args, this.ParsingContext);
			var expectedResult = this.GetTime(9, 50, 10);
			Assert.AreEqual(expectedResult, result.Result);
		}

		[TestMethod]
		public void TimeWithLastTwoArgsNegativeReturnsCorrectValue()
		{
			var function = new Time();

			var args = FunctionsHelper.CreateArgs(10, -10, -10);
			var result = function.Execute(args, this.ParsingContext);
			var expectedResult = this.GetTime(9, 49, 50);
			Assert.AreEqual(expectedResult, result.Result);
		}

		[TestMethod]
		public void TimeWithFirstArgNegativeReturnsPoundNum()
		{
			//This test case tests all four cases where the first argument is negative. They all should return #NUM!
			var function = new Time();

			var case1Args = FunctionsHelper.CreateArgs(-10, 10, 10);
			var case2Args = FunctionsHelper.CreateArgs(-10, 10, -10);
			var case3Args = FunctionsHelper.CreateArgs(-10, -10, 10);
			var case4Args = FunctionsHelper.CreateArgs(-10, -10, -10);

			var case1Result = function.Execute(case1Args, this.ParsingContext);
			var case2Result = function.Execute(case2Args, this.ParsingContext);
			var case3Result = function.Execute(case3Args, this.ParsingContext);
			var case4Result = function.Execute(case4Args, this.ParsingContext);

			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)case1Result.Result).Type);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)case2Result.Result).Type);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)case3Result.Result).Type);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)case4Result.Result).Type);
		}

		[TestMethod]
		public void TimeWithMaxTimeInputsReturnsCorrectValue()
		{
			var function = new Time();

			var args = FunctionsHelper.CreateArgs(32767, 32767, 32767);
			var result = function.Execute(args, this.ParsingContext);
			var expectedResult = this.GetTime(10, 13, 7);
			Assert.AreEqual(expectedResult, result.Result);
		}

		[TestMethod]
		public void TimeWithMaxTimeAsFirstInputOnlyReturnsCorrectValue()
		{
			var function = new Time();
			var args = FunctionsHelper.CreateArgs(32767,0,0);
			var result = function.Execute(args, this.ParsingContext);
			var expectedResult = this.GetTime(7, 0, 0);
			Assert.AreEqual(expectedResult, result.Result);
		}

		[TestMethod]
		public void TimeWithMinutesOver59ReturnsCorrectValue()
		{
			var function = new Time();
			var args = FunctionsHelper.CreateArgs(0, 750, 0);
			var result = function.Execute(args, this.ParsingContext);
			var expectedResult = this.GetTime(12, 30, 0);
			Assert.AreEqual(expectedResult, result.Result);
		}

		[TestMethod]
		public void TimeWithSecondsOver59ReturnsCorrecctValue()
		{
			var function = new Time();
			var args = FunctionsHelper.CreateArgs(0, 0, 2000);
			var result = function.Execute(args, this.ParsingContext);
			var expectedResult = this.GetTime(0, 33, 20);
			Assert.AreEqual(expectedResult, result.Result);
		}

		[TestMethod]
		public void TimeWithOnePastMaxTimeInputReturnsPoundNum()
		{
			var function = new Time();

			var args = FunctionsHelper.CreateArgs(32768, 32768, 32768);
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void TimeWithGenericStringOrEmptyStringReturnsPoundValue()
		{
			//This test case tests all three cases where the input is a generic string or an empty string. They should all return #VALUE!
			var function = new Time();

			var case1Args = FunctionsHelper.CreateArgs("string", 10, 10);
			var case2Args = FunctionsHelper.CreateArgs(10, "string", 10);
			var case3Args = FunctionsHelper.CreateArgs(10, 10, "string");
			var case4Args = FunctionsHelper.CreateArgs("", 10, 10);
			var case5Args = FunctionsHelper.CreateArgs(10, "", 10);
			var case6Args = FunctionsHelper.CreateArgs(10, 10, "");

			var case1Result = function.Execute(case1Args, this.ParsingContext);
			var case2Result = function.Execute(case2Args, this.ParsingContext);
			var case3Result = function.Execute(case3Args, this.ParsingContext);
			var case4Result = function.Execute(case4Args, this.ParsingContext);
			var case5Result = function.Execute(case5Args, this.ParsingContext);
			var case6Result = function.Execute(case6Args, this.ParsingContext);

			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)case1Result.Result).Type);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)case2Result.Result).Type);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)case3Result.Result).Type);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)case4Result.Result).Type);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)case5Result.Result).Type);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)case6Result.Result).Type);
		}

		[TestMethod]
		public void TimeWithArgsAsNumericStringReturnsCorrectResult()
		{
			// This test case tests all three cases where the input is a numeric string. 
			var function = new Time();

			var case1Args = FunctionsHelper.CreateArgs("10", 10, 10);
			var case2Args = FunctionsHelper.CreateArgs(10, "10", 10);
			var case3Args = FunctionsHelper.CreateArgs(10, 10, "10");

			var case1Result = function.Execute(case1Args, this.ParsingContext);
			var case2Result = function.Execute(case2Args, this.ParsingContext);
			var case3Result = function.Execute(case3Args, this.ParsingContext);

			var expectedResult = this.GetTime(10, 10, 10);

			Assert.AreEqual(expectedResult, case1Result.Result);
			Assert.AreEqual(expectedResult, case2Result.Result);
			Assert.AreEqual(expectedResult, case3Result.Result);
		}

		[TestMethod]
		public void TimeWithOmittedThirdParamReturnsCorrectResult()
		{
			var function = new Time();
			var args = FunctionsHelper.CreateArgs(10, 10);
			var result = function.Execute(args, this.ParsingContext);
			var exptectedResult = this.GetTime(10, 10, 0);
			Assert.AreEqual(exptectedResult, result.Result);
		}

		[TestMethod]
		public void TimeWithOmittedSecondAndThirdParametersReturnsCorrectResult()
		{
			var function = new Time();
			var args = FunctionsHelper.CreateArgs(10);
			var result = function.Execute(args, this.ParsingContext);
			var expectedResult = this.GetTime(10, 0, 0);
			Assert.AreEqual(expectedResult, result.Result);
		}


		[TestMethod]
		public void MaxTimeInARegularDayReturnsCorrectResult()
		{
			var function = new Time();
			var args = FunctionsHelper.CreateArgs(23, 59, 59);
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(0.999988425925926, result.Result);
		}
		#endregion
	}
}
