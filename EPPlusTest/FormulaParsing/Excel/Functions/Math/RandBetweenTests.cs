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
using System.Collections.Generic;
using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math
{
	[TestClass]
	public class RandBetweenTests : MathFunctionsTestBase
	{
		#region RandBetweenTests Function(Execute) Tests
		[TestMethod]
		public void RandBetweenIsChecked1000TimesToMakeSureItStaysBetween0And10()
		{
			var function = new RandBetween();
			var input1 = 0;
			var input2 = 10;
			for (int i = 0; i < 1000; i++)
			{
				var result = function.Execute(FunctionsHelper.CreateArgs(input1, input2), this.ParsingContext);
				if ((double)result.ResultValue < 0 || (double)result.ResultValue > 10)
					Assert.Fail("RAND returned a number outside of the range 0 and 10.");
			}
		}

		[TestMethod]
		public void RandBetweenIsChecked1000TimesToMakeSureItStaysBetweenNegative10And10()
		{
			var function = new RandBetween();
			var input1 = -10;
			var input2 = 10;
			for (int i = 0; i < 1000; i++)
			{
				var result = function.Execute(FunctionsHelper.CreateArgs(input1, input2), this.ParsingContext);
				if ((double) result.ResultValue < -10 || (double) result.ResultValue > 10)
					Assert.Fail("RAND returned a number outside of the range 0 and 10.");
			}
		}

		[TestMethod]
		public void RandBetweenIsGivenABooleanAsAnInputShouldReturnPoundValue()
		{
			var function = new RandBetween();
			var booleanInputTrue = true;
			var booleanInputFalse = false;
			var result1 = function.Execute(FunctionsHelper.CreateArgs(booleanInputTrue, booleanInputTrue), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(booleanInputTrue, booleanInputFalse), this.ParsingContext);
			var result3 = function.Execute(FunctionsHelper.CreateArgs(booleanInputFalse, booleanInputTrue), this.ParsingContext);
			var result4 = function.Execute(FunctionsHelper.CreateArgs(booleanInputFalse, booleanInputFalse), this.ParsingContext);

			Assert.AreEqual(eErrorType.Value, result1.Result);
			Assert.AreEqual(eErrorType.Value, result2.Result);
			Assert.AreEqual(eErrorType.Value, result3.Result);
			Assert.AreEqual(eErrorType.Value, result4.Result);
		}

		[TestMethod]
		public void RandBetweenIsGivenARangeOfDatesSeperatedByDashes()
		{
			var function = new RandBetween();
			var input1 = "1-1-2017 00:00";
			var input2 = "12-31-2017 11:59";
			var result1 = function.Execute(FunctionsHelper.CreateArgs(input1, input2), this.ParsingContext);

			if ((double) result1.ResultValue < 42736 || (double) result1.ResultValue > 43100.49931)
				Assert.Fail("A value outside of the given range was returned.");
		}

		[TestMethod]
		public void RandBetweenIsGivenARangeOfDatesSeperatedBySlashes()
		{
			var function = new RandBetween();
			var input1 = "1/11/2011 11:00 am";
			var input2 = "12/11/2011 11:00 am";
			var result1 = function.Execute(FunctionsHelper.CreateArgs(input1, input2), this.ParsingContext);

			if ((double)result1.ResultValue < 40554.45833 || (double)result1.ResultValue > 40888.45833)
				Assert.Fail("A value outside of the given range was returned.");
		}

		[TestMethod]
		public void RandBetweenIsGivenMidnightTwice()
		{
			var function = new RandBetween();
			var input1 = "12:00 am";
			var input2 = "12:00 am";
			var result1 = function.Execute(FunctionsHelper.CreateArgs(input1, input2), this.ParsingContext);

			Assert.AreEqual(0, result1.ResultNumeric);
		}

		[TestMethod]
		public void RandBetweenIsGivenSixAmAndSixPm()
		{
			var function = new RandBetween();
			var input1 = "6:00 am";
			var input2 = "6:00 pm";
			var result1 = function.Execute(FunctionsHelper.CreateArgs(input1, input2), this.ParsingContext);

			Assert.AreEqual(1, result1.ResultNumeric);
		}

		[TestMethod]
		public void RandBetweenIsGivenSixPmAndSixAm()
		{
			var function = new RandBetween();
			var input1 = "6:00 pm";
			var input2 = "6:00 am";
			var result1 = function.Execute(FunctionsHelper.CreateArgs(input1, input2), this.ParsingContext);

			Assert.AreEqual(eErrorType.Value, result1.Result);
		}

		[TestMethod]
		public void RandBetweenIsGivenARangeOf12HourTimes()
		{
			var function = new RandBetween();
			var input1 = "00:01 am";
			var input2 = "11:59 pm";
			var result1 = function.Execute(FunctionsHelper.CreateArgs(input1, input2), this.ParsingContext);

			Assert.AreEqual(1, result1.ResultNumeric);
		}

		#region RandBetweenTests from the MathFunctionTests.cs

		[TestMethod]
		public void RandBetweenShouldReturnAnIntegerValueBetweenSuppliedValues()
		{
			var func = new RandBetween();
			var args = FunctionsHelper.CreateArgs(1, 5);
			var result = func.Execute(args, ParsingContext);
			CollectionAssert.Contains(new List<double> { 1d, 2d, 3d, 4d, 5d }, result.Result);
		}

		[TestMethod]
		public void RandBetweenShouldReturnAnIntegerValueBetweenSuppliedValuesWhenLowIsNegative()
		{
			var func = new RandBetween();
			var args = FunctionsHelper.CreateArgs(-5, 0);
			var result = func.Execute(args, ParsingContext);
			CollectionAssert.Contains(new List<double> { 0d, -1d, -2d, -3d, -4d, -5d }, result.Result);
		}

		[TestMethod]
		public void RandBetweenWithInvalidArgumentReturnsPoundValue()
		{
			var func = new RandBetween();
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}
#endregion
		#endregion
	}
}