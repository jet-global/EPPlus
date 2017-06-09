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
using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math
{
	[TestClass]
	public class ExpTests : MathFunctionsTestBase
	{
		#region
		[TestMethod]
		public void ExpWithTooFewArgumentsReturnsPoundValue()
		{
			var func = new Exp();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void ExpWithNormalInputProducesCorrectResult()
		{
			var func = new Exp();
			var args = FunctionsHelper.CreateArgs(23);
			var result = func.Execute(args, this.ParsingContext);
			var roundedResult = System.Math.Round((double)result.Result, 4);
			Assert.AreEqual(9744803446.2489, roundedResult);
		}

		[TestMethod]
		public void ExpWithValueNearMaxValidValueReturnsCorrectResult()
		{
			var func = new Exp();
			var args = FunctionsHelper.CreateArgs(709);
			var result = func.Execute(args, this.ParsingContext);
			var expectedValue = 8.21840746155497 * System.Math.Pow(10, 307);
			var actualValue = (double)result.Result / System.Math.Pow(10, 307);
			var roundedResult = System.Math.Round(actualValue, 14);
			roundedResult = roundedResult * System.Math.Pow(10, 307);
			//actualValue = System.Math.Truncate(actualValue);
			//Assert.AreEqual(821840746155497, actualValue);
			Assert.AreEqual(expectedValue, roundedResult);
		}

		[TestMethod]
		public void ExpWithPositiveIntegerReturnsCorrectResult()
		{
			var func = new Exp();
			var args = FunctionsHelper.CreateArgs(1);
			var result = func.Execute(args, this.ParsingContext);
			var roundedResult = System.Math.Round((double)result.Result, 13);
			Assert.AreEqual(2.718281828459, roundedResult);
		}

		[TestMethod]
		public void ExpWithFractionReturnsCorrectResult()
		{
			var func = new Exp();
			var args = FunctionsHelper.CreateArgs(0.5);
			var result = func.Execute(args, this.ParsingContext);
			var roundedResult = System.Math.Round((double)result.Result, 14);
			Assert.AreEqual(1.64872127070013, roundedResult);
		}

		[TestMethod]
		public void ExpWithZeroReturnsCorrectResult()
		{
			var func = new Exp();
			var args = FunctionsHelper.CreateArgs(0);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(1.0, result.Result);
		}

		[TestMethod]
		public void ExpWithNegativeIntegerReturnsCorrectResult()
		{
			var func = new Exp();
			var args = FunctionsHelper.CreateArgs(-1);
			var result = func.Execute(args, this.ParsingContext);
			var roundedResult = System.Math.Round((double)result.Result, 15);
			Assert.AreEqual(0.367879441171442, roundedResult);
		}

		[TestMethod]
		public void ExpWithNegativeDoubleReturnsCorrectResult()
		{
			var func = new Exp();
			var args = FunctionsHelper.CreateArgs(-2.5);
			var result = func.Execute(args, this.ParsingContext);
			var roundedResult = System.Math.Round((double)result.Result, 15);
			Assert.AreEqual(0.082084998623899, roundedResult);
		}

		[TestMethod]
		public void ExpWithPositiveIntegerInStringReturnsCorrectResult()
		{
			var func = new Exp();
			var args = FunctionsHelper.CreateArgs("1");
			var result = func.Execute(args, this.ParsingContext);
			var roundedResult = System.Math.Round((double)result.Result, 13);
			Assert.AreEqual(2.718281828459, roundedResult);
		}

		[TestMethod]
		public void ExpWithFractionInStringReturnsCorrectResult()
		{
			var func = new Exp();
			var args = FunctionsHelper.CreateArgs("0.5");
			var result = func.Execute(args, this.ParsingContext);
			var roundedResult = System.Math.Round((double)result.Result, 14);
			Assert.AreEqual(1.64872127070013, roundedResult);
		}

		[TestMethod]
		public void ExpWithZeroInStringReturnsCorrectResult()
		{
			var func = new Exp();
			var args = FunctionsHelper.CreateArgs("0");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(1.0, result.Result);
		}

		[TestMethod]
		public void ExpWithNegativeIntegerInStringReturnsCorrectResult()
		{
			var func = new Exp();
			var args = FunctionsHelper.CreateArgs("-1");
			var result = func.Execute(args, this.ParsingContext);
			var roundedResult = System.Math.Round((double)result.Result, 15);
			Assert.AreEqual(0.367879441171442, roundedResult);
		}

		[TestMethod]
		public void ExpWithNegativeDoubleInStringReturnsCorrectResult()
		{
			var func = new Exp();
			var args = FunctionsHelper.CreateArgs("-2.5");
			var result = func.Execute(args, this.ParsingContext);
			var roundedResult = System.Math.Round((double)result.Result, 15);
			Assert.AreEqual(0.082084998623899, roundedResult);
		}

		[TestMethod]
		public void ExpWithNonNumericStringReturnsPoundValue()
		{
			var func = new Exp();
			var args = FunctionsHelper.CreateArgs("word");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void ExpWithEmptyStringReturnsPoundValue()
		{
			var func = new Exp();
			var args = FunctionsHelper.CreateArgs(string.Empty);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void ExpWithDateInStringReturnsCorrectResult()
		{
			var func = new Exp();
			var args = FunctionsHelper.CreateArgs("3/1/1900");
			var result = func.Execute(args, this.ParsingContext);
			var expectedValue = 3.10429793570192 * System.Math.Pow(10, 26);
			Assert.AreEqual(expectedValue, result.Result);
		}

		[TestMethod]
		public void ExpWithErrorValueInputReturnsThatErrorValue()
		{
			var func = new Exp();
			var args = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.NA));
			var resultNA = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)resultNA.Result).Type);
			args = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Div0));
			var resultDiv0 = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)resultDiv0.Result).Type);
		}

		//[TestMethod]
		//public void Exp()
		//{
		//	var func = new Exp();
		//	var args = FunctionsHelper.CreateArgs();
		//	var result = func.Execute(args, this.ParsingContext);
		//	var roundedResult = System.Math.Round((double)result.Result, );
		//	Assert.AreEqual(, result.Result);
		//}
		#endregion
	}
}
