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
	public class SqrtTests : MathFunctionsTestBase
	{
		#region
		[TestMethod]
		public void SqrtFunctionWithTooFewArgumentsReturnsPoundValue()
		{
			var function = new Sqrt();
			var arguments = FunctionsHelper.CreateArgs();
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SqrtFunctionWithPositiveIntegerReturnsCorrectResult()
		{
			var function = new Sqrt();
			var arguments = FunctionsHelper.CreateArgs(4);
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(2d, result.Result);
		}

		[TestMethod]
		public void SqrtFunctionWithPositiveDoubleReturnsCorrectResult()
		{
			var function = new Sqrt();
			var arguments = FunctionsHelper.CreateArgs(6.25);
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(2.5, result.Result);
		}

		[TestMethod]
		public void SqrtFunctionWithZeroReturnsCorrectResult()
		{
			var function = new Sqrt();
			var arguments = FunctionsHelper.CreateArgs(0);
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(0, result.Result);
		}

		[TestMethod]
		public void SqrtFunctionWithNegativeIntegerReturnsPoundNum()
		{
			var function = new Sqrt();
			var arguments = FunctionsHelper.CreateArgs(-1);
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SqrtFunctionWithNegativeDoubleReturnsPoundNum()
		{
			var function = new Sqrt();
			var arguments = FunctionsHelper.CreateArgs(-1.5);
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SqrtFunctionWithPositiveIntegerInStringReturnsCorrectResult()
		{
			var function = new Sqrt();
			var arguments = FunctionsHelper.CreateArgs("4");
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(2d, result.Result);
		}

		[TestMethod]
		public void SqrtFunctionWithPositiveDoubleInStringReturnsCorrectResult()
		{
			var function = new Sqrt();
			var arguments = FunctionsHelper.CreateArgs("6.25");
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(2.5, result.Result);
		}

		[TestMethod]
		public void SqrtFunctionWithZeroInStringReturnsCorrectResult()
		{
			var function = new Sqrt();
			var arguments = FunctionsHelper.CreateArgs("0");
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(0, result.Result);
		}

		[TestMethod]
		public void SqrtFunctionWithNegativeIntegerInStringReturnsPoundNum()
		{
			var function = new Sqrt();
			var arguments = FunctionsHelper.CreateArgs("-1");
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SqrtFunctionWithNegativeDoubleInStringReturnsPoundNum()
		{
			var function = new Sqrt();
			var arguments = FunctionsHelper.CreateArgs("-1.5");
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SqrtFunctionWithNonNumericsStringReturnsPoundValue()
		{
			var function = new Sqrt();
			var arguments = FunctionsHelper.CreateArgs("word");
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SqrtFunctionWithEmptyStringReturnsPoundValue()
		{
			var function = new Sqrt();
			var arguments = FunctionsHelper.CreateArgs(string.Empty);
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SqrtFunctionWithDateInStringReturnsCorrectResult()
		{
			var function = new Sqrt();
			var arguments = FunctionsHelper.CreateArgs("3/4/1900");
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(8d, result.Result);
		}

		[TestMethod]
		public void SqrtFunctionWithErrorValuesAsInputReturnsTheInputErrorValue()
		{
			var func = new Sqrt();
			var argNA = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.NA));
			var argNAME = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Name));
			var argVALUE = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Value));
			var argNUM = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Num));
			var argDIV0 = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Div0));
			var argREF = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Ref));
			var resultNA = func.Execute(argNA, this.ParsingContext);
			var resultNAME = func.Execute(argNAME, this.ParsingContext);
			var resultVALUE = func.Execute(argVALUE, this.ParsingContext);
			var resultNUM = func.Execute(argNUM, this.ParsingContext);
			var resultDIV0 = func.Execute(argDIV0, this.ParsingContext);
			var resultREF = func.Execute(argREF, this.ParsingContext);
			Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)resultNA.Result).Type);
			Assert.AreEqual(eErrorType.Name, ((ExcelErrorValue)resultNAME.Result).Type);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)resultVALUE.Result).Type);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)resultNUM.Result).Type);
			Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)resultDIV0.Result).Type);
			Assert.AreEqual(eErrorType.Ref, ((ExcelErrorValue)resultREF.Result).Type);
		}

		//[TestMethod]
		//public void SqrtFunctionWith()
		//{
		//	var function = new Sqrt();
		//	var arguments = FunctionsHelper.CreateArgs();
		//	var result = function.Execute(arguments, this.ParsingContext);
		//	Assert.AreEqual(, result.Result);
		//}
		#endregion
	}
}
