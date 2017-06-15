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
	public class SqrtPiTests : MathFunctionsTestBase
	{
		#region SqrtPi Tests
		[TestMethod]
		public void SqrtPiWithTooFewArgumentsReturnsPoundValue()
		{
			var function = new SqrtPi();
			var arguments = FunctionsHelper.CreateArgs();
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SqrtPiWithPositiveIntegerReturnsCorrectResult()
		{
			var function = new SqrtPi();
			var arguments = FunctionsHelper.CreateArgs(4);
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(3.54490770181103, (double)result.Result, 0.00000000000001);
		}

		[TestMethod]
		public void SqrtPiWithPositiveDoubleReturnsCorrectResult()
		{
			var function = new SqrtPi();
			var arguments = FunctionsHelper.CreateArgs(6.25);
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(4.43113462726379, (double)result.Result, 0.00000000000001);
		}

		[TestMethod]
		public void SqrtPiWithZeroReturnsCorrectResult()
		{
			var function = new SqrtPi();
			var arguments = FunctionsHelper.CreateArgs(0);
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(0d, result.Result);
		}

		[TestMethod]
		public void SqrtPiWithNegativeIntegerReturnsPoundNum()
		{
			var function = new SqrtPi();
			var arguments = FunctionsHelper.CreateArgs(-1);
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SqrtPiWithNegativeDoubleReturnsPoundNum()
		{
			var function = new SqrtPi();
			var arguments = FunctionsHelper.CreateArgs(-1.5);
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SqrtPiWithPositiveIntegerInStringReturnsCorrectResult()
		{
			var function = new SqrtPi();
			var arguments = FunctionsHelper.CreateArgs("4");
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(3.54490770181103, (double)result.Result, 0.00000000000001);
		}

		[TestMethod]
		public void SqrtPiWithPositiveDoubleInStringReturnsCorrectResult()
		{
			var function = new SqrtPi();
			var arguments = FunctionsHelper.CreateArgs("6.25");
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(4.43113462726379, (double)result.Result, 0.00000000000001);
		}

		[TestMethod]
		public void SqrtPiWithZeroInStringReturnsCorrectResult()
		{
			var function = new SqrtPi();
			var arguments = FunctionsHelper.CreateArgs("0");
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(0d, result.Result);
		}

		[TestMethod]
		public void SqrtPiWithNegativeIntegerInStringReturnsPoundNum()
		{
			var function = new SqrtPi();
			var arguments = FunctionsHelper.CreateArgs("-1");
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SqrtPiWithNegativeDoubleInStringReturnsPoundNum()
		{
			var function = new SqrtPi();
			var arguments = FunctionsHelper.CreateArgs("-1.5");
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SqrtPiWithNonNumericStringReturnsPoundValue()
		{
			var function = new SqrtPi();
			var arguments = FunctionsHelper.CreateArgs("word");
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SqrtPiWithEmptyStringReturnsPoundValue()
		{
			var function = new SqrtPi();
			var arguments = FunctionsHelper.CreateArgs(string.Empty);
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SqrtPiWithDateInStringReturnsCorrectResult()
		{
			var function = new SqrtPi();
			var arguments = FunctionsHelper.CreateArgs("3/4/1900");
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(14.1796308072441, (double)result.Result, 0.0000000000001);
		}

		[TestMethod]
		public void SqrtPiFunctionWithErrorValuesAsInputReturnsTheInputErrorValue()
		{
			var function = new SqrtPi();
			var argumentNA = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.NA));
			var argumentNAME = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Name));
			var argumentVALUE = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Value));
			var argumentNUM = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Num));
			var argumentDIV0 = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Div0));
			var argumentREF = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Ref));
			var resultNA = function.Execute(argumentNA, this.ParsingContext);
			var resultNAME = function.Execute(argumentNAME, this.ParsingContext);
			var resultVALUE = function.Execute(argumentVALUE, this.ParsingContext);
			var resultNUM = function.Execute(argumentNUM, this.ParsingContext);
			var resultDIV0 = function.Execute(argumentDIV0, this.ParsingContext);
			var resultREF = function.Execute(argumentREF, this.ParsingContext);
			Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)resultNA.Result).Type);
			Assert.AreEqual(eErrorType.Name, ((ExcelErrorValue)resultNAME.Result).Type);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)resultVALUE.Result).Type);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)resultNUM.Result).Type);
			Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)resultDIV0.Result).Type);
			Assert.AreEqual(eErrorType.Ref, ((ExcelErrorValue)resultREF.Result).Type);
		}
		#endregion
	}
}
