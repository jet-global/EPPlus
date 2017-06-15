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
	public class FactDoubleTests : MathFunctionsTestBase
	{
		#region FactDouble Tests
		[TestMethod]
		public void FactDoubleWithTooFewArgumentsReturnsPoundValue()
		{
			var func = new FactDouble();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void FactDoubleWithInputAsOneReturnsCorrectResult()
		{
			var func = new FactDouble();
			var args = FunctionsHelper.CreateArgs(1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(1d, result.Result);
		}

		[TestMethod]
		public void FactDoubleWithInputAsTwoReturnsCorrectResult()
		{
			var func = new FactDouble();
			var args = FunctionsHelper.CreateArgs(2);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(2d, result.Result);
		}

		[TestMethod]
		public void FactDoubleWithInputAsThreeReturnsCorrectResult()
		{
			var func = new FactDouble();
			var args = FunctionsHelper.CreateArgs(3);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(3d, result.Result);
		}

		[TestMethod]
		public void FactDoubleWithPositiveEvenIntegerReturnsCorrectResult()
		{
			var func = new FactDouble();
			var args = FunctionsHelper.CreateArgs(4);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(8d, result.Result);
		}

		[TestMethod]
		public void FactDoubleWithPositiveOddIntegerReturnsCorrectResult()
		{
			var func = new FactDouble();
			var args = FunctionsHelper.CreateArgs(5);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(15d, result.Result);
		}

		[TestMethod]
		public void FactDoubleWithPositiveEvenDoubleReturnsCorrectResult()
		{
			var func = new FactDouble();
			var args = FunctionsHelper.CreateArgs(4.9);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(8d, result.Result);
		}

		[TestMethod]
		public void FactDoubleWithPositiveOddDoubleReturnsCorrectResult()
		{
			var func = new FactDouble();
			var args = FunctionsHelper.CreateArgs(5.9);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(15d, result.Result);
		}

		[TestMethod]
		public void FactDoubleWithInputAsZeroReturnsCorrectResult()
		{
			var func = new FactDouble();
			var args = FunctionsHelper.CreateArgs(0);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(1d, result.Result);
		}

		[TestMethod]
		public void FactDoubleWithNegativeFractionReturnsCorrectResult()
		{
			var func = new FactDouble();
			var args = FunctionsHelper.CreateArgs(-0.5);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(1d, result.Result);
		}

		[TestMethod]
		public void FactDoubleWithNegativeOneReturnsCorrectResult()
		{
			var func = new FactDouble();
			var args = FunctionsHelper.CreateArgs(-1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(1d, result.Result);
		}

		[TestMethod]
		public void FactDoubleWithNegativeDoubleLessThanNegativeOneReturnsPoundNum()
		{
			var func = new FactDouble();
			var args = FunctionsHelper.CreateArgs(-1.5);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void FactDoubleWithNegativeIntegerLessThenNegativeOneReturnsPoundNum()
		{
			var func = new FactDouble();
			var args = FunctionsHelper.CreateArgs(-2);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void FactDoubleWithPositiveEvenIntegerInStringReturnsCorrectResult()
		{
			var func = new FactDouble();
			var args = FunctionsHelper.CreateArgs("4");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(8d, result.Result);
		}

		[TestMethod]
		public void FactDoubleWithPositiveOddIntegerInStringReturnsCorrectResult()
		{
			var func = new FactDouble();
			var args = FunctionsHelper.CreateArgs("5");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(15d, result.Result);
		}

		[TestMethod]
		public void FactDoubleWithPositiveEvenDoubleInStringReturnsCorrectResult()
		{
			var func = new FactDouble();
			var args = FunctionsHelper.CreateArgs("4.9");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(8d, result.Result);
		}

		[TestMethod]
		public void FactDoubleWithPositiveOddDoubleInStringReturnsCorrectResult()
		{
			var func = new FactDouble();
			var args = FunctionsHelper.CreateArgs("5.9");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(15d, result.Result);
		}

		[TestMethod]
		public void FactDoubleWithZeroInStringReturnsCorrectResult()
		{
			var func = new FactDouble();
			var args = FunctionsHelper.CreateArgs("0");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(1d, result.Result);
		}

		[TestMethod]
		public void FactDoubleWithNegativeOneInStringReturnsCorrectResult()
		{
			var func = new FactDouble();
			var args = FunctionsHelper.CreateArgs("-1");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(1d, result.Result);
		}

		[TestMethod]
		public void FactDoubleWithNegativeIntegerLessThanNegativeOneInStringReturnsPoundNum()
		{
			var func = new FactDouble();
			var args = FunctionsHelper.CreateArgs("-2");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void FactDoubleWithNegativeDoubleLessThanNegativeOneInStringReturnsPoundNum()
		{
			var func = new FactDouble();
			var args = FunctionsHelper.CreateArgs("-2.5");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void FactDoubleWithNonNumericStringReturnsPoundValue()
		{
			var func = new FactDouble();
			var args = FunctionsHelper.CreateArgs("word");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void FactDoubleWithEmptyStringReturnsPoundValue()
		{
			var func = new FactDouble();
			var args = FunctionsHelper.CreateArgs(string.Empty);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void FactDoubleWithDateInStringReturnsCorrectResult()
		{
			var func = new FactDouble();
			var args = FunctionsHelper.CreateArgs("1/1/1900");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(1d, result.Result);
		}

		[TestMethod]
		public void FactDoubleFunctionWithErrorValuesAsInputReturnsTheInputErrorValue()
		{
			var func = new FactDouble();
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
		#endregion
	}
}
