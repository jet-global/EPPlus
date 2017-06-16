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
	public class DegreesTests : MathFunctionsTestBase
	{
		#region Degrees Tests
		[TestMethod]
		public void DegreesFunctionWithTooFewArgumentsReturnsPoundValue()
		{
			var func = new Degrees();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void DegreesFunctionWithInputZeroRadiansReturnsCorrectResult()
		{
			var func = new Degrees();
			var args = FunctionsHelper.CreateArgs(0);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(0.0, result.Result);
		}

		[TestMethod]
		public void DegreesFunctionWithInputPiOver2RadiansReturnsCorrectResult()
		{
			var func = new Degrees();
			var pi = System.Math.PI;
			var args = FunctionsHelper.CreateArgs(pi/2);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(90.0, result.Result);
		}

		[TestMethod]
		public void DegreesFunctionWithInput2TimesPiRadiansReturnsCorrectResult()
		{
			var func = new Degrees();
			var pi = System.Math.PI;
			var args = FunctionsHelper.CreateArgs(2*pi);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(360.0, result.Result);
		}

		[TestMethod]
		public void DegreesFunctionWithNegativeIntegerReturnsCorrectResult()
		{
			var func = new Degrees();
			var args = FunctionsHelper.CreateArgs(-1);
			var result = func.Execute(args, this.ParsingContext);
			var roundedResult = System.Math.Round((double)result.Result, 13);
			Assert.AreEqual(-57.2957795130823, roundedResult);
		}

		[TestMethod]
		public void DegreesFunctionWithNegativeDoubleReturnsCorrectResult()
		{
			var func = new Degrees();
			var args = FunctionsHelper.CreateArgs(-1.5);
			var result = func.Execute(args, this.ParsingContext);
			var roundedResult = System.Math.Round((double)result.Result, 13);
			Assert.AreEqual(-85.9436692696235, roundedResult);
		}

		[TestMethod]
		public void DegreesFunctionWithValueGreaterThan2PiReturnsCorrectResult()
		{
			var func = new Degrees();
			var args = FunctionsHelper.CreateArgs(10);
			var result = func.Execute(args, this.ParsingContext);
			var roundedResult = System.Math.Round((double)result.Result, 12);
			Assert.AreEqual(572.957795130823, roundedResult);
		}

		[TestMethod]
		public void DegreesFunctionWithIntegerInStringReturnsCorrectResult()
		{
			var func = new Degrees();
			var args = FunctionsHelper.CreateArgs("1");
			var result = func.Execute(args, this.ParsingContext);
			var roundedResult = System.Math.Round((double)result.Result, 13);
			Assert.AreEqual(57.2957795130823, roundedResult);
		}

		[TestMethod]
		public void DegreesFunctionWithDoubleInStringReturnsCorrectResult()
		{
			var func = new Degrees();
			var args = FunctionsHelper.CreateArgs("1.5");
			var result = func.Execute(args, this.ParsingContext);
			var roundedResult = System.Math.Round((double)result.Result, 13);
			Assert.AreEqual(85.9436692696235, roundedResult);
		}

		[TestMethod]
		public void DegreesFunctionWithNegativeIntegerInStringReturnsCorrectResult()
		{
			var func = new Degrees();
			var args = FunctionsHelper.CreateArgs("-1");
			var result = func.Execute(args, this.ParsingContext);
			var roundedResult = System.Math.Round((double)result.Result, 13);
			Assert.AreEqual(-57.2957795130823, roundedResult);
		}

		[TestMethod]
		public void DegreesFunctionWithNegativeDoubleInStringReturnsCorrectResult()
		{
			var func = new Degrees();
			var args = FunctionsHelper.CreateArgs("-1.5");
			var result = func.Execute(args, this.ParsingContext);
			var roundedResult = System.Math.Round((double)result.Result, 13);
			Assert.AreEqual(-85.9436692696235, roundedResult);
		}

		[TestMethod]
		public void DegreesFunctionWithNonNumericStringReturnsPoundValue()
		{
			var func = new Degrees();
			var args = FunctionsHelper.CreateArgs("word");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void DegreesFunctionWithEmptyStringReturnsPoundValue()
		{
			var func = new Degrees();
			var args = FunctionsHelper.CreateArgs(string.Empty);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void DegreesFunctionWithDateInStringReturnsCorrectResult()
		{
			var func = new Degrees();
			var args = FunctionsHelper.CreateArgs("6/7/2017");
			var result = func.Execute(args, this.ParsingContext);
			var roundedResult = System.Math.Round((double)result.Result, 8);
			Assert.AreEqual(2457587.87065464, roundedResult);
		}

		[TestMethod]
		public void DegreesFunctionWithErrorValuesAsInputReturnsTheInputErrorValue()
		{
			var func = new Degrees();
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
