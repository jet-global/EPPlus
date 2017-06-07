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
	public class RadiansTests : MathFunctionsTestBase
	{
		#region Radians Tests
		[TestMethod]
		public void RadiansFunctionWithTooFewArgumentsReturnsPoundValue()
		{
			var func = new Radians();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void RadiansFunctionWithInputZeroDegreesReturnsCorrectResult()
		{
			var func = new Radians();
			var args = FunctionsHelper.CreateArgs(0);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(0.0, result.Result);
		}

		[TestMethod]
		public void RadiansFunctionWithInput90DegreesReturnsCorrectResult()
		{
			var func = new Radians();
			var args = FunctionsHelper.CreateArgs(90);
			var result = func.Execute(args, this.ParsingContext);
			var roundedResult = System.Math.Round((double)result.Result, 14);
			Assert.AreEqual(1.5707963267949, roundedResult);
		}

		[TestMethod]
		public void RadiansFunctionWithInput360DegreesReturnsCorrectResult()
		{
			var func = new Radians();
			var args = FunctionsHelper.CreateArgs(360);
			var result = func.Execute(args, this.ParsingContext);
			var roundedResult = System.Math.Round((double)result.Result, 14);
			Assert.AreEqual(6.28318530717959, roundedResult);
		}

		[TestMethod]
		public void RadiansFunctionWithNegativeIntegerReturnsCorrectResult()
		{
			var func = new Radians();
			var args = FunctionsHelper.CreateArgs(-1);
			var result = func.Execute(args, this.ParsingContext);
			var roundedResult = System.Math.Round((double)result.Result, 15);
			Assert.AreEqual(-0.017453292519943, roundedResult);
		}

		[TestMethod]
		public void RadiansFunctionWithNegativeDoubleReturnsCorrectResult()
		{
			var func = new Radians();
			var args = FunctionsHelper.CreateArgs(-1.5);
			var result = func.Execute(args, this.ParsingContext);
			var roundedResult = System.Math.Round((double)result.Result, 15);
			Assert.AreEqual(-0.026179938779915, roundedResult);
		}

		[TestMethod]
		public void RadiansFunctionWithValueGreaterThan360DegreesReturnsCorrectResult()
		{
			var func = new Radians();
			var args = FunctionsHelper.CreateArgs(720);
			var result = func.Execute(args, this.ParsingContext);
			var roundedResult = System.Math.Round((double)result.Result, 13);
			Assert.AreEqual(12.5663706143592, roundedResult);
		}

		[TestMethod]
		public void RadiansFunctionWithIntegerInStringReturnsCorrectResult()
		{
			var func = new Radians();
			var args = FunctionsHelper.CreateArgs("1");
			var result = func.Execute(args, this.ParsingContext);
			var roundedResult = System.Math.Round((double)result.Result, 15);
			Assert.AreEqual(0.017453292519943, roundedResult);
		}

		[TestMethod]
		public void RadiansFunctionWithDoubleInStringReturnsCorrectResult()
		{
			var func = new Radians();
			var args = FunctionsHelper.CreateArgs("1.5");
			var result = func.Execute(args, this.ParsingContext);
			var roundedResult = System.Math.Round((double)result.Result, 15);
			Assert.AreEqual(0.026179938779915, roundedResult);
		}

		[TestMethod]
		public void RadiansFunctionWithNegativeIntegerInStringReturnsCorrectResult()
		{
			var func = new Radians();
			var args = FunctionsHelper.CreateArgs("-1");
			var result = func.Execute(args, this.ParsingContext);
			var roundedResult = System.Math.Round((double)result.Result, 15);
			Assert.AreEqual(-0.017453292519943, roundedResult);
		}

		[TestMethod]
		public void RadiansFunctionWithNegativeDoubleInStringReturnsCorrectResult()
		{
			var func = new Radians();
			var args = FunctionsHelper.CreateArgs("-1.5");
			var result = func.Execute(args, this.ParsingContext);
			var roundedResult = System.Math.Round((double)result.Result, 15);
			Assert.AreEqual(-0.026179938779915, roundedResult);
		}

		[TestMethod]
		public void RadiansFunctionWithNonNumericStringReturnsPoundValue()
		{
			var func = new Radians();
			var args = FunctionsHelper.CreateArgs("word");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void RadiansFunctionWithEmptyStringReturnsPoundValue()
		{
			var func = new Radians();
			var args = FunctionsHelper.CreateArgs(string.Empty);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void RadiansFunctionWithDateReturnsCorrectResult()
		{
			var func = new Radians();
			var args = FunctionsHelper.CreateArgs("6/7/2017");
			var result = func.Execute(args, this.ParsingContext);
			var roundedResult = System.Math.Round((double)result.Result, 12);
			Assert.AreEqual(748.624076057928, roundedResult);
		}

		//[TestMethod]
		//public void RadiansFunction()
		//{
		//	var func = new Radians();
		//	var args = FunctionsHelper.CreateArgs();
		//	var result = func.Execute(args, this.ParsingContext);
		//	var roundedResult = System.Math.Round((double)result.Result, );
		//	Assert.AreEqual(, roundedResult);
		//}
		#endregion
	}
}
