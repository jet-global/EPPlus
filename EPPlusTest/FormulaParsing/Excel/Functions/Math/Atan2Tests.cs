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
	public class Atan2Tests : MathFunctionsTestBase
	{
		#region Atan2Tests Function(Execute) Tests
		[TestMethod]
		public void Atan2IsGivenAlternatingInputsOfOneAndZero()
		{
			var function = new Atan2();

			var input1 = 0;
			var input2 = 1;
			var input3 = -1;

			var result1 = function.Execute(FunctionsHelper.CreateArgs(input1, input1), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(input2, input1), this.ParsingContext);
			var result3 = function.Execute(FunctionsHelper.CreateArgs(input1, input2), this.ParsingContext);
			var result4 = function.Execute(FunctionsHelper.CreateArgs(input2, input2), this.ParsingContext);
			var result5 = function.Execute(FunctionsHelper.CreateArgs(input3, input1), this.ParsingContext);
			var result6 = function.Execute(FunctionsHelper.CreateArgs(input1, input3), this.ParsingContext);
			var result7 = function.Execute(FunctionsHelper.CreateArgs(input3, input3), this.ParsingContext);
			var result8 = function.Execute(FunctionsHelper.CreateArgs(input3, input2), this.ParsingContext);
			var result9 = function.Execute(FunctionsHelper.CreateArgs(input2, input3), this.ParsingContext);

			Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)result1.Result).Type);
			Assert.AreEqual(0, result2.ResultNumeric, .00001);
			Assert.AreEqual(1.570796327, result3.ResultNumeric, .00001);
			Assert.AreEqual(0.785398163, result4.ResultNumeric, .00001);
			Assert.AreEqual(3.141592654, result5.ResultNumeric, .00001);
			Assert.AreEqual(-1.570796327, result6.ResultNumeric, .00001);
			Assert.AreEqual(-2.35619449, result7.ResultNumeric, .00001);
			Assert.AreEqual(2.35619449, result8.ResultNumeric, .00001);
			Assert.AreEqual(-0.785398163, result9.ResultNumeric, .00001);

		}

		[TestMethod]
		public void Atan2HandlesPi()
		{
			var function = new Atan2();
			var Pi = System.Math.PI;

			var input1 = Pi;
			var input2 = 1;
			var input3 = -1;

			var result1 = function.Execute(FunctionsHelper.CreateArgs(input1, input1), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(input1, input2), this.ParsingContext);
			var result3 = function.Execute(FunctionsHelper.CreateArgs(input2, input1), this.ParsingContext);
			var result4 = function.Execute(FunctionsHelper.CreateArgs(input1, input3), this.ParsingContext);

			//Note: Neither Excel or EPPlus handle Pi perfectly. Both seem to have a small rounding issue that is not a problem if you are aware of it.
			Assert.AreEqual(0.785398163, System.Math.Round(result1.ResultNumeric, 9));
			Assert.AreEqual(0.308169071, System.Math.Round(result2.ResultNumeric, 9));
			Assert.AreEqual(1.262627256, System.Math.Round(result3.ResultNumeric, 9));
			Assert.AreEqual(-0.308169071, System.Math.Round(result4.ResultNumeric, 9));
		}

		[TestMethod]
		public void Atan2HandlesLargeInputs()
		{
			var function = new Atan2();

			var input1 = 1000;
			var input2 = -1000;

			var result1 = function.Execute(FunctionsHelper.CreateArgs(input1, input1), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(input2, input2), this.ParsingContext);
			var result3 = function.Execute(FunctionsHelper.CreateArgs(input1, input2), this.ParsingContext);

			Assert.AreEqual(0.785398163, System.Math.Round(result1.ResultNumeric, 9));
			Assert.AreEqual(-2.35619449, System.Math.Round(result2.ResultNumeric, 9));
			Assert.AreEqual(-0.785398163, System.Math.Round(result3.ResultNumeric, 9));
		}

		[TestMethod]
		public void Atan2HandlesStringInputs()
		{
			var function = new Atan2();

			var input1 = "string";
			var input2 = 5;
			var input3 = "five";
			var input4 = "5";

			var result1 = function.Execute(FunctionsHelper.CreateArgs(input1, input1), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(input1, input2), this.ParsingContext);
			var result3 = function.Execute(FunctionsHelper.CreateArgs(input2, input1), this.ParsingContext);
			var result4 = function.Execute(FunctionsHelper.CreateArgs(input3, input3), this.ParsingContext);
			var result5 = function.Execute(FunctionsHelper.CreateArgs(input4, input4), this.ParsingContext);

			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result1.Result).Type);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result2.Result).Type);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result3.Result).Type);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result4.Result).Type);
			Assert.AreEqual(0.785398163, System.Math.Round(result5.ResultNumeric, 9));
		}

		[TestMethod]
		public void Atan2HandlesDoublesAsInputs()
		{
			var function = new Atan2();

			var input1 = 5.5;
			var input2 = "5.5";
			var input3 = -5.5;
			var input4 = "-5.5";

			var result1 = function.Execute(FunctionsHelper.CreateArgs(input1, input1), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(input2, input2), this.ParsingContext);
			var result3 = function.Execute(FunctionsHelper.CreateArgs(input3, input3), this.ParsingContext);
			var result4 = function.Execute(FunctionsHelper.CreateArgs(input4, input4), this.ParsingContext);

			Assert.AreEqual(0.785398163, System.Math.Round(result1.ResultNumeric, 9));
			Assert.AreEqual(0.785398163, System.Math.Round(result2.ResultNumeric, 9));
			Assert.AreEqual(-2.35619449, System.Math.Round(result3.ResultNumeric, 9));
			Assert.AreEqual(-2.35619449, System.Math.Round(result4.ResultNumeric, 9));
		}

		[TestMethod]
		public void Atan2HandlesBooleans()
		{
			var function = new Atan2();

			var input1 = true;
			var input2 = false;
			var input3 = 5;

			var result1 = function.Execute(FunctionsHelper.CreateArgs(input1, input1), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(input1, input2), this.ParsingContext);
			var result3 = function.Execute(FunctionsHelper.CreateArgs(input2, input1), this.ParsingContext);
			var result4 = function.Execute(FunctionsHelper.CreateArgs(input2, input2), this.ParsingContext);
			var result5 = function.Execute(FunctionsHelper.CreateArgs(input1, input3), this.ParsingContext);
			var result6 = function.Execute(FunctionsHelper.CreateArgs(input3, input1), this.ParsingContext);
			var result7 = function.Execute(FunctionsHelper.CreateArgs(input2, input3), this.ParsingContext);
			var result8 = function.Execute(FunctionsHelper.CreateArgs(input3, input2), this.ParsingContext);

			Assert.AreEqual(0.785398163, System.Math.Round(result1.ResultNumeric, 9));
			Assert.AreEqual(0, System.Math.Round(result2.ResultNumeric, 9));
			Assert.AreEqual(1.570796327, System.Math.Round(result3.ResultNumeric, 9));
			Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)result4.Result).Type);
			Assert.AreEqual(1.373400767, System.Math.Round(result5.ResultNumeric, 9));
			Assert.AreEqual(0.19739556, System.Math.Round(result6.ResultNumeric, 9));
			Assert.AreEqual(1.570796327, System.Math.Round(result7.ResultNumeric, 9));
			Assert.AreEqual(0, System.Math.Round(result8.ResultNumeric, 9));
		}

		[TestMethod]
		public void Atan2HandlesDateAndTimeInputs()
		{
			var function = new Atan2();

			var input1 = "1/17/2011 2:00";
			var input2 = "1-17-2017 2:00";
			var input3 = "16:30";
			var input4 = "4:30 pm";

			var result1 = function.Execute(FunctionsHelper.CreateArgs(input1, input1), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(input2, input2), this.ParsingContext);
			var result3 = function.Execute(FunctionsHelper.CreateArgs(input3, input3), this.ParsingContext);
			var result4 = function.Execute(FunctionsHelper.CreateArgs(input4, input4), this.ParsingContext);
			var result5 = function.Execute(FunctionsHelper.CreateArgs(input3, input4), this.ParsingContext);


			Assert.AreEqual(0.785398163, System.Math.Round(result1.ResultNumeric, 9));
			Assert.AreEqual(0.785398163, System.Math.Round(result2.ResultNumeric, 9));
			Assert.AreEqual(0.785398163, System.Math.Round(result3.ResultNumeric, 9));
			Assert.AreEqual(0.785398163, System.Math.Round(result4.ResultNumeric, 9));
			Assert.AreEqual(0.785398163, System.Math.Round(result5.ResultNumeric, 9));

			#endregion
		}
	}
}