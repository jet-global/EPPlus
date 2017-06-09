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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math
{
	[TestClass]
	public class CosTests : MathFunctionsTestBase
	{
		#region TimeValue Function(Execute) Tests
		[TestMethod]
		public void CosIsGivenAStringAsInput()
		{
			var function = new Cos();

			var input1 = "string";
			var input2 = "0";
			var input3 = "1";

			var result1 = function.Execute(FunctionsHelper.CreateArgs(input1), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(input2), this.ParsingContext);
			var result3 = function.Execute(FunctionsHelper.CreateArgs(input3), this.ParsingContext);

			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result1.Result).Type);
			Assert.AreEqual(1, result2.ResultNumeric);
			Assert.AreEqual(0.540302306, System.Math.Round(result3.ResultNumeric, 9));
		}

		[TestMethod]
		public void CosIsGivenValuesRanginFromNegative10to10()
		{
			var function = new Cos();

			var input1 = -10;
			var input2 = -1;
			var input3 = 0;
			var input4 = 1;
			var input5 = 10;

			var result1 = function.Execute(FunctionsHelper.CreateArgs(input1), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(input2), this.ParsingContext);
			var result3 = function.Execute(FunctionsHelper.CreateArgs(input3), this.ParsingContext);
			var result4 = function.Execute(FunctionsHelper.CreateArgs(input4), this.ParsingContext);
			var result5 = function.Execute(FunctionsHelper.CreateArgs(input5), this.ParsingContext);

			Assert.AreEqual(-0.839071529, System.Math.Round(result1.ResultNumeric, 9));
			Assert.AreEqual(0.540302306, System.Math.Round(result2.ResultNumeric, 9));
			Assert.AreEqual(1, result3.ResultNumeric);
			Assert.AreEqual(0.540302306, System.Math.Round(result4.ResultNumeric, 9));
			Assert.AreEqual(-0.839071529, System.Math.Round(result5.ResultNumeric, 9));

		}

		[TestMethod]
		public void CosIntAndDoublesAsInputs()
		{
			var function = new Cos();

			var input1 = 20;
			var input2 = 100;
			var input3 = 1;
			var input4 = 1.0;
			var input5 = 1.5;
			var input6 = 1000;

			var result1 = function.Execute(FunctionsHelper.CreateArgs(input1), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(input2), this.ParsingContext);
			var result3 = function.Execute(FunctionsHelper.CreateArgs(input3), this.ParsingContext);
			var result4 = function.Execute(FunctionsHelper.CreateArgs(input4), this.ParsingContext);
			var result5 = function.Execute(FunctionsHelper.CreateArgs(input5), this.ParsingContext);
			var result6 = function.Execute(FunctionsHelper.CreateArgs(input6), this.ParsingContext);

			Assert.AreEqual(0.408082062, System.Math.Round(result1.ResultNumeric, 9));
			Assert.AreEqual(0.862318872, System.Math.Round(result2.ResultNumeric, 9));
			Assert.AreEqual(0.540302306, System.Math.Round(result3.ResultNumeric, 9));
			Assert.AreEqual(0.540302306, System.Math.Round(result4.ResultNumeric, 9));
			Assert.AreEqual(0.070737202, System.Math.Round(result5.ResultNumeric, 9));
			Assert.AreEqual(0.562379076, System.Math.Round(result6.ResultNumeric, 9));

		}

		[TestMethod]
		public void CosHandlesPi()
		{
			var function = new Cos();
			var Pi = System.Math.PI;

			var input1 = Pi;
			var input2 = Pi/2;
			var input3 = 2*Pi;
			var input4 = 60*Pi/180;

			var result1 = function.Execute(FunctionsHelper.CreateArgs(input1), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(input2), this.ParsingContext);
			var result3 = function.Execute(FunctionsHelper.CreateArgs(input3), this.ParsingContext);
			var result4 = function.Execute(FunctionsHelper.CreateArgs(input4), this.ParsingContext);

			Assert.AreEqual(-1, result1.ResultNumeric);
			Assert.AreEqual(6.12303176911189E-17, result2.ResultNumeric,0.000001);//Neither Excel or EPPlus return 0.
			Assert.AreEqual(1, result3.ResultNumeric);
			Assert.AreEqual(0.5, result4.ResultNumeric, .000001);


		}

		#endregion
	}
}
