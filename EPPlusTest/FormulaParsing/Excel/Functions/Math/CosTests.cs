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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;

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

		[TestMethod]
		public void CosHandlesMilitaryTime()
		{
			var function = new Cos();

			var input1 = "00:00";
			var input2 = "00:01";
			var input4 = "23:59:59";

			var result1 = function.Execute(FunctionsHelper.CreateArgs(input1), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(input2), this.ParsingContext);
			var result4 = function.Execute(FunctionsHelper.CreateArgs(input4), this.ParsingContext);

			Assert.AreEqual(1, result1.ResultNumeric);
			Assert.AreEqual(0.999999759, System.Math.Round(result2.ResultNumeric, 8), .000001);
			Assert.AreEqual(0.540312045, System.Math.Round(result4.ResultNumeric, 9), .000001);
		}

		[TestMethod]
		public void CosHandlesMilitaryTimesPast2400()
		{
			var function = new Cos();

			var input2 = "01:00";
			var input4 = "02:00";

			var result2 = function.Execute(FunctionsHelper.CreateArgs(input2), this.ParsingContext);
			var result4 = function.Execute(FunctionsHelper.CreateArgs(input4), this.ParsingContext);

			Assert.AreEqual(0.99913207, System.Math.Round(result2.ResultNumeric, 9));
			Assert.AreEqual(0.996529787, System.Math.Round(result4.ResultNumeric, 9));
		}

		[TestMethod]
		public void CosHandlesDateTimeInputs()
		{
			var function = new Cos();

			var input1 = "1/17/2011 2:00";
			var input2 = "1/17/2011 2:00 AM";
			var input3 = "17/1/2011 2:00 AM";
			var input4 = "17/Jan/2011 2:00 AM";

			var result1 = function.Execute(FunctionsHelper.CreateArgs(input1), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(input2), this.ParsingContext);
			var result3 = function.Execute(FunctionsHelper.CreateArgs(input3), this.ParsingContext);
			var result4 = function.Execute(FunctionsHelper.CreateArgs(input4), this.ParsingContext);

			Assert.AreEqual(-0.523862501, System.Math.Round(result1.ResultNumeric, 9));
			Assert.AreEqual(-0.523862501, System.Math.Round(result2.ResultNumeric, 9));
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result3.Result).Type);
			Assert.AreEqual(-0.523862501, System.Math.Round(result4.ResultNumeric, 9));
		}

		[TestMethod]
		public void CosHandlesNormal12HourClockInputs()
		{
			var function = new Cos();

			var input1 = "00:00:00 AM";
			var input2 = "00:01:32 AM";
			var input3 = "12:00 PM";
			var input4 = "12:00 AM";
			var input6 = "1:00 PM";
			var input8 = "1:10:32 am";
			var input9 = "3:42:32 pm";

			var result1 = function.Execute(FunctionsHelper.CreateArgs(input1), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(input2), this.ParsingContext);
			var result3 = function.Execute(FunctionsHelper.CreateArgs(input3), this.ParsingContext);
			var result4 = function.Execute(FunctionsHelper.CreateArgs(input4), this.ParsingContext);
			var result6 = function.Execute(FunctionsHelper.CreateArgs(input6), this.ParsingContext);
			var result8 = function.Execute(FunctionsHelper.CreateArgs(input8), this.ParsingContext);
			var result9 = function.Execute(FunctionsHelper.CreateArgs(input9), this.ParsingContext);

			Assert.AreEqual(1, result1.ResultNumeric);
			Assert.AreEqual(0.999999433, System.Math.Round(result2.ResultNumeric, 9));
			Assert.AreEqual(0.877582562, result3.ResultNumeric, .00001);
			Assert.AreEqual(1, result4.ResultNumeric);
			Assert.AreEqual(0.856850597, System.Math.Round(result6.ResultNumeric, 9));
			Assert.AreEqual(0.998800647, System.Math.Round(result8.ResultNumeric, 9));
			Assert.AreEqual(0.793329861, System.Math.Round(result9.ResultNumeric, 9));
		}

		[TestMethod]
		public void CosTestMilitaryTimeAndNormalTimeComparisions()
		{
			var function = new Cos();

			var input1 = "16:30";
			var input2 = "04:30 pm";
			var input3 = "02:30";
			var input4 = "2:30 am";

			var result1 = function.Execute(FunctionsHelper.CreateArgs(input1), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(input2), this.ParsingContext);
			var result3 = function.Execute(FunctionsHelper.CreateArgs(input3), this.ParsingContext);
			var result4 = function.Execute(FunctionsHelper.CreateArgs(input4), this.ParsingContext);

			Assert.AreEqual(0.772834946, System.Math.Round(result1.ResultNumeric, 9));
			Assert.AreEqual(0.772834946, System.Math.Round(result2.ResultNumeric, 9));
			Assert.AreEqual(0.994579557, System.Math.Round(result3.ResultNumeric, 9));
			Assert.AreEqual(0.994579557, System.Math.Round(result4.ResultNumeric, 9));
		}

		[TestMethod]
		public void CosTestInputsWithDatesThatHaveSlashesInThem()
		{
			var function = new Cos();

			var input1 = "1/17/2011 2:00 am";
			var input2 = "17/01/2011 2:00 AM";
			var input3 = "17/Jan/2011 2:00 AM";
			var input4 = "17/January/2011 2:00 am";
			var input5 = "1/17/2011 2:00:00 am";
			var input6 = "17/01/2011 2:00:00 AM";
			var input7 = "17/Jan/2011 2:00:00 AM";
			var input8 = "17/January/2011 2:00:00 am";

			var result1 = function.Execute(FunctionsHelper.CreateArgs(input1), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(input2), this.ParsingContext);
			var result3 = function.Execute(FunctionsHelper.CreateArgs(input3), this.ParsingContext);
			var result4 = function.Execute(FunctionsHelper.CreateArgs(input4), this.ParsingContext);
			var result5 = function.Execute(FunctionsHelper.CreateArgs(input5), this.ParsingContext);
			var result6 = function.Execute(FunctionsHelper.CreateArgs(input6), this.ParsingContext);
			var result7 = function.Execute(FunctionsHelper.CreateArgs(input7), this.ParsingContext);
			var result8 = function.Execute(FunctionsHelper.CreateArgs(input8), this.ParsingContext);

			Assert.AreEqual(-0.523862501, System.Math.Round(result1.ResultNumeric, 9));
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result2.Result).Type);
			Assert.AreEqual(-0.523862501, System.Math.Round(result3.ResultNumeric, 9));
			Assert.AreEqual(-0.523862501, System.Math.Round(result4.ResultNumeric, 9));
			Assert.AreEqual(-0.523862501, System.Math.Round(result5.ResultNumeric, 9));
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result6.Result).Type);
			Assert.AreEqual(-0.523862501, System.Math.Round(result7.ResultNumeric, 9));
			Assert.AreEqual(-0.523862501, System.Math.Round(result8.ResultNumeric, 9));
		}

		[TestMethod]
		public void CosHandlesInputsWithDatesInTheFormMonthDateCommaYearTime()
		{
			var function = new Cos();

			var input1 = "Jan 17, 2011 2:00 am";
			var input2 = "June 5, 2017 11:00 pm";
			var input3 = "Jan 17, 2011 2:00:00 am";
			var input4 = "June 5, 2017 11:00:00 pm";

			var result1 = function.Execute(FunctionsHelper.CreateArgs(input1), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(input2), this.ParsingContext);
			var result3 = function.Execute(FunctionsHelper.CreateArgs(input3), this.ParsingContext);
			var result4 = function.Execute(FunctionsHelper.CreateArgs(input4), this.ParsingContext);

			Assert.AreEqual(-0.523862501, System.Math.Round(result1.ResultNumeric, 9));
			Assert.AreEqual(-0.978822933, System.Math.Round(result2.ResultNumeric, 9));
			Assert.AreEqual(-0.523862501, System.Math.Round(result3.ResultNumeric, 9));
			Assert.AreEqual(-0.978822933, System.Math.Round(result4.ResultNumeric, 9));
		}

		[TestMethod]
		public void CosHandlesInputDatesAreSeperatedByDashes()
		{
			var function = new Cos();

			var input1 = "1-17-2017 2:00";
			var input4 = "1-17-2017 2:00 am";

			var result1 = function.Execute(FunctionsHelper.CreateArgs(input1), this.ParsingContext);
			var result4 = function.Execute(FunctionsHelper.CreateArgs(input4), this.ParsingContext);

			Assert.AreEqual(0.276637268, System.Math.Round(result1.ResultNumeric, 9));
			Assert.AreEqual(0.276637268, System.Math.Round(result4.ResultNumeric, 9));
		}

		[TestMethod]
		public void CosFunctionWithErrorValuesAsInputReturnsTheInputErrorValue()
		{
			var func = new Cos();
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
