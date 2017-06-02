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
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;


namespace EPPlusTest.FormulaParsing.Excel.Functions.DateTimeFunctions
{
	[TestClass]
	public class TimeValueTests : DateTimeFunctionsTestBase
	{
		#region TimeValue Function (Execute) Tests
		[TestMethod]
		public void TimeValueWithInvalidArgumentReturnsPoundValue()
		{
			var func = new TimeValue();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void TimeValueIsGivenAString()
		{
			var function = new TimeValue();
			
			var input1 = "Noon";
			var input2 = "midnight";
			var input3 = "one o'clock";

			var result1 = function.Execute(FunctionsHelper.CreateArgs(input1), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(input2), this.ParsingContext);
			var result3 = function.Execute(FunctionsHelper.CreateArgs(input3), this.ParsingContext);

			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result1.Result).Type);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result2.Result).Type);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result3.Result).Type);
		}

		[TestMethod]
		public void TimeValueIsGivenMilitaryTime()
		{
			var function = new TimeValue();

			var input1 = "00:00";
			var input2 = "00:01";
			var input3 = "24:00";

			var result1 = function.Execute(FunctionsHelper.CreateArgs(input1), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(input2), this.ParsingContext);
			var result3 = function.Execute(FunctionsHelper.CreateArgs(input3), this.ParsingContext);
			
			Assert.AreEqual(0.0,result1.ResultNumeric);
			Assert.AreEqual( 0.00069444, System.Math.Round(result2.ResultNumeric, 8));
			Assert.AreEqual(0.0, result3.ResultNumeric);
		}

		[TestMethod]
		public void TimeValueMillitaryTimePast2400ActsLikeItWasModded()
		{
			var function = new TimeValue();

			var input1 = "25:00";
			var input2 = "01:00";
			var input3 = "26:00";
			var input4 = "02:00";

			var result1 = function.Execute(FunctionsHelper.CreateArgs(input1), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(input2), this.ParsingContext);
			var result3 = function.Execute(FunctionsHelper.CreateArgs(input3), this.ParsingContext);
			var result4 = function.Execute(FunctionsHelper.CreateArgs(input4), this.ParsingContext);

			Assert.AreEqual(0.041666667, System.Math.Round(result1.ResultNumeric,9));
			Assert.AreEqual(0.041666667, System.Math.Round(result2.ResultNumeric,9));
			Assert.AreEqual(0.083333333, System.Math.Round(result3.ResultNumeric, 9));
			Assert.AreEqual(0.083333333, System.Math.Round(result4.ResultNumeric, 9));


		}

		[TestMethod]
		public void TimeValueDatesAreIgnored()
		{
			var function = new TimeValue();

			var input1 = "1/11/2011 2:00";
			var input2 = "1/11/2011 2:00 am";
			var input3 = "17/1/2011 2:00 am";
			var input4 = "11/Jan/2011 2:00 am";
			var input5 = "Jan 11, 2011 2:00 am";

			var result1 = function.Execute(FunctionsHelper.CreateArgs(input1), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(input2), this.ParsingContext);
			var result3 = function.Execute(FunctionsHelper.CreateArgs(input3), this.ParsingContext);
			var result4 = function.Execute(FunctionsHelper.CreateArgs(input4), this.ParsingContext);
			var result5 = function.Execute(FunctionsHelper.CreateArgs(input5), this.ParsingContext);

			Assert.AreEqual(0.083333333, System.Math.Round(result1.ResultNumeric, 9));
			Assert.AreEqual(0.083333333, System.Math.Round(result2.ResultNumeric, 9));
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result3.Result).Type);
			Assert.AreEqual(0.083333333, System.Math.Round(result4.ResultNumeric, 9));
			Assert.AreEqual(0.083333333, System.Math.Round(result5.ResultNumeric, 9));
		}

		[TestMethod]
		public void TimeValueIsNormal12HourClock()
		{
			var function = new TimeValue();

			var input1 = "00:00 am";
			var input2 = "00:01 am";
			var input3 = "12:00 pm";
			var input4 = "12:00 am";
			var input6 = "1:00 pm";
			var input7 = "13:00";

			var result1 = function.Execute(FunctionsHelper.CreateArgs(input1), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(input2), this.ParsingContext);
			var result3 = function.Execute(FunctionsHelper.CreateArgs(input3), this.ParsingContext);
			var result4 = function.Execute(FunctionsHelper.CreateArgs(input4), this.ParsingContext);
			var result6 = function.Execute(FunctionsHelper.CreateArgs(input6), this.ParsingContext);
			var result7 = function.Execute(FunctionsHelper.CreateArgs(input7), this.ParsingContext);

			Assert.AreEqual(0.0, result1.ResultNumeric);
			Assert.AreEqual(0.000694444, System.Math.Round(result2.ResultNumeric, 9));
			Assert.AreEqual(0.5, result3.ResultNumeric);
			Assert.AreEqual(0.0, result4.ResultNumeric);
			Assert.AreEqual(0.541666667, System.Math.Round(result6.ResultNumeric, 9));
			Assert.AreEqual(0.541666667, System.Math.Round(result7.ResultNumeric, 9));


		}

		[TestMethod]
		public void TimeValueEPPisModdingTimesOver12Pm()
		{
			var function = new TimeValue();

			var input1 = "13:00 pm";
			var input2 = "1:00 pm";
			var input3 = "16:00 pm";
			var input4 = "4:00 pm";

			var result1 = function.Execute(FunctionsHelper.CreateArgs(input1), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(input2), this.ParsingContext);
			var result3 = function.Execute(FunctionsHelper.CreateArgs(input3), this.ParsingContext);
			var result4 = function.Execute(FunctionsHelper.CreateArgs(input4), this.ParsingContext);

			Assert.AreEqual(0.541666667, System.Math.Round(result1.ResultNumeric, 9));
			Assert.AreEqual(0.541666667, System.Math.Round(result2.ResultNumeric, 9));
			Assert.AreEqual(0.666666667, System.Math.Round(result3.ResultNumeric, 9));
			Assert.AreEqual(0.666666667, System.Math.Round(result4.ResultNumeric, 9));
		}

		[TestMethod]
		public void TimeValueTestMilitaryTimeAndNormalTimeComparisions()
		{
			var function = new TimeValue();

			var input1 = "16:30";
			var input2 = "4:30 pm";
			var input3 = "02:30";
			var input4 = "2:30 am";

			var result1 = function.Execute(FunctionsHelper.CreateArgs(input1), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(input2), this.ParsingContext);
			var result3 = function.Execute(FunctionsHelper.CreateArgs(input3), this.ParsingContext);
			var result4 = function.Execute(FunctionsHelper.CreateArgs(input4), this.ParsingContext);

			Assert.AreEqual(0.6875, System.Math.Round(result1.ResultNumeric, 9));
			Assert.AreEqual(0.6875, System.Math.Round(result2.ResultNumeric, 9));
			Assert.AreEqual(0.104166667, System.Math.Round(result3.ResultNumeric, 9));
			Assert.AreEqual(0.104166667, System.Math.Round(result4.ResultNumeric, 9));

		}





		#endregion
	}
}
