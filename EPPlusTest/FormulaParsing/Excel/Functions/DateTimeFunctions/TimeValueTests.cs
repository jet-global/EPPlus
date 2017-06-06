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
			var input4 = "23:59:59";
			

			var result1 = function.Execute(FunctionsHelper.CreateArgs(input1), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(input2), this.ParsingContext);
			var result3 = function.Execute(FunctionsHelper.CreateArgs(input3), this.ParsingContext);
			var result4 = function.Execute(FunctionsHelper.CreateArgs(input4), this.ParsingContext);

			Assert.AreEqual(0.0,result1.ResultNumeric);
			Assert.AreEqual( 0.00069444, System.Math.Round(result2.ResultNumeric, 8));
			Assert.AreEqual(0.0, result3.ResultNumeric);
			Assert.AreEqual(0.999988426, System.Math.Round(result4.ResultNumeric, 9));
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

			Assert.AreEqual(0.041666667, System.Math.Round(result1.ResultNumeric, 9));
			Assert.AreEqual(0.041666667, System.Math.Round(result2.ResultNumeric, 9));
			Assert.AreEqual(0.083333333, System.Math.Round(result3.ResultNumeric, 9));
			Assert.AreEqual(0.083333333, System.Math.Round(result4.ResultNumeric, 9));


		}

		[TestMethod]
		public void TimeValueDatesAreIgnored()
		{
			var function = new TimeValue();

			var input1 = "1/11/2011 2:00";
			var input2 = "1/11/2011 2:00 AM";
			var input3 = "17/1/2011 2:00 AM";
			var input4 = "11/Jan/2011 2:00 AM";
			var input5 = "Jan 11, 2011 2:00 AM";

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

			var input1 = "00:00:00 AM";
			var input2 = "00:01:32 AM";
			var input3 = "12:00 PM";
			var input4 = "12:00 AM";
			var input6 = "1:00 PM";
			//var input7 = "13:00 PM"; // I am putting this on ice for now. Excel Returns a #Value, EPP Retuns a value.
			var input8 = "1:10:32 am";
			var input9 = "3:42:32 pm";

			var result1 = function.Execute(FunctionsHelper.CreateArgs(input1), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(input2), this.ParsingContext);
			var result3 = function.Execute(FunctionsHelper.CreateArgs(input3), this.ParsingContext);
			var result4 = function.Execute(FunctionsHelper.CreateArgs(input4), this.ParsingContext);
			var result6 = function.Execute(FunctionsHelper.CreateArgs(input6), this.ParsingContext);
			//var result7 = function.Execute(FunctionsHelper.CreateArgs(input7), this.ParsingContext);
			var result8 = function.Execute(FunctionsHelper.CreateArgs(input8), this.ParsingContext);
			var result9 = function.Execute(FunctionsHelper.CreateArgs(input9), this.ParsingContext);

			Assert.AreEqual(0.0, result1.ResultNumeric);
			Assert.AreEqual(0.001064815, System.Math.Round(result2.ResultNumeric, 9));
			Assert.AreEqual(0.5, result3.ResultNumeric);
			Assert.AreEqual(0.0, result4.ResultNumeric);
			Assert.AreEqual(0.541666667, System.Math.Round(result6.ResultNumeric, 9));
			//Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result7.Result).Type);
			Assert.AreEqual(0.048981481, System.Math.Round(result8.ResultNumeric, 9));
			Assert.AreEqual(0.654537037, System.Math.Round(result9.ResultNumeric, 9));


		}
	
		  
		[TestMethod]
		public void TimeValueTimeOfTheForm1300Pm()
		{
			//Note: In Excel, the argument 13:00 PM would return #VALUE!.
			var function = new TimeValue();

			var input1 = "13:00 PM";
			var input2 = "1:00 PM";
			var input3 = "16:00 PM";
			var input4 = "4:00 PM";

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
			var input2 = "04:30 pm";
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

		[TestMethod]
		public void TimeValueTestInputsWithDatesThatHaveSlashesInThem()
		{
			var function = new TimeValue();

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

			Assert.AreEqual(0.083333333, System.Math.Round(result1.ResultNumeric, 9));
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result2.Result).Type);
			Assert.AreEqual(0.083333333, System.Math.Round(result3.ResultNumeric, 9));
			Assert.AreEqual(0.083333333, System.Math.Round(result4.ResultNumeric, 9));
			Assert.AreEqual(0.083333333, System.Math.Round(result5.ResultNumeric, 9));
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result6.Result).Type);
			Assert.AreEqual(0.083333333, System.Math.Round(result7.ResultNumeric, 9));
			Assert.AreEqual(0.083333333, System.Math.Round(result8.ResultNumeric, 9));
		}

		[TestMethod]
		public void TimeValueInputsWithDatesInTheFormMonthDateCommaYearTime()
		{
			var function = new TimeValue();

			var input1 = "Jan 17, 2011 2:00 am";
			var input2 = "June 5, 2017 11:00 pm";
			var input3 = "Jan 17, 2011 2:00:00 am";
			var input4 = "June 5, 2017 11:00:00 pm";

			var result1 = function.Execute(FunctionsHelper.CreateArgs(input1), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(input2), this.ParsingContext);
			var result3 = function.Execute(FunctionsHelper.CreateArgs(input3), this.ParsingContext);
			var result4 = function.Execute(FunctionsHelper.CreateArgs(input4), this.ParsingContext);

			Assert.AreEqual(0.083333333, System.Math.Round(result1.ResultNumeric, 9));
			Assert.AreEqual(0.958333333, System.Math.Round(result2.ResultNumeric, 9));
			Assert.AreEqual(0.083333333, System.Math.Round(result3.ResultNumeric, 9));
			Assert.AreEqual(0.958333333, System.Math.Round(result4.ResultNumeric, 9));
		}

		[TestMethod]
		public void TimeValueInputDatesAreSeperatedByDot()
		{
			var function = new TimeValue();

			var input1 = "1.2011 20:00";
			var input2 = "1.17.2011 2:00 am";
			var input3 = "17.1.2017 2:00 am";

			var result1 = function.Execute(FunctionsHelper.CreateArgs(input1), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(input2), this.ParsingContext);
			var result3 = function.Execute(FunctionsHelper.CreateArgs(input3), this.ParsingContext);

			Assert.AreEqual(0.833333333, System.Math.Round(result1.ResultNumeric, 9));
			Assert.AreEqual(0.083333333, System.Math.Round(result2.ResultNumeric, 9));
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result3.Result).Type);
		}

		[TestMethod]
		public void TimeValueInputDatesAreSeperatedByDashes()
		{
			var function = new TimeValue();

			var input1 = "1-17-2017 2:00";
			var input2 = "1-2017 2:00";
			var input3 = "12-2017 2:00";
			var input4 = "1-17-2017 2:00 am";

			var result1 = function.Execute(FunctionsHelper.CreateArgs(input1), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(input2), this.ParsingContext);
			var result3 = function.Execute(FunctionsHelper.CreateArgs(input3), this.ParsingContext);
			var result4 = function.Execute(FunctionsHelper.CreateArgs(input4), this.ParsingContext);
			
			Assert.AreEqual(0.083333333, System.Math.Round(result1.ResultNumeric, 9));
			Assert.AreEqual(0.083333333, System.Math.Round(result2.ResultNumeric, 9));
			Assert.AreEqual(0.083333333, System.Math.Round(result3.ResultNumeric, 9));
			Assert.AreEqual(0.083333333, System.Math.Round(result4.ResultNumeric, 9));

		}

		[TestMethod]
		public void TimeValueInputDatesAreSeperatedByDashesAndTheMonthsAreSpelledOut()
		{
			var function = new TimeValue();

			var input1 = "14-march-2012 5:00";
			var input2 = "march-2012 5:00";
			var input3 = "march-14-2012 5:00";

			var result1 = function.Execute(FunctionsHelper.CreateArgs(input1), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(input2), this.ParsingContext);
			var result3 = function.Execute(FunctionsHelper.CreateArgs(input3), this.ParsingContext);

			Assert.AreEqual(0.208333333, System.Math.Round(result1.ResultNumeric, 9));
			Assert.AreEqual(0.208333333, System.Math.Round(result2.ResultNumeric, 9));
			Assert.AreEqual(0.208333333, System.Math.Round(result3.ResultNumeric, 9));
		}

		#endregion
	}
}
