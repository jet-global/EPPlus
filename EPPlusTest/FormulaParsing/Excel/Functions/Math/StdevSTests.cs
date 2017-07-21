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
	public class StdevSTests : MathFunctionsTestBase
	{
		#region StdevS Function(Execute) Tests

		[TestMethod]
		public void StdevSIsGivenBooleanInputs()
		{
			var function = new StdevS();
			var boolInputTrue = true;
			var boolInputFalse = false;
			var result1 = function.Execute(FunctionsHelper.CreateArgs(boolInputTrue, boolInputTrue, boolInputFalse), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(boolInputTrue, boolInputTrue, boolInputTrue), this.ParsingContext);
			var result3 = function.Execute(FunctionsHelper.CreateArgs(boolInputTrue, boolInputFalse, boolInputFalse), this.ParsingContext);

			Assert.AreEqual(0.577350269, result1.ResultNumeric, .00001);
			Assert.AreEqual(0, result2.ResultNumeric, .00001);
			Assert.AreEqual(0.577350269, result3.ResultNumeric, .00001);
		}

		[TestMethod]
		public void StdevSIsGivenBooleanInputsFromCellRefrence()
		{
			var function = new StdevS();

			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = true;
				worksheet.Cells["B2"].Value = false;
				worksheet.Cells["B3"].Formula = "=stdev.s(B1,B1,B2)";
				worksheet.Cells["B4"].Formula = "=stdev.s(B1,B1,B1)";
				worksheet.Cells["B5"].Formula = "=stdev.s(B1,B2,B2)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B3"].Value).Type);
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B4"].Value).Type);
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B5"].Value).Type);
			}
		}

		[TestMethod]
		public void StdevSIsGivenADateSeperatedByABackslash()
		{
			var function = new StdevS();
			var input1 = "1/17/2011 2:00";
			var input2 = "6/17/2011 2:00";
			var input3 = "1/17/2012 2:00";
			var input4 = "1/17/2013 2:00";
			var result1 = function.Execute(FunctionsHelper.CreateArgs(input1, input1, input1), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(input1, input2, input1), this.ParsingContext);
			var result3 = function.Execute(FunctionsHelper.CreateArgs(input1, input3, input4), this.ParsingContext);
			Assert.AreEqual(0, result1.ResultNumeric, .00001);
			Assert.AreEqual(87.17989065, result2.ResultNumeric, .00001);
			Assert.AreEqual(365.500114, result3.ResultNumeric, .00001);
		}

		[TestMethod]
		public void StdevSIsGivenADateSeperatedByABackslashInputFromCellRefrence()
		{
			var function = new StdevS();

			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "1/17/2011 2:00";
				worksheet.Cells["B2"].Value = "6/17/2011 2:00";
				worksheet.Cells["B3"].Value = "1/17/2012 2:00";
				worksheet.Cells["B4"].Value = "1/17/2013 2:00";
				worksheet.Cells["B5"].Formula = "=stdev.s(B1,B1,B1)";
				worksheet.Cells["B6"].Formula = "=stdev.s(B1,B2,B1)";
				worksheet.Cells["B7"].Formula = "=stdev.s(B1,B3,B4)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B5"].Value).Type);
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B6"].Value).Type);
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B7"].Value).Type);
			}
		}

		[TestMethod]
		public void StdevSIsGivenMaxAmountOfInputs254()
		{
			var function = new StdevS();
			var input1 = 1;

			var result1 = function.Execute(FunctionsHelper.CreateArgs(
				100, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1,
				input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1,
				input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1,
				input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1,
				input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1,
				input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1,
				input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1,
				input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1,
				input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1,
				input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1,
				input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1,
				input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1,
				input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1,
				input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1,
				input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1,
				input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1,
				input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1, input1
				), this.ParsingContext);
			Assert.AreEqual(6.211812471, result1.ResultNumeric, .00001);
		}

		[TestMethod]
		public void StdevSIsGivenNumberInputFromCellRefrence()
		{
			var function = new StdevS();

			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 66;
				worksheet.Cells["B2"].Value = 52;
				worksheet.Cells["B3"].Value = 77;
				worksheet.Cells["B4"].Value = 71;
				worksheet.Cells["B5"].Value = 30;
				worksheet.Cells["B6"].Value = 90;
				worksheet.Cells["B7"].Value = 26;
				worksheet.Cells["B8"].Value = 56;
				worksheet.Cells["B9"].Value = 7;
				worksheet.Cells["A10"].Formula = "=stdev.s(B:B)";
				worksheet.Cells["A11"].Formula = "=stdev.s(B1,B3,B5,B6,B9)";
				worksheet.Cells["A12"].Formula = "=stdev.s(B1,B3,B5,B6)";
				worksheet.Calculate();
				Assert.AreEqual(26.97581221, (double)worksheet.Cells["A10"].Value, .00001);
				Assert.AreEqual(34.47462835, (double)worksheet.Cells["A11"].Value, .00001);
				Assert.AreEqual(25.77304794, (double)worksheet.Cells["A12"].Value, .00001);
			}
		}

		[TestMethod]
		public void StdevSIsGivenNumberInputFromCellRefrenceWithEmptyCellsFirst()
		{
			var function = new StdevS();

			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B3"].Value = 66;
				worksheet.Cells["B4"].Value = 52;
				worksheet.Cells["B5"].Value = 77;
				worksheet.Cells["B6"].Value = 71;
				worksheet.Cells["B7"].Value = 30;
				worksheet.Cells["B8"].Value = 90;
				worksheet.Cells["B9"].Value = 26;
				worksheet.Cells["B10"].Value = 56;
				worksheet.Cells["B12"].Value = 7;
				worksheet.Cells["A12"].Formula = "=stdev.s(B1:B12)";
				worksheet.Calculate();
				Assert.AreEqual(26.97581221, (double)worksheet.Cells["A12"].Value, .00001);
			}
		}

		[TestMethod]
		public void StdevSIsGivenNumberInputFromCellRefrenceWithEmptyCellsFirstAndAnInvalidCellInTheMiddle()
		{
			var function = new StdevS();

			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B3"].Value = 66;
				worksheet.Cells["B4"].Value = 52;
				worksheet.Cells["B5"].Value = 77;
				worksheet.Cells["B6"].Value = 71;
				//B7 is an empty cell
				worksheet.Cells["B8"].Value = 90;
				worksheet.Cells["B9"].Value = 26;
				worksheet.Cells["B10"].Value = 56;
				worksheet.Cells["B11"].Value = 7;
				worksheet.Cells["A12"].Formula = "=stdev.s(B1:B11)";
				worksheet.Calculate();
				Assert.AreEqual(27.35448514, (double)worksheet.Cells["A12"].Value, .00001);
			}
		}

		[TestMethod]
		public void StdevSIsGivenNumbersAsInputs()
		{
			var function = new StdevS();
			var result1 = function.Execute(FunctionsHelper.CreateArgs(-1, 0, 1), this.ParsingContext);
			Assert.AreEqual(1, result1.ResultNumeric, .00001);
		}

		[TestMethod]
		public void StdevSIsGivenAMixOfInputTypes()
		{
			var function = new StdevS();
			var result1 = function.Execute(FunctionsHelper.CreateArgs(1, true, null, "6/17/2011 2:00", "02:00 am"), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(1, true, "6/17/2011 2:00", "02:00 am"), this.ParsingContext);
			Assert.AreEqual(18206.31704, result1.ResultNumeric, .00001);
			Assert.AreEqual(20355.19445, result2.ResultNumeric, .00001);
		}


		[TestMethod]
		public void StdevSIsGivenAMixOfInputTypesByCellRange()
		{
			var function = new StdevS();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 15;
				worksheet.Cells["B2"].Value = 1;
				worksheet.Cells["B3"].Value = "6/17/2011 2:00";
				worksheet.Cells["B4"].Value = "02:00 am";
				worksheet.Cells["B9"].Formula = "=stdev.s(B1:B4)";
				worksheet.Calculate();
				Assert.AreEqual(9.899494937, (double)worksheet.Cells["B9"].Value, .00001);
			}
		}

		[TestMethod]
		public void StdevSIsGivenAMixOfInputTypesWithANullInTheCenterByCellRefrenceAndRange()
		{
			var function = new StdevS();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 15;
				worksheet.Cells["B2"].Value = 1;
				worksheet.Cells["B3"].Value = "6/17/2011 2:00";
				worksheet.Cells["B4"].Value = "02:00 am";
				worksheet.Cells["B9"].Formula = "=stdev.s(B1:B4)";
				worksheet.Calculate();
				Assert.AreEqual(9.899494937, (double)worksheet.Cells["B9"].Value, .00001);//This is returning 9.899494937. That is the same as std.s(15,1) or (15,true)
			}
		}

		[TestMethod]
		public void StdevSIsGivenAMixOfInputTypesWithANullInTheCenterByCellRefrenceAndTwoRangeInputs()
		{
			var function = new StdevS();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 15;
				worksheet.Cells["B2"].Value = 1;
				//empty B3 cell
				worksheet.Cells["B4"].Value = "6/17/2011 2:00";
				worksheet.Cells["B5"].Value = "02:00 am";
				worksheet.Cells["B8"].Formula = "=stdev.s(B1,B2)";
				worksheet.Cells["B9"].Formula = "=stdev.s(B1:B5,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(9.899494937, (double)worksheet.Cells["B8"].Value, .00001);
				Assert.AreEqual(8.082903769, (double)worksheet.Cells["B9"].Value, .00001);
			}
		}

		[TestMethod]
		public void StdevSIsGivenAMixOfTwoTypesAndTwoRangesThatShouldHaveAnOutput()
		{
			var function = new StdevS();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 23;
				worksheet.Cells["B2"].Value = 1;
				worksheet.Cells["B3"].Value = "6/17/2011 2:00";
				worksheet.Cells["B4"].Value = "02:00 am";
				worksheet.Cells["A1"].Value = 23;
				worksheet.Cells["A2"].Value = 15;
				worksheet.Cells["A3"].Value = "02:00 am";
				worksheet.Cells["B9"].Formula = "=stdev.s(B1:B4, A1:A3)";
				worksheet.Calculate();
				Assert.AreEqual(10.37625494, (double)worksheet.Cells["B9"].Value, .0001);
			}
		}

		[TestMethod]
		public void StdevSIsTheSameTestsAsGivenAMixOfInputTypesByCellRefrenceExceptTheyAreAllOnes()
		{
			var function = new StdevS();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 1;
				//empty B3 cell
				worksheet.Cells["B4"].Value = 1;
				worksheet.Cells["B5"].Value = 1;
				worksheet.Cells["B6"].Formula = "=stdev.s(B1,B2,B4,B5)";
				worksheet.Cells["B7"].Formula = "=stdev.s(B1,B2,B3,B4,B5)";
				worksheet.Cells["B8"].Formula = "=stdev.s(B1,B2)";
				worksheet.Cells["B9"].Formula = "=stdev.s(B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B6"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B7"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B8"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B9"].Value);
			}
		}

		[TestMethod]
		public void StdevSTestingDirectInputVsRangeInputWithAGapInTheMiddle()
		{
			var function = new StdevS();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B3"].Value = 1;
				worksheet.Cells["B8"].Formula = "=stdev.s(B1,B3)";
				worksheet.Cells["B9"].Formula = "=stdev.s(B1:B3)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B8"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B9"].Value);
			}
		}

		[TestMethod]
		public void StdevSTestingDirectINputVsRangeInputTest2()
		{
			var function = new StdevS();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 0;
				worksheet.Cells["B8"].Formula = "=stdev.s(B1,B2)";
				worksheet.Cells["B9"].Formula = "=stdev.s(B1:B2)";
				worksheet.Calculate();
				Assert.AreEqual(0.707106781, (double)worksheet.Cells["B8"].Value, .00001);
				Assert.AreEqual(0.707106781, (double)worksheet.Cells["B9"].Value, .00001);
			}
		}

		[TestMethod]
		public void StdevSIsGivenAStringInputWithAEmptyCellInTheMiddle()
		{
			var function = new StdevS();
			var result1 = function.Execute(FunctionsHelper.CreateArgs(66, 52, 77, 71, 30, 90, 26, 56, 7), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(66, 52, 77, 71,null, 30, 90, 26, 56, 7), this.ParsingContext);
			Assert.AreEqual(26.97581221, result1.ResultNumeric, .00001);
			Assert.AreEqual(30.42020527, result2.ResultNumeric, .00001);
		}

		[TestMethod]
		public void StdevSIsGivenAMixedStringInputWithAEmptyCellInTheMiddle()
		{
			var function = new StdevS();
			var result1 = function.Execute(FunctionsHelper.CreateArgs(10, 2, "6/17/2011 2:00", "02:00 am"), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(10, 2, "6/17/2011 2:00",null, "02:00 am"), this.ParsingContext);
			Assert.AreEqual(20353.52823, result1.ResultNumeric, .00001);
			Assert.AreEqual(18205.19946, result2.ResultNumeric, .00001);
		}

		[TestMethod]
		public void StdevSIsGivenAMixOfTwoTypesAndTwoRangesThatShouldHaveAnOutputTestTwo()
		{
			var function = new StdevS();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 10;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = "6/17/2011 2:00";
				worksheet.Cells["B4"].Value = "02:00 am";
				worksheet.Cells["B9"].Formula = "=stdev.s(B1:B4)";
				worksheet.Calculate();
				Assert.AreEqual(5.656854249, (double) worksheet.Cells["B9"].Value, .0001);
			}
		}

		[TestMethod]
		public void StdevSIsGivenTwoRangesOfIntsAsInputs()
		{
			var function = new StdevS();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 34;
				worksheet.Cells["B4"].Value = 56;
				worksheet.Cells["B5"].Value = 32;
				worksheet.Cells["B6"].Value = 76;
				worksheet.Cells["B7"].Value = 2;
				worksheet.Cells["B8"].Value = 3;
				worksheet.Cells["B9"].Value = 5;
				worksheet.Cells["B10"].Value = 7;
				worksheet.Cells["B11"].Value = 45;
				worksheet.Cells["A9"].Formula = "=stdev.s(B1:B6, B7:B11)";
				worksheet.Calculate();
				Assert.AreEqual(26.30762074, (double)worksheet.Cells["A9"].Value, .00001);
			}
		}

		[TestMethod]
		public void StdevSIsGivenTwoRangesOfIntsAsInputsWithATimeInTheMiddleOfTheFirstRange()
		{
			var function = new StdevS();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 34;
				worksheet.Cells["B4"].Value = "12:00";
				worksheet.Cells["B5"].Value = 32;
				worksheet.Cells["B6"].Value = 76;
				worksheet.Cells["B7"].Value = 2;
				worksheet.Cells["B8"].Value = 3;
				worksheet.Cells["B9"].Value = 5;
				worksheet.Cells["B10"].Value = 7;
				worksheet.Cells["B11"].Value = 45;
				worksheet.Cells["A9"].Formula = "=stdev.s(B1:B6, B7:B11)";
				worksheet.Calculate();
				Assert.AreEqual(25.35985454, (double)worksheet.Cells["A9"].Value, .00001);
			}
		}

		[TestMethod]
		public void StdevSIsGivenTwoRangesOfIntsAsInputsWithATimeInTheMiddleOfTheFirstRangeTestTwo()
		{
			var function = new StdevS();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 34;
				worksheet.Cells["B4"].Value = "12:00";
				worksheet.Cells["B5"].Value = "6/17/2011 2:00";
				worksheet.Cells["B6"].Value = 76;
				worksheet.Cells["B7"].Value = 2;
				worksheet.Cells["B8"].Value = 3;
				worksheet.Cells["B9"].Value = 5;
				worksheet.Cells["B10"].Value = 7;
				worksheet.Cells["B11"].Value = 45;
				worksheet.Cells["A9"].Formula = "=stdev.s(B1:B6, B7:B11)";
				worksheet.Calculate();
				Assert.AreEqual(26.56647846, (double)worksheet.Cells["A9"].Value, .00001);
			}
		}

		[TestMethod]
		public void StdevSIsGivenAStringInputWithAEmptyCellInTheMiddleByCellRefrence()
		{
			var function = new StdevS();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B3"].Value = 66;
				worksheet.Cells["B4"].Value = 52;
				worksheet.Cells["B5"].Value = 77;
				worksheet.Cells["B6"].Value = 71;
				//empty B7 cell
				worksheet.Cells["B8"].Value = 90;
				worksheet.Cells["B9"].Value = 26;
				worksheet.Cells["B10"].Value = 56;
				worksheet.Cells["B11"].Value = 7;
				worksheet.Cells["B12"].Value = 30;
				worksheet.Cells["A12"].Formula = "=stdev.s(B3:B12)";
				worksheet.Calculate();
				Assert.AreEqual(26.97581221, (double)worksheet.Cells["A12"].Value, .00001);
			}
		}

		[TestMethod]
		public void StdevSIsGivenMilitaryTimesAsInputs()
		{
			var function = new StdevS();
			var result1 = function.Execute(FunctionsHelper.CreateArgs("00:00", "02:00", "13:00"), this.ParsingContext);
			Assert.AreEqual(0.291666667, result1.ResultNumeric, .00001);
		}

		[TestMethod]
		public void StdevSIsGivenNumbersAsInputstakeTwo()
		{
			var function = new StdevS();
			var result1 = function.Execute(FunctionsHelper.CreateArgs(-1, -1, -1), this.ParsingContext);
			Assert.AreEqual(0, result1.ResultNumeric, .00001);
		}

		[TestMethod]
		public void StdevSIsGivenMilitaryTimesAsInputsByCellRefrence()
		{
			var function = new StdevS();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "00:00";
				worksheet.Cells["B2"].Value = "02:00";
				worksheet.Cells["B3"].Value = "13:00";
				worksheet.Cells["B4"].Formula = "=stdev.s(B1,B2,B3)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B4"].Value).Type);
			}
		}

		[TestMethod]
		public void StdevSIsGivenMilitaryTimesAsInputsByCellRange()
		{
			var function = new StdevS();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "00:00";
				worksheet.Cells["B2"].Value = "02:00";
				worksheet.Cells["B3"].Value = "13:00";
				worksheet.Cells["B4"].Formula = "=stdev.s(B1:B3)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B4"].Value).Type);
			}
		}

		[TestMethod]
		public void StdevSIsGiven12HourTimesAsInputs()
		{
			var function = new StdevS();
			var result1 = function.Execute(FunctionsHelper.CreateArgs("12:00 am", "02:00 am", "01:00 pm"), this.ParsingContext);
			Assert.AreEqual(0.291666667, result1.ResultNumeric, .00001);
		}

		[TestMethod]
		public void StdevSIsGiven12HourTimesAsInputsByCellRefrence()
		{
			var function = new StdevS();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "12:00 am";
				worksheet.Cells["B2"].Value = "02:00 am";
				worksheet.Cells["B3"].Value = "01:00 pm";
				worksheet.Cells["B4"].Formula = "=stdev.s(B1,B2,B3)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B4"].Value).Type);
			}
		}

		[TestMethod]
		public void StdevSIsGiven12HourTimesAsInputsByCellRefrenceTestTwo()
		{
			var function = new StdevS();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "12:00 am";
				worksheet.Cells["B2"].Value = "02:00 am";
				worksheet.Cells["B3"].Value = "01:00 pm";
				worksheet.Cells["B4"].Formula = "=stdev.s(B1,B2,B3)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B4"].Value).Type);
			}
		}

		[TestMethod]
		public void StdevSIsGiven12HourTimesAsInputsByCellRange()
		{
			var function = new StdevS();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "12:00 am";
				worksheet.Cells["B2"].Value = "02:00 am";
				worksheet.Cells["B3"].Value = "01:00 pm";
				worksheet.Cells["B4"].Formula = "=stdev.s(B1:B3)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B4"].Value).Type);
			}
		}

		[TestMethod]
		public void StdevSIsGivenMonthDayYear12HourTimeAsInputs()
		{
			var function = new StdevS();
			var result1 = function.Execute(FunctionsHelper.CreateArgs("Jan 17, 2011 2:00 am", "June 5, 2017 11:00 pm", "June 15, 2017 11:00 pm"), this.ParsingContext);
			Assert.AreEqual(1349.204675, result1.ResultNumeric, .00001);
		}

		[TestMethod]
		public void StdevSIsGivenMonthDayYear12HourTimeAsInputsByCellRefrence()
		{
			var function = new StdevS();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "Jan 17, 2011 2:00 am";
				worksheet.Cells["B2"].Value = "June 5, 2017 11:00 pm";
				worksheet.Cells["B3"].Value = "June 15, 2017 11:00 pm";
				worksheet.Cells["B4"].Formula = "=stdev.s(B1,B2,B3)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B4"].Value).Type);
			}
		}

		[TestMethod]
		public void StdevSIsGivenMonthDayYear12HourTimeAsInputsByCellRange()
		{
			var function = new StdevS();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "Jan 17, 2011 2:00 am";
				worksheet.Cells["B2"].Value = "June 5, 2017 11:00 pm";
				worksheet.Cells["B3"].Value = "June 15, 2017 11:00 pm";
				worksheet.Cells["B4"].Formula = "=stdev.s(B1:B3)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B4"].Value).Type);
			}
		}

		[TestMethod]
		public void StdevSIsGivenDateTimeInputsSeperatedByADashAsInputs()
		{
			var function = new StdevS();
			var result1 = function.Execute(FunctionsHelper.CreateArgs("1-17-2017 2:00", "6-17-2017 2:00", "9-17-2017 2:00"), this.ParsingContext);
			Assert.AreEqual(122.6879511, result1.ResultNumeric, .00001);
		}

		[TestMethod]
		public void StdevSIsGivenDateTimeInputsSeperatedByADashAsInputsByCellRefrence()
		{
			var function = new StdevS();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "1-17-2017 2:00";
				worksheet.Cells["B2"].Value = "6-17-2017 2:00";
				worksheet.Cells["B3"].Value = "9-17-2017 2:00";
				worksheet.Cells["B4"].Formula = "=stdev.s(B1,B2,B3)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B4"].Value).Type);
			}
		}

		[TestMethod]
		public void StdevSIsGivenDateTimeInputsSeperatedByADashAsInputsByCellRange()
		{
			var function = new StdevS();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "1-17-2017 2:00";
				worksheet.Cells["B2"].Value = "6-17-2017 2:00";
				worksheet.Cells["B3"].Value = "9-17-2017 2:00";
				worksheet.Cells["B4"].Formula = "=stdev.s(B1:B3)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B4"].Value).Type);
			}
		}

		[TestMethod]
		public void StdevSIsGivenStringsAsInputs()
		{
			var function = new StdevS();
			var result1 = function.Execute(FunctionsHelper.CreateArgs("string", "another string", "a third string"), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result1.Result).Type);
		}

		[TestMethod]
		public void StdevSIsGivenStringsAsInputsByCellRefrence()
		{
			var function = new StdevS();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "string";
				worksheet.Cells["B2"].Value = "another string";
				worksheet.Cells["B3"].Value = "a third string";
				worksheet.Cells["B4"].Formula = "=stdev.s(B1,B2,B3)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B4"].Value).Type);
			}
		}

		[TestMethod]
		public void StdevSIsGivenStringsAsInputsByCellRange()
		{
			var function = new StdevS();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "string";
				worksheet.Cells["B2"].Value = "another string";
				worksheet.Cells["B3"].Value = "a third string";
				worksheet.Cells["B4"].Formula = "=stdev.s(B1:B3)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B4"].Value).Type);
			}
		}

		[TestMethod]
		public void StdevSIsGivenStringNumbersAsInputs()
		{
			var function = new StdevS();
			var result1 = function.Execute(FunctionsHelper.CreateArgs("5", "6", "7"), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs("5.5", "6.6", "7.7"), this.ParsingContext);
			Assert.AreEqual(1, result1.ResultNumeric, .00001);
			Assert.AreEqual(1.1, result2.ResultNumeric, .00001);
		}

		[TestMethod]
		public void StdevSIsGivenStringNumbersAsInputsByCellRefrence()
		{
			var function = new StdevS();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["A1"].Value = "5";
				worksheet.Cells["A2"].Value = "6";
				worksheet.Cells["A3"].Value = "7";
				worksheet.Cells["B1"].Value = "5.5";
				worksheet.Cells["B2"].Value = "6.6";
				worksheet.Cells["B3"].Value = "7.7";
				worksheet.Cells["A4"].Formula = "=stdev.s(A1,A2,A3)";
				worksheet.Cells["B4"].Formula = "=stdev.s(B1,B2,B3)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["A4"].Value).Type);
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B4"].Value).Type);
			}
		}

		[TestMethod]
		public void StdevSIsGivenStringNumbersAsInputsByCellRange()
		{
			var function = new StdevS();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["A1"].Value = "5";
				worksheet.Cells["A2"].Value = "6";
				worksheet.Cells["A3"].Value = "7";
				worksheet.Cells["B1"].Value = "5.5";
				worksheet.Cells["B2"].Value = "6.6";
				worksheet.Cells["B3"].Value = "7.7";
				worksheet.Cells["A4"].Formula = "=stdev.s(A1:A3)";
				worksheet.Cells["B4"].Formula = "=stdev.s(B1:B3)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["A4"].Value).Type);
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["B4"].Value).Type);
			}
		}

		[TestMethod]
		public void StdevSIsGivenASingleTrueBooleanInput()
		{
			var function = new StdevS();
			var result1 = function.Execute(FunctionsHelper.CreateArgs(true), this.ParsingContext);

			Assert.AreEqual(0, result1.ResultNumeric, .00001); ;
		}

		[TestMethod]
		public void StdevSIsGivenASingleFalseBooleanInput()
		{
			var function = new StdevS();
			var result1 = function.Execute(FunctionsHelper.CreateArgs(false), this.ParsingContext);

			Assert.AreEqual(0, result1.ResultNumeric, .00001); ;
		}

		[TestMethod]
		public void StdevSIsGivenASingleStringInput()
		{
			var function = new StdevS();
			var result1 = function.Execute(FunctionsHelper.CreateArgs("string"), this.ParsingContext);

			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result1.Result).Type);
		}
		#endregion
	}
}
