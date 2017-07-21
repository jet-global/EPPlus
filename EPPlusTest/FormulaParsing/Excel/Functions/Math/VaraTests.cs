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
	public class VaraTests : MathFunctionsTestBase
	{
		#region Vara Function(Execute) Tests

		[TestMethod]
		public void VaraIsGivenBooleanInputs()
		{
			var function = new Vara();
			var boolInputTrue = true;
			var boolInputFalse = false;
			var result1 = function.Execute(FunctionsHelper.CreateArgs(boolInputTrue, boolInputTrue, boolInputFalse), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(boolInputTrue, boolInputTrue, boolInputTrue), this.ParsingContext);
			var result3 = function.Execute(FunctionsHelper.CreateArgs(boolInputTrue, boolInputFalse, boolInputFalse), this.ParsingContext);

			Assert.AreEqual(0.333333333, result1.ResultNumeric, .00001);
			Assert.AreEqual(0, result2.ResultNumeric, .00001);
			Assert.AreEqual(0.333333333, result3.ResultNumeric, .00001);
		}

		[TestMethod]
		public void VaraIsGivenBooleanInputsFromCellRefrence()
		{
			var function = new Vara();

			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = true;
				worksheet.Cells["B2"].Value = false;
				worksheet.Cells["B3"].Formula = "=Vara(B1,B1,B2)";
				worksheet.Cells["B4"].Formula = "=Vara(B1,B1,B1)";
				worksheet.Cells["B5"].Formula = "=Vara(B1,B2,B2)";
				worksheet.Calculate();
				Assert.AreEqual(0.333333333, (double)worksheet.Cells["B3"].Value, .00001);
				Assert.AreEqual(0d, (double)worksheet.Cells["B4"].Value, .00001);
				Assert.AreEqual(0.333333333, (double)worksheet.Cells["B5"].Value, .00001);
			}
		}

		[TestMethod]
		public void VaraIsGivenADateSeperatedByABackslash()
		{
			var function = new Vara();
			var input1 = "1/17/2011 2:00";
			var input2 = "6/17/2011 2:00";
			var input3 = "1/17/2012 2:00";
			var input4 = "1/17/2013 2:00";
			var result1 = function.Execute(FunctionsHelper.CreateArgs(input1, input1, input1), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(input1, input2, input1), this.ParsingContext);
			var result3 = function.Execute(FunctionsHelper.CreateArgs(input1, input3, input4), this.ParsingContext);

			Assert.AreEqual(0, result1.ResultNumeric, .00001);
			Assert.AreEqual(7600.333333, result2.ResultNumeric, .00001);
			Assert.AreEqual(133590.3333, result3.ResultNumeric, .0001);
		}

		[TestMethod]
		public void VaraIsGivenADateSeperatedByABackslashInputFromCellRefrence()
		{
			var function = new Vara();

			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "1/17/2011 2:00";
				worksheet.Cells["B2"].Value = "6/17/2011 2:00";
				worksheet.Cells["B3"].Value = "1/17/2012 2:00";
				worksheet.Cells["B4"].Value = "1/17/2013 2:00";
				worksheet.Cells["B5"].Formula = "=Vara(B1,B1,B1)";
				worksheet.Cells["B6"].Formula = "=Vara(B1,B2,B1)";
				worksheet.Cells["B7"].Formula = "=Vara(B1,B3,B4)";
				worksheet.Calculate();
				Assert.AreEqual(0d, (double)worksheet.Cells["B5"].Value, .00001);
				Assert.AreEqual(0d, (double)worksheet.Cells["B6"].Value, .00001);
				Assert.AreEqual(0d, (double)worksheet.Cells["B7"].Value, .00001);
			}
		}

		[TestMethod]
		public void VaraIsGivenMaxAmountOfInputs254()
		{
			var function = new Vara();
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


			Assert.AreEqual(38.58661417, result1.ResultNumeric, .00001);
		}

		[TestMethod]
		public void VaraIsGivenNumberInputFromCellRefrence()
		{
			var function = new Vara();

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
				worksheet.Cells["A10"].Formula = "=Vara(B:B)";
				worksheet.Cells["A11"].Formula = "=Vara(B1,B3,B5,B6,B9)";
				worksheet.Cells["A12"].Formula = "=Vara(B1,B3,B5,B6)";
				worksheet.Calculate();

				Assert.AreEqual(727.6944444, (double)worksheet.Cells["A10"].Value, .00001);
				Assert.AreEqual(1188.5, (double)worksheet.Cells["A11"].Value, .00001);
				Assert.AreEqual(664.25, (double)worksheet.Cells["A12"].Value, .00001);
			}
		}

		[TestMethod]
		public void VaraIsGivenNumberInputFromCellRefrenceWithEmptyCellsFirst()
		{
			var function = new Vara();

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
				worksheet.Cells["A12"].Formula = "=Vara(B1:B12)";
				worksheet.Calculate();
				Assert.AreEqual(727.6944444, (double)worksheet.Cells["A12"].Value, .00001);
			}
		}

		[TestMethod]
		public void VaraIsGivenNumberInputFromCellRefrenceWithEmptyCellsFirstAndAnInvalidCellInTheMiddle()
		{
			var function = new Vara();

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
				worksheet.Cells["A12"].Formula = "=Vara(B1:B11)";
				worksheet.Calculate();
				Assert.AreEqual(748.2678571, (double)worksheet.Cells["A12"].Value, .00001);
			}
		}

		[TestMethod]
		public void VaraIsGivenNumbersAsInputs()
		{
			var function = new Vara();
			var result1 = function.Execute(FunctionsHelper.CreateArgs(-1, 0, 1), this.ParsingContext);
			Assert.AreEqual(1, result1.ResultNumeric, .00001);
		}

		[TestMethod]
		public void VaraIsGivenAMixOfInputTypes()
		{
			var function = new Vara();
			var result1 = function.Execute(FunctionsHelper.CreateArgs(1, true, null, "6/17/2011 2:00", "02:00 am"), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(1, true, "6/17/2011 2:00", "02:00 am"), this.ParsingContext);
			Assert.AreEqual(331469980, result1.ResultNumeric, .1);
			Assert.AreEqual(414333941.1, result2.ResultNumeric, .1);
		}


		[TestMethod]
		public void VaraIsGivenAMixOfInputTypesByCellRange()
		{
			var function = new Vara();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 15;
				worksheet.Cells["B2"].Value = 1;
				worksheet.Cells["B3"].Value = "6/17/2011 2:00";
				worksheet.Cells["B4"].Value = "02:00 am";
				worksheet.Cells["B9"].Formula = "=Vara(B1:B4)";
				worksheet.Calculate();

				Assert.AreEqual(54, (double)worksheet.Cells["B9"].Value, .00001);
			}
		}

		[TestMethod]
		public void VaraIsGivenAMixOfInputTypesWithANullInTheCenterByCellRefrenceAndTwoRangeInputs()
		{
			var function = new Vara();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 15;
				worksheet.Cells["B2"].Value = 1;
				//empty B3 cell
				worksheet.Cells["B4"].Value = "6/17/2011 2:00";
				worksheet.Cells["B5"].Value = "02:00 am";
				worksheet.Cells["B8"].Formula = "=Vara(B1,B2)";
				worksheet.Cells["B9"].Formula = "=Vara(B1:B5,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(98, (double)worksheet.Cells["B8"].Value, .00001);
				Assert.AreEqual(46.28571, (double)worksheet.Cells["B9"].Value, .00001);
			}
		}

		[TestMethod]
		public void VaraIsGivenAMixOfTwoTypesAndTwoRangesThatShouldHaveAnOutput()
		{
			var function = new Vara();
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
				worksheet.Cells["B9"].Formula = "=Vara(B1:B4, A1:A3)";
				worksheet.Calculate();
				Assert.AreEqual(122.4762, (double)worksheet.Cells["B9"].Value, .0001);
			}
		}

		[TestMethod]
		public void VaraIsTheSameTestsAsGivenAMixOfInputTypesByCellRefrenceExceptTheyAreAllOnes()
		{
			var function = new Vara();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 1;
				//empty B3 cell
				worksheet.Cells["B4"].Value = 1;
				worksheet.Cells["B5"].Value = 1;
				worksheet.Cells["B6"].Formula = "=Vara(B1,B2,B4,B5)";
				worksheet.Cells["B7"].Formula = "=Vara(B1,B2,B3,B4,B5)";
				worksheet.Cells["B8"].Formula = "=Vara(B1,B2)";
				worksheet.Cells["B9"].Formula = "=Vara(B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B6"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B7"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B8"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B9"].Value);
			}
		}

		[TestMethod]
		public void VaraTestingDirectInputVsRangeInputWithAGapInTheMiddle()
		{
			var function = new Vara();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B3"].Value = 1;
				worksheet.Cells["B8"].Formula = "=Vara(B1,B3)";
				worksheet.Cells["B9"].Formula = "=Vara(B1:B3)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B8"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B9"].Value);
			}
		}

		[TestMethod]
		public void VaraTestingDirectINputVsRangeInputTest2()
		{
			var function = new Vara();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 0;
				worksheet.Cells["B8"].Formula = "=Vara(B1,B2)";
				worksheet.Cells["B9"].Formula = "=Vara(B1:B2)";
				worksheet.Calculate();
				Assert.AreEqual(0.5, (double)worksheet.Cells["B8"].Value, .00001);
				Assert.AreEqual(0.5, (double)worksheet.Cells["B9"].Value, .00001);
			}
		}

		[TestMethod]
		public void VaraIsGivenAStringInputWithAEmptyCellInTheMiddle()
		{
			var function = new Vara();
			var result1 = function.Execute(FunctionsHelper.CreateArgs(66, 52, 77, 71, 30, 90, 26, 56, 7), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(66, 52, 77, 71, null, 30, 90, 26, 56, 7), this.ParsingContext);
			Assert.AreEqual(727.6944444, result1.ResultNumeric, .00001);
			Assert.AreEqual(925.3888889, result2.ResultNumeric, .00001);
		}

		[TestMethod]
		public void VaraIsGivenAMixedStringInputWithAEmptyCellInTheMiddle()
		{
			var function = new Vara();
			var result1 = function.Execute(FunctionsHelper.CreateArgs(10, 2, "6/17/2011 2:00", "02:00 am"), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs(10, 2, "6/17/2011 2:00", null, "02:00 am"), this.ParsingContext);
			Assert.AreEqual(414266111.4, result1.ResultNumeric, .1);
			Assert.AreEqual(331429287.4, result2.ResultNumeric, .1);
		}

		[TestMethod]
		public void VaraIsGivenAMixOfTwoTypesAndTwoRangesThatShouldHaveAnOutputTestTwo()
		{
			var function = new Vara();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 10;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = "6/17/2011 2:00";
				worksheet.Cells["B4"].Value = "02:00 am";
				worksheet.Cells["B9"].Formula = "=Vara(B1:B4)";
				worksheet.Calculate();
				Assert.AreEqual(22.66666667, (double)worksheet.Cells["B9"].Value, .0001);
			}
		}

		[TestMethod]
		public void VaraIsGivenTwoRangesOfIntsAsInputs()
		{
			var function = new Vara();
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


				worksheet.Cells["A9"].Formula = "=Vara(B1:B6, B7:B11)";
				worksheet.Calculate();
				Assert.AreEqual(692.0909091, (double)worksheet.Cells["A9"].Value, .00001);
			}
		}

		[TestMethod]
		public void VaraIsGivenTwoRangesOfIntsAsInputsWithATimeInTheMiddleOfTheFirstRange()
		{
			var function = new Vara();
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


				worksheet.Cells["A9"].Formula = "=Vara(B1:B6, B7:B11)";
				worksheet.Calculate();
				Assert.AreEqual(617.7636364, (double)worksheet.Cells["A9"].Value, .00001);
			}
		}

		[TestMethod]
		public void VaraIsGivenTwoRangesOfIntsAsInputsWithATimeInTheMiddleOfTheFirstRangeTestTwo()
		{
			var function = new Vara();
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


				worksheet.Cells["A9"].Formula = "=Vara(B1:B6, B7:B11)";
				worksheet.Calculate();
				Assert.AreEqual(626.4909091, (double)worksheet.Cells["A9"].Value, .00001);
			}
		}

		[TestMethod]
		public void VaraIsGivenAStringInputWithAEmptyCellInTheMiddleByCellRefrence()
		{
			var function = new Vara();
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
				worksheet.Cells["A12"].Formula = "=Vara(B3:B12)";
				worksheet.Calculate();
				Assert.AreEqual(727.6944444, (double)worksheet.Cells["A12"].Value, .00001);
			}
		}

		[TestMethod]
		public void VaraIsGivenMilitaryTimesAsInputs()
		{
			var function = new Vara();
			var result1 = function.Execute(FunctionsHelper.CreateArgs("00:00", "02:00", "13:00"), this.ParsingContext);
			Assert.AreEqual(0.085069444, result1.ResultNumeric, .00001);
		}

		[TestMethod]
		public void VaraIsGivenNumbersAsInputstakeTwo()
		{
			var function = new Vara();
			var result1 = function.Execute(FunctionsHelper.CreateArgs(-1, -1, -1), this.ParsingContext);
			Assert.AreEqual(0, result1.ResultNumeric, .00001);
		}

		[TestMethod]
		public void VaraIsGivenMilitaryTimesAsInputsByCellRefrence()
		{
			var function = new Vara();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "00:00";
				worksheet.Cells["B2"].Value = "02:00";
				worksheet.Cells["B3"].Value = "13:00";
				worksheet.Cells["B4"].Formula = "=Vara(B1,B2,B3)";
				worksheet.Calculate();
				Assert.AreEqual(0d, (worksheet.Cells["B4"].Value));
			}
		}

		[TestMethod]
		public void VaraIsGivenMilitaryTimesAsInputsByCellRange()
		{
			var function = new Vara();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "00:00";
				worksheet.Cells["B2"].Value = "02:00";
				worksheet.Cells["B3"].Value = "13:00";
				worksheet.Cells["B4"].Formula = "=Vara(B1:B3)";
				worksheet.Calculate();
				Assert.AreEqual(0d, (worksheet.Cells["B4"].Value));
			}
		}

		[TestMethod]
		public void VaraIsGiven12HourTimesAsInputs()
		{
			var function = new Vara();
			var result1 = function.Execute(FunctionsHelper.CreateArgs("12:00 am", "02:00 am", "01:00 pm"), this.ParsingContext);
			Assert.AreEqual(0.085069444, result1.ResultNumeric, .00001);
		}

		[TestMethod]
		public void VaraIsGiven12HourTimesAsInputsByCellRefrence()
		{
			var function = new Vara();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "12:00 am";
				worksheet.Cells["B2"].Value = "02:00 am";
				worksheet.Cells["B3"].Value = "01:00 pm";
				worksheet.Cells["B4"].Formula = "=Vara(B1,B2,B3)";
				worksheet.Calculate();
				Assert.AreEqual(0d, (worksheet.Cells["B4"].Value));
			}
		}

		[TestMethod]
		public void VaraIsGiven12HourTimesAsInputsByCellRange()
		{
			var function = new Vara();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "12:00 am";
				worksheet.Cells["B2"].Value = "02:00 am";
				worksheet.Cells["B3"].Value = "01:00 pm";
				worksheet.Cells["B4"].Formula = "=Vara(B1:B3)";
				worksheet.Calculate();
				Assert.AreEqual(0d, (worksheet.Cells["B4"].Value));
			}
		}

		[TestMethod]
		public void VaraIsGivenMonthDayYear12HourTimeAsInputs()
		{
			var function = new Vara();
			var result1 = function.Execute(FunctionsHelper.CreateArgs("Jan 17, 2011 2:00 am", "June 5, 2017 11:00 pm", "June 15, 2017 11:00 pm"), this.ParsingContext);
			Assert.AreEqual(1820353.255, result1.ResultNumeric, .001);
		}

		[TestMethod]
		public void VaraIsGivenMonthDayYear12HourTimeAsInputsByCellRefrence()
		{
			var function = new Vara();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "Jan 17, 2011 2:00 am";
				worksheet.Cells["B2"].Value = "June 5, 2017 11:00 pm";
				worksheet.Cells["B3"].Value = "June 15, 2017 11:00 pm";
				worksheet.Cells["B4"].Formula = "=Vara(B1,B2,B3)";
				worksheet.Calculate();
				Assert.AreEqual(0d, (worksheet.Cells["B4"].Value));
			}
		}

		[TestMethod]
		public void VaraIsGivenMonthDayYear12HourTimeAsInputsByCellRange()
		{
			var function = new Vara();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "Jan 17, 2011 2:00 am";
				worksheet.Cells["B2"].Value = "June 5, 2017 11:00 pm";
				worksheet.Cells["B3"].Value = "June 15, 2017 11:00 pm";
				worksheet.Cells["B4"].Formula = "=Vara(B1:B3)";
				worksheet.Calculate();
				Assert.AreEqual(0d, (worksheet.Cells["B4"].Value));
			}
		}

		[TestMethod]
		public void VaraIsGivenDateTimeInputsSeperatedByADashAsInputs()
		{
			var function = new Vara();
			var result1 = function.Execute(FunctionsHelper.CreateArgs("1-17-2017 2:00", "6-17-2017 2:00", "9-17-2017 2:00"), this.ParsingContext);
			Assert.AreEqual(15052.33333, result1.ResultNumeric, .00001);
		}

		[TestMethod]
		public void VaraIsGivenDateTimeInputsSeperatedByADashAsInputsByCellRefrence()
		{
			var function = new Vara();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "1-17-2017 2:00";
				worksheet.Cells["B2"].Value = "6-17-2017 2:00";
				worksheet.Cells["B3"].Value = "9-17-2017 2:00";
				worksheet.Cells["B4"].Formula = "=Vara(B1,B2,B3)";
				worksheet.Calculate();
				Assert.AreEqual(0d, (worksheet.Cells["B4"].Value));
			}
		}

		[TestMethod]
		public void VaraIsGivenDateTimeInputsSeperatedByADashAsInputsByCellRange()
		{
			var function = new Vara();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "1-17-2017 2:00";
				worksheet.Cells["B2"].Value = "6-17-2017 2:00";
				worksheet.Cells["B3"].Value = "9-17-2017 2:00";
				worksheet.Cells["B4"].Formula = "=Vara(B1:B3)";
				worksheet.Calculate();
				Assert.AreEqual(0d, (worksheet.Cells["B4"].Value));
			}
		}

		[TestMethod]
		public void VaraIsGivenStringsAsInputs()
		{
			var function = new Vara();
			var result1 = function.Execute(FunctionsHelper.CreateArgs("string", "another string", "a third string"), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result1.Result).Type);
		}

		[TestMethod]
		public void VaraIsGivenStringsAsInputsByCellRefrence()
		{
			var function = new Vara();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "string";
				worksheet.Cells["B2"].Value = "another string";
				worksheet.Cells["B3"].Value = "a third string";
				worksheet.Cells["B4"].Formula = "=Vara(B1,B2,B3)";
				worksheet.Calculate();
				Assert.AreEqual(0d, (worksheet.Cells["B4"].Value));
			}
		}

		[TestMethod]
		public void VaraIsGivenStringsAsInputsByCellRange()
		{
			var function = new Vara();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "string";
				worksheet.Cells["B2"].Value = "another string";
				worksheet.Cells["B3"].Value = "a third string";
				worksheet.Cells["B4"].Formula = "=Vara(B1:B3)";
				worksheet.Calculate();
				Assert.AreEqual(0d, (worksheet.Cells["B4"].Value));
			}
		}

		[TestMethod]
		public void VaraIsGivenStringNumbersAsInputs()
		{
			var function = new Vara();
			var result1 = function.Execute(FunctionsHelper.CreateArgs("5", "6", "7"), this.ParsingContext);
			var result2 = function.Execute(FunctionsHelper.CreateArgs("5.5", "6.6", "7.7"), this.ParsingContext);
			Assert.AreEqual(1, result1.ResultNumeric, .00001);
			Assert.AreEqual(1.21, result2.ResultNumeric, .00001);
		}

		[TestMethod]
		public void VaraIsGivenStringNumbersAsInputsByCellRefrence()
		{
			var function = new Vara();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["A1"].Value = "5";
				worksheet.Cells["A2"].Value = "6";
				worksheet.Cells["A3"].Value = "7";
				worksheet.Cells["B1"].Value = "5.5";
				worksheet.Cells["B2"].Value = "6.6";
				worksheet.Cells["B3"].Value = "7.7";
				worksheet.Cells["A4"].Formula = "=Vara(A1,A2,A3)";
				worksheet.Cells["B4"].Formula = "=Vara(B1,B2,B3)";
				worksheet.Calculate();
				Assert.AreEqual(0d, (worksheet.Cells["B4"].Value));
				Assert.AreEqual(0d, (worksheet.Cells["B4"].Value));
			}
		}

		[TestMethod]
		public void VaraIsGivenStringNumbersAsInputsByCellRange()
		{
			var function = new Vara();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["A1"].Value = "5";
				worksheet.Cells["A2"].Value = "6";
				worksheet.Cells["A3"].Value = "7";
				worksheet.Cells["B1"].Value = "5.5";
				worksheet.Cells["B2"].Value = "6.6";
				worksheet.Cells["B3"].Value = "7.7";
				worksheet.Cells["A4"].Formula = "=Vara(A1:A3)";
				worksheet.Cells["B4"].Formula = "=Vara(B1:B3)";
				worksheet.Calculate();
				Assert.AreEqual(0d, (worksheet.Cells["B4"].Value));
				Assert.AreEqual(0d, (worksheet.Cells["B4"].Value));
			}
		}

		[TestMethod]
		public void VaraIsGivenASingleTrueBooleanInput()
		{
			var function = new Vara();
			var result1 = function.Execute(FunctionsHelper.CreateArgs(true), this.ParsingContext);

			Assert.AreEqual(0, result1.ResultNumeric, .00001); ;
		}

		[TestMethod]
		public void VaraIsGivenASingleFalseBooleanInput()
		{
			var function = new Vara();
			var result1 = function.Execute(FunctionsHelper.CreateArgs(false), this.ParsingContext);

			Assert.AreEqual(0, result1.ResultNumeric, .00001); ;
		}

		[TestMethod]
		public void VaraIsGivenASingleStringInput()
		{
			var function = new Vara();
			var result1 = function.Execute(FunctionsHelper.CreateArgs("string"), this.ParsingContext);

			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result1.Result).Type);
		}
		#endregion
	}
}
