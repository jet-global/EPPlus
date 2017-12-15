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
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math
{
	[TestClass]
	public class MedianTests : MathFunctionsTestBase
	{
		#region Median Function (Execute) Tests
		[TestMethod]
		public void MedianWithNoArgumentsReturnsPoundValue()
		{
			var function = new Median();
			var result = function.Execute(FunctionsHelper.CreateArgs(), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void MedianWithMaxArgumentsReturnsCorrectValue()
		{
			// This functionality is different from that of Excel's. Normally when too many arguments are entered
			// into a function it won't let you calculate the function, however in EPPlus it will return a pound
			// NA error instead. 
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				for(int i = 1; i < 270; i++)
				{
					for (int j = 1; j < 2; j++)
					{
						worksheet.Cells[i, j].Value = 3;
					}
				}
				worksheet.Cells["B3"].Formula = "MEDIAN(A1:A255)";
				worksheet.Cells["B4"].Formula = "MEDIAN(A1:A270)";
				worksheet.Calculate();
				Assert.AreEqual(3d, worksheet.Cells["B3"].Value);
				Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)worksheet.Cells["B4"].Value).Type);
			}
		}	

		[TestMethod]
		public void MedianWithOneInputReturnsCorrectValue()
		{
			var function = new Median();
			var result = function.Execute(FunctionsHelper.CreateArgs(15), this.ParsingContext);
			Assert.AreEqual(15d, result.Result);
		}

		[TestMethod]
		public void MedianWithNumericInputsReturnsCorrectValue()
		{
			var function = new Median();
			var result = function.Execute(FunctionsHelper.CreateArgs(16, 55, 19, 20), this.ParsingContext);
			Assert.AreEqual(19.5d, result.Result);
		}

		[TestMethod]
		public void MedianWithGenericStringInputReturnsPoundValue()
		{
			var function = new Median();
			var result = function.Execute(FunctionsHelper.CreateArgs("string", "string", "string"), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void MedianWithNumericStringInputReturnsCorrectValue()
		{
			var function = new Median();
			var result = function.Execute(FunctionsHelper.CreateArgs("16", "55", "19", "20"), this.ParsingContext);
			Assert.AreEqual(19.5d, result.Result);
		}

		[TestMethod]
		public void MedianWithReferenceToNumbersReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 16;
				worksheet.Cells["B2"].Value = 6;
				worksheet.Cells["B3"].Value = 5;
				worksheet.Cells["B10"].Formula = "MEDIAN(B1:B3)";
				worksheet.Calculate();
				Assert.AreEqual(6d, worksheet.Cells["B10"].Value);
			}
		}

		[TestMethod]
		public void MedianWithReferencesTypedOutReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 16;
				worksheet.Cells["B2"].Value = 6;
				worksheet.Cells["B3"].Value = 5;
				worksheet.Cells["B10"].Formula = "MEDIAN(B1,B2,B3)";
				worksheet.Calculate();
				Assert.AreEqual(6d, worksheet.Cells["B10"].Value);
			}
		}

		[TestMethod]
		public void MedianWithReferenceToNumericStringsReturnsPoundNum()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "5";
				worksheet.Cells["B2"].Value = "45";
				worksheet.Cells["B3"].Value = "76";
				worksheet.Cells["B10"].Formula = "MEDIAN(B1:B3)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)worksheet.Cells["B10"].Value).Type);
			}
		}

		[TestMethod]
		public void MedianWithReferencesToGeneralStringsReturnsPoundNum()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "string!";
				worksheet.Cells["B2"].Value = "string";
				worksheet.Cells["B3"].Value = "string";
				worksheet.Cells["B10"].Formula = "MEDIAN(B1:B3)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)worksheet.Cells["B10"].Value).Type);
			}
		}

		[TestMethod]
		public void MedianWithLogicInputsReturnsCorrectValue()
		{
			var function = new Median();
			var result = function.Execute(FunctionsHelper.CreateArgs(true, false), this.ParsingContext);
			Assert.AreEqual(0.5d, result.Result);
		}

		[TestMethod]
		public void MedianWithReferenceToLogicInputsReturnsPoundNum()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "TRUE";
				worksheet.Cells["B2"].Value = "FALSE";
				worksheet.Cells["B10"].Formula = "MEDIAN(B1:B2)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)worksheet.Cells["B10"].Value).Type);
			}
		}

		[TestMethod]
		public void MedianWithReferenceToCellsWithZeroReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 0;
				worksheet.Cells["B2"].Value = 16;
				worksheet.Cells["B3"].Value = 6;
				worksheet.Cells["B4"].Value = 5;
				worksheet.Cells["B10"].Formula = "MEDIAN(B1:B4)";
				worksheet.Calculate();
				Assert.AreEqual(5.5d, worksheet.Cells["B10"].Value);
			}
		}

		[TestMethod]
		public void MedianWithInputCellsThatHaveErrorsReturnsRespectiveError()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 99;
				worksheet.Cells["B2"].Value = 6;
				worksheet.Cells["B3"].Formula = "SQRT(-1)";
				worksheet.Cells["B10"].Formula = "MEDIAN(B1:B3)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)worksheet.Cells["B10"].Value).Type);
			}
		}

		[TestMethod]
		public void MedianWithReferenceToEmptyCellsReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 16;
				worksheet.Cells["B2"].Value = 5;
				worksheet.Cells["B3"].Value = 6;
				worksheet.Cells["B10"].Formula = "MEDIAN(B1:B4)";
				worksheet.Calculate();
				Assert.AreEqual(6d, worksheet.Cells["B10"].Value);
			}
		}

		[TestMethod]
		public void MedianWithReferenceToStringsAndNumbersReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 5;
				worksheet.Cells["B2"].Value = 64;
				worksheet.Cells["B3"].Value = 0;
				worksheet.Cells["B4"].Value = "string";
				worksheet.Cells["B10"].Formula = "MEDIAN(B1:B4)";
				worksheet.Calculate();
				Assert.AreEqual(5d, worksheet.Cells["B10"].Value);
			}
		}

		[TestMethod]
		public void MedianWithDateObjectAsInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "MEDIAN(DATE(2017, 6, 12))";
				worksheet.Calculate();
				Assert.AreEqual(42898d, worksheet.Cells["B1"].Value);
			}
		}
		
		[TestMethod]
		public void MedianWithDateAsStringReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "MEDIAN(\"6/12/2017\", \"5/4/2016\")";
				worksheet.Calculate();
				Assert.AreEqual(42696d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void MedianWithDoubleInputsReturnsCorrectValue()
		{
			var function = new Median();
			var result = function.Execute(FunctionsHelper.CreateArgs(5.5, 2.3, 15.6, 11.2), this.ParsingContext);
			Assert.AreEqual(8.35d, result.Result);
		}

		[TestMethod]
		public void MedianWithFractionInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "MEDIAN((2/3),(5/8),(99/5))";
				worksheet.Calculate();
				Assert.AreEqual(0.66666666666667d, (double)worksheet.Cells["B1"].Value, 0.000001);
			}
		}

		[TestMethod]
		public void MedianWithMixedTypesReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "MEDIAN(4, TRUE, \"78\")";
				worksheet.Calculate();
				Assert.AreEqual(4d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void MedianWithCommaButNoArgsReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "MEDIAN(,)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void MedianWithSecondArgOnlyReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "MEDIAN(, 5)";
				worksheet.Calculate();
				Assert.AreEqual(2.5d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void MedianWithOnlyFirstArgReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "MEDIAN(1, )";
				worksheet.Cells["B2"].Formula = "MEDIAN(1, , , ,)";
				worksheet.Calculate();
				Assert.AreEqual(0.5d, worksheet.Cells["B1"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void MedianShouldPoundNumIfNoArgs()
		{
			var function = new Median();
			var arguments = new FunctionArgument[] { new FunctionArgument(null) };
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(OfficeOpenXml.FormulaParsing.ExpressionGraph.DataType.ExcelError, result.DataType);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)(result.Result)).Type);
		}

		[TestMethod]
		public void MedianShouldCalculateCorrectlyWithOneMember()
		{
			var function = new Median();
			var arguments = FunctionsHelper.CreateArgs(1);
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(1d, result.Result);
		}

		[TestMethod]
		public void MedianShouldCalculateCorrectlyWithOddMembers()
		{
			var function = new Median();
			var arguments = FunctionsHelper.CreateArgs(3, 5, 1, 4, 2);
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(3d, result.Result);
		}

		[TestMethod]
		public void MedianShouldCalculateCorrectlyWithEvenMembers()
		{
			var function = new Median();
			var arguments = FunctionsHelper.CreateArgs(1, 2, 3, 4);
			var result = function.Execute(arguments, this.ParsingContext);
			Assert.AreEqual(2.5d, result.Result);
		}
		#endregion
	}
}
