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
using System.Linq;
using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math
{
	[TestClass]
	public class SumTests : MathFunctionsTestBase
	{
		#region Sum Function (Execute) Tests
		[TestMethod]
		public void SumWithFourNumbersReturnsCorrectValue()
		{
			var function = new Sum();
			var result = function.Execute(FunctionsHelper.CreateArgs(1.5, 2, 3.5, 7), this.ParsingContext);
			Assert.AreEqual(14d, result.Result);
		}

		[TestMethod]
		public void SumWithFourNegativeNumbersReturnsCorrectValue()
		{
			var function = new Sum();
			var result = function.Execute(FunctionsHelper.CreateArgs(-1.5, -2, -3.5, -7), this.ParsingContext);
			Assert.AreEqual(-14d, result.Result);
		}

		[TestMethod]
		public void SumWithOneIntegerReturnsCorrectValue()
		{
			var function = new Sum();
			var result = function.Execute(FunctionsHelper.CreateArgs(2), this.ParsingContext);
			Assert.AreEqual(2d, result.Result);
		}

		[TestMethod]
		public void SumWithOneDoubleReturnsCorrectValue()
		{
			var function = new Sum();
			var result = function.Execute(FunctionsHelper.CreateArgs(2.5), this.ParsingContext);
			Assert.AreEqual(2.5d, result.Result);
		}

		[TestMethod]
		public void SumWithOneNumericStringReturnsCorrectValue()
		{
			var function = new Sum();
			var result = function.Execute(FunctionsHelper.CreateArgs("2"), this.ParsingContext);
			Assert.AreEqual(2d, result.Result);
		}

		[TestMethod]
		public void SumWithNonNumericStringReturnsPoundValue()
		{
			var function = new Sum();
			var result = function.Execute(FunctionsHelper.CreateArgs("word"), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SumWithDateInStringReturnsCorrectValue()
		{
			var function = new Sum();
			var result = function.Execute(FunctionsHelper.CreateArgs("6/16/2017"), this.ParsingContext);
			Assert.AreEqual(42902d, result.Result);
		}

		[TestMethod]
		public void SumWithBooleanValueReturnsCorrectValue()
		{
			var function = new Sum();
			var booleanTrue = function.Execute(FunctionsHelper.CreateArgs(true), this.ParsingContext);
			var booleanFalse = function.Execute(FunctionsHelper.CreateArgs(false), this.ParsingContext);
			Assert.AreEqual(1d, booleanTrue.Result);
			Assert.AreEqual(0d, booleanFalse.Result);
		}

		[TestMethod]
		public void SumWithErrorInputsReturnRespectiveErrors()
		{
			var func = new Sum();
			var argNA = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.NA), 5);
			var argNAME = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Name), 5);
			var argVALUE = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Value), 5);
			var argNUM = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Num), 5);
			var argDIV0 = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Div0), 5);
			var argREF = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Ref), 5);
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

		[TestMethod]
		public void SumWithOneIntegerCellReferenceReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 2;
				worksheet.Cells["B2"].Formula = "SUM(B1)";
				worksheet.Calculate();
				Assert.AreEqual(2d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void SumWithOneDoubleCellReferenceReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 2.5;
				worksheet.Cells["B2"].Formula = "SUM(B1)";
				worksheet.Calculate();
				Assert.AreEqual(2.5d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void SumWithNumericStringCellReferenceReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "2";
				worksheet.Cells["B2"].Formula = "SUM(B1)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void SumWithNonNumericStringCellReferenceReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "word";
				worksheet.Cells["B2"].Formula = "SUM(B1)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void SumWithDateInStringCellReferenceReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "6/16/2017";
				worksheet.Cells["B2"].Formula = "SUM(B1)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void SumWithBooleanCellReferenceReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = true;
				worksheet.Cells["B2"].Formula = "SUM(B1)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void SumWithEmptyCellReferenceReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "SUM(A2)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void SumWithIntegerAndNumericStringReturnsCorrectValue()
		{
			var function = new Sum();
			var result = function.Execute(FunctionsHelper.CreateArgs("1", 3), this.ParsingContext);
			Assert.AreEqual(4d, result.Result);
		}

		[TestMethod]
		public void SumWithIntegerAndNonNumericStringReturnsCorrectValue()
		{
			var function = new Sum();
			var result = function.Execute(FunctionsHelper.CreateArgs(1, "word"), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SumWithIntegerAndBooleanReturnsCorrectValue()
		{
			var function = new Sum();
			var result = function.Execute(FunctionsHelper.CreateArgs(3, true), this.ParsingContext);
			Assert.AreEqual(4d, result.Result);
		}

		[TestMethod]
		public void SumWithIntegerAndNullValueReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "SUM(1, )";
				worksheet.Calculate();
				Assert.AreEqual(1d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void SumWithIntegerAndNumericStringCellReferenceReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "2";
				worksheet.Cells["B2"].Formula = "SUM(1, B1)";
				worksheet.Calculate();
				Assert.AreEqual(1d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void SumWithIntegerAndNonNumericStringCellReferenceReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "word";
				worksheet.Cells["B2"].Formula = "SUM(1, B1)";
				worksheet.Calculate();
				Assert.AreEqual(1d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void SumWithIntegerAndBooleanCellReferenceReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = true;
				worksheet.Cells["B2"].Formula = "SUM(1, B1)";
				worksheet.Calculate();
				Assert.AreEqual(1d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void SumWithIntegerAndEmptyCellReferenceReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "SUM(2, B2)";
				worksheet.Calculate();
				Assert.AreEqual(2d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void SumWithNumericStringAndNonNumericStringReturnsPoundValue()
		{
			var function = new Sum();
			var result = function.Execute(FunctionsHelper.CreateArgs("2", "Word"), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SumWithNonNumericStringAndNumericStringReturnsPoundValue()
		{
			var function = new Sum();
			var result = function.Execute(FunctionsHelper.CreateArgs("word", "2"), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SumWithBooleanValueAndNonNumericStringReturnsPoundValue()
		{
			var function = new Sum();
			var result = function.Execute(FunctionsHelper.CreateArgs(true, "word"), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SumWthNonNumericStringAndBooleanValueReturnsPoundValue()
		{
			var function = new Sum();
			var result = function.Execute(FunctionsHelper.CreateArgs("word", true), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SumWithNumericStringAndBooleanValueReturnsCorrectValue()
		{
			var function = new Sum();
			var result = function.Execute(FunctionsHelper.CreateArgs("2", true), this.ParsingContext);
			Assert.AreEqual(3d, result.Result);
		}

		[TestMethod]
		public void SumWithBooleanValueAndNumericStringReturnsCorrectValue()
		{
			var function = new Sum();
			var result = function.Execute(FunctionsHelper.CreateArgs(true, "2"), this.ParsingContext);
			Assert.AreEqual(3d, result.Result);
		}

		[TestMethod]
		public void SumWithNumericStringAndNonNumericStringCellReferenceReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "2";
				worksheet.Cells["B2"].Value = "word";
				worksheet.Cells["B3"].Formula = "SUM(B1:B2)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B3"].Value);
			}
		}

		[TestMethod]
		public void SumWithNonNumericStringAndNumericStringCellReferenceReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "word";
				worksheet.Cells["B2"].Value = "2";
				worksheet.Cells["B3"].Formula = "SUM(B1:B2)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B3"].Value);
			}
		}

		[TestMethod]
		public void SumWithBooleanAndNonNumericStringCellReferenceReturnsZero()
		{
			using (var packge = new ExcelPackage())
			{
				var worksheet = packge.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = true;
				worksheet.Cells["B2"].Value = "word";
				worksheet.Cells["B3"].Formula = "SUM(B1:B2)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B3"].Value);
			}
		}

		[TestMethod]
		public void SumWithNonNumericStringAndBooleanCellReferenceReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "word";
				worksheet.Cells["B2"].Value = true;
				worksheet.Cells["B3"].Formula = "SUM(B1:B2)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B3"].Value);
			}
		}

		[TestMethod]
		public void SumWithNumericStringAndBooleanValueCellReferenceReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "2";
				worksheet.Cells["B2"].Value = true;
				worksheet.Cells["B3"].Formula = "SUM(B1:B2)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B3"].Value);
			}
		}

		[TestMethod]
		public void SumWithBooleanValueAndNumericStringCellReferenceReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = true;
				worksheet.Cells["B2"].Value = "2";
				worksheet.Cells["B3"].Formula = "SUM(B1:B2)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B3"].Value);
			}
		}

		[TestMethod]
		public void SumWithCellRangeReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B2"].Value = 1.5;
				worksheet.Cells["B3"].Value = 2;
				worksheet.Cells["B4"].Value = 3.5;
				worksheet.Cells["B5"].Value = 7;
				worksheet.Cells["B6"].Formula = "SUM(B2:B5)";
				worksheet.Calculate();
				Assert.AreEqual(14d, worksheet.Cells["B6"].Value);
			}
		}

		[TestMethod]
		public void SumWithIntegersAnNumericStringCellReferenceReturnCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 1.5;
				worksheet.Cells["B2"].Value = "2";
				worksheet.Cells["B3"].Value = 3.5;
				worksheet.Cells["B4"].Value = 7;
				worksheet.Cells["B5"].Formula = "SUM(B1:B4)";
				worksheet.Calculate();
				Assert.AreEqual(12d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void SumWithIntegersAndNonNumericStringCellReferenceReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 1.5;
				worksheet.Cells["B2"].Value = "word";
				worksheet.Cells["B3"].Value = 3.5;
				worksheet.Cells["B4"].Value = 7;
				worksheet.Cells["B5"].Formula = "SUM(B1:B4)";
				worksheet.Calculate();
				Assert.AreEqual(12d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void SumWithIntegersAndBooleanValueCellReferenceReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 1.5;
				worksheet.Cells["B2"].Value = true;
				worksheet.Cells["B3"].Value = 3.5;
				worksheet.Cells["B4"].Value = 7;
				worksheet.Cells["B5"].Formula = "SUM(B1:B4)";
				worksheet.Calculate();
				Assert.AreEqual(12d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void SumWithIntegersAndEmptyCellReferenceReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 1.5;
				worksheet.Cells["B3"].Value = 3.5;
				worksheet.Cells["B4"].Value = 7;
				worksheet.Cells["B5"].Formula = "SUM(B1:B4)";
				worksheet.Calculate();
				Assert.AreEqual(12d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void SumWithNumericStringInArrayReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "SUM({\"1\", \"2\", \"3\"})";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void SumWithNumericStringsEnteredDirectlyReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "SUM(\"1\", \"2\", \"3\")";
				worksheet.Calculate();
				Assert.AreEqual(6d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void SumWithNumbersInArrayReturnsCorectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "SUM({1, 2, 3})";
				worksheet.Calculate();
				Assert.AreEqual(6d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void SumWithNumberAsValueReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 2;
				worksheet.Cells["B2"].Formula = "SUM(B1)";
				worksheet.Calculate();
				Assert.AreEqual(2d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void SumWithNumberAsFormulaReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "2";
				worksheet.Cells["B2"].Formula = "SUM(B1)";
				worksheet.Calculate();
				Assert.AreEqual(2d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void SumWithStringAsValueReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "2";
				worksheet.Cells["B2"].Formula = "SUM(B1)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void SumWithStringAsFormulaReturnsZero()
		{
			using (var packge = new ExcelPackage())
			{
				var worksheet = packge.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "\"2\"";
				worksheet.Cells["B2"].Formula = "SUM(B1)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void SumWithBooleanAsValueReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = true;
				worksheet.Cells["B2"].Formula = "SUM(B1)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void SumWithBooleanAsFormulaReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "TRUE";
				worksheet.Cells["B2"].Formula = "SUM(B1)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void SumWithDateInStringAsValueReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "6/20/2017";
				worksheet.Cells["B2"].Formula = "SUM(B1)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void SumWithDateInStringAsFormulaReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "\"6/20/2017\"";
				worksheet.Cells["B2"].Formula = "SUM(B1)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void SumShouldCalculate2Plus3AndReturn5()
		{
			var func = new Sum();
			var args = FunctionsHelper.CreateArgs(2, 3);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(5d, result.Result);
		}

		[TestMethod]
		public void SumShouldCalculateEnumerableOf2Plus5Plus3AndReturn10()
		{
			var func = new Sum();
			var args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(2, 5), 3);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(10d, result.Result);
		}

		[TestMethod]
		public void SumShouldIgnoreHiddenValuesWhenIgnoreHiddenValuesIsSet()
		{
			var func = new Sum();
			func.IgnoreHiddenValues = true;
			var args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(2, 5), 3, 4);
			args.Last().SetExcelStateFlag(ExcelCellState.HiddenCell);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(10d, result.Result);
		}

		[TestMethod]
		public void SumPropagatesRangeErrors()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet1");
				sheet.Cells[2, 2].Value = "#NAME?";
				sheet.Cells[2, 3].Value = 0;
				sheet.Cells[2, 4].Value = 1;
				sheet.Cells[2, 5].Formula = "SUM(B2:D2)";
				sheet.Calculate();
				Assert.AreEqual(eErrorType.Name, ((ExcelErrorValue)sheet.Cells[2, 5].Value).Type);
			}
		}

		[TestMethod]
		public void SumIgnoresTextThatLooksLikeAnErrorValue()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet1");
				sheet.Cells[2, 2].Formula = @"""#NAME?""";
				sheet.Cells[2, 3].Value = 0;
				sheet.Cells[2, 4].Value = 1;
				sheet.Cells[2, 5].Formula = "SUM(B2:D2)";
				sheet.Calculate();
				Assert.AreEqual(1d, sheet.Cells[2, 5].Value);
			}
		}
		#endregion
	}
}
