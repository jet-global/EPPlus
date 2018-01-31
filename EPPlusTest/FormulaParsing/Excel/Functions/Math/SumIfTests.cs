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
using System;
using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using static OfficeOpenXml.FormulaParsing.ExcelDataProvider;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math
{
	[TestClass]
	public class SumIfTests : MathFunctionsTestBase
	{
		#region SumIf Function (Execute) Tests
		private ExcelPackage _package;
		private EpplusExcelDataProvider _provider;
		private ParsingContext _parsingContext;
		private ExcelWorksheet _worksheet;

		[TestInitialize]
		public void Initialize()
		{
			_package = new ExcelPackage();
			_provider = new EpplusExcelDataProvider(_package);
			_parsingContext = ParsingContext.Create();
			_parsingContext.Scopes.NewScope(RangeAddress.Empty);
			_worksheet = _package.Workbook.Worksheets.Add("testsheet");
		}

		[TestCleanup]
		public void Cleanup()
		{
			_package.Dispose();
		}

		[TestMethod]
		public void SumIfNumeric()
		{
			_worksheet.Cells["A1"].Value = 1d;
			_worksheet.Cells["A2"].Value = 2d;
			_worksheet.Cells["A3"].Value = 3d;
			_worksheet.Cells["B1"].Value = 1d;
			_worksheet.Cells["B2"].Value = 3d;
			_worksheet.Cells["B3"].Value = 5d;
			var func = new SumIf();
			IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
			IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
			var args = FunctionsHelper.CreateArgs(range1, ">1", range2);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(8d, result.Result);
		}

		[TestMethod]
		public void SumIfNonNumeric()
		{
			_worksheet.Cells["A1"].Value = "Monday";
			_worksheet.Cells["A2"].Value = "Tuesday";
			_worksheet.Cells["A3"].Value = "Thursday";
			_worksheet.Cells["B1"].Value = 1d;
			_worksheet.Cells["B2"].Value = 3d;
			_worksheet.Cells["B3"].Value = 5d;
			var func = new SumIf();
			IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
			IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
			var args = FunctionsHelper.CreateArgs(range1, "T*day", range2);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(8d, result.Result);
		}

		[TestMethod]
		public void SumIfNumericExpression()
		{
			_worksheet.Cells["A1"].Value = null;
			_worksheet.Cells["A2"].Value = 1d;
			_worksheet.Cells["A3"].Value = "Not Empty";
			var func = new SumIf();
			IRangeInfo range = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
			var args = FunctionsHelper.CreateArgs(range, 1d);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(1d, result.Result);
		}

		[TestMethod]
		public void SumIfNotEqualToNull()
		{
			_worksheet.Cells["A1"].Value = null;
			_worksheet.Cells["A2"].Value = string.Empty;
			_worksheet.Cells["A3"].Value = "Not Empty";
			_worksheet.Cells["B1"].Value = 1d;
			_worksheet.Cells["B2"].Value = 3d;
			_worksheet.Cells["B3"].Value = 5d;
			var func = new SumIf();
			IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
			IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
			var args = FunctionsHelper.CreateArgs(range1, "<>", range2);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(8d, result.Result);
		}

		[TestMethod]
		public void SumIfEqualToZero()
		{
			_worksheet.Cells["A1"].Value = null;
			_worksheet.Cells["A2"].Value = string.Empty;
			_worksheet.Cells["A3"].Value = 0d;
			_worksheet.Cells["B1"].Value = 1d;
			_worksheet.Cells["B2"].Value = 3d;
			_worksheet.Cells["B3"].Value = 5d;
			var func = new SumIf();
			IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
			IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
			var args = FunctionsHelper.CreateArgs(range1, "0", range2);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(5d, result.Result);
		}

		[TestMethod]
		public void SumIfNotEqualToZero()
		{
			_worksheet.Cells["A1"].Value = null;
			_worksheet.Cells["A2"].Value = string.Empty;
			_worksheet.Cells["A3"].Value = 0d;
			_worksheet.Cells["B1"].Value = 1d;
			_worksheet.Cells["B2"].Value = 3d;
			_worksheet.Cells["B3"].Value = 5d;
			var func = new SumIf();
			IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
			IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
			var args = FunctionsHelper.CreateArgs(range1, "<>0", range2);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(4d, result.Result);
		}

		[TestMethod]
		public void SumIfGreaterThanZero()
		{
			_worksheet.Cells["A1"].Value = null;
			_worksheet.Cells["A2"].Value = string.Empty;
			_worksheet.Cells["A3"].Value = 1d;
			_worksheet.Cells["B1"].Value = 1d;
			_worksheet.Cells["B2"].Value = 3d;
			_worksheet.Cells["B3"].Value = 5d;
			var func = new SumIf();
			IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
			IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
			var args = FunctionsHelper.CreateArgs(range1, ">0", range2);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(5d, result.Result);
		}

		[TestMethod]
		public void SumIfGreaterThanOrEqualToZero()
		{
			_worksheet.Cells["A1"].Value = null;
			_worksheet.Cells["A2"].Value = string.Empty;
			_worksheet.Cells["A3"].Value = 1d;
			_worksheet.Cells["B1"].Value = 1d;
			_worksheet.Cells["B2"].Value = 3d;
			_worksheet.Cells["B3"].Value = 5d;
			var func = new SumIf();
			IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
			IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
			var args = FunctionsHelper.CreateArgs(range1, ">=0", range2);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(5d, result.Result);
		}

		[TestMethod]
		public void SumIfLessThanZero()
		{
			_worksheet.Cells["A1"].Value = null;
			_worksheet.Cells["A2"].Value = string.Empty;
			_worksheet.Cells["A3"].Value = -1d;
			_worksheet.Cells["B1"].Value = 1d;
			_worksheet.Cells["B2"].Value = 3d;
			_worksheet.Cells["B3"].Value = 5d;
			var func = new SumIf();
			IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
			IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
			var args = FunctionsHelper.CreateArgs(range1, "<0", range2);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(5d, result.Result);
		}

		[TestMethod]
		public void SumIfLessThanOrEqualToZero()
		{
			_worksheet.Cells["A1"].Value = null;
			_worksheet.Cells["A2"].Value = string.Empty;
			_worksheet.Cells["A3"].Value = -1d;
			_worksheet.Cells["B1"].Value = 1d;
			_worksheet.Cells["B2"].Value = 3d;
			_worksheet.Cells["B3"].Value = 5d;
			var func = new SumIf();
			IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
			IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
			var args = FunctionsHelper.CreateArgs(range1, "<=0", range2);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(5d, result.Result);
		}

		[TestMethod]
		public void SumIfLessThanCharacter()
		{
			_worksheet.Cells["A1"].Value = null;
			_worksheet.Cells["A2"].Value = string.Empty;
			_worksheet.Cells["A3"].Value = "Not Empty";
			_worksheet.Cells["B1"].Value = 1d;
			_worksheet.Cells["B2"].Value = 3d;
			_worksheet.Cells["B3"].Value = 5d;
			var func = new SumIf();
			IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
			IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
			var args = FunctionsHelper.CreateArgs(range1, "<a", range2);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(3d, result.Result);
		}

		[TestMethod]
		public void SumIfLessThanOrEqualToCharacter()
		{
			_worksheet.Cells["A1"].Value = null;
			_worksheet.Cells["A2"].Value = string.Empty;
			_worksheet.Cells["A3"].Value = "Not Empty";
			_worksheet.Cells["B1"].Value = 1d;
			_worksheet.Cells["B2"].Value = 3d;
			_worksheet.Cells["B3"].Value = 5d;
			var func = new SumIf();
			IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
			IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
			var args = FunctionsHelper.CreateArgs(range1, "<=a", range2);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(3d, result.Result);
		}

		[TestMethod]
		public void SumIfGreaterThanCharacter()
		{
			_worksheet.Cells["A1"].Value = null;
			_worksheet.Cells["A2"].Value = string.Empty;
			_worksheet.Cells["A3"].Value = "Not Empty";
			_worksheet.Cells["B1"].Value = 1d;
			_worksheet.Cells["B2"].Value = 3d;
			_worksheet.Cells["B3"].Value = 5d;
			var func = new SumIf();
			IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
			IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
			var args = FunctionsHelper.CreateArgs(range1, ">a", range2);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(5d, result.Result);
		}

		[TestMethod]
		public void SumIfGreaterThanOrEqualToCharacter()
		{
			_worksheet.Cells["A1"].Value = null;
			_worksheet.Cells["A2"].Value = string.Empty;
			_worksheet.Cells["A3"].Value = "Not Empty";
			_worksheet.Cells["B1"].Value = 1d;
			_worksheet.Cells["B2"].Value = 3d;
			_worksheet.Cells["B3"].Value = 5d;
			var func = new SumIf();
			IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
			IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
			var args = FunctionsHelper.CreateArgs(range1, ">=a", range2);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(5d, result.Result);
		}

		[TestMethod]
		public void SumIfHandleDates()
		{
			_worksheet.Cells["A1"].Value = null;
			_worksheet.Cells["A2"].Value = string.Empty;
			_worksheet.Cells["A3"].Value = "Not Empty";
			_worksheet.Cells["B1"].Value = 1d;
			_worksheet.Cells["B2"].Value = 3d;
			_worksheet.Cells["B3"].Value = 5d;
			var func = new SumIf();
			IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
			IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
			var args = FunctionsHelper.CreateArgs(range1, ">=a", range2);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(5d, result.Result);
		}

		[TestMethod]
		public void SumIfShouldHandleBooleanArg()
		{
			using (var pck = new ExcelPackage())
			{
				var sheet = pck.Workbook.Worksheets.Add("test");
				sheet.Cells["A1"].Value = true;
				sheet.Cells["B1"].Value = 1;
				sheet.Cells["A2"].Value = false;
				sheet.Cells["B2"].Value = 1;
				sheet.Cells["C1"].Formula = "SUMIF(A1:A2,TRUE,B1:B2)";
				sheet.Calculate();
				Assert.AreEqual(1d, sheet.Cells["C1"].Value);
			}
		}

		[TestMethod]
		public void SumIfDateComparison()
		{
			_worksheet.Cells[2, 3].Value = new DateTime(2012, 1, 1);
			_worksheet.Cells[3, 3].Value = new DateTime(2012, 6, 1);
			_worksheet.Cells[4, 3].Value = new DateTime(2012, 12, 1);
			_worksheet.Cells[5, 3].Value = new DateTime(2014, 1, 1);
			_worksheet.Cells[6, 3].Value = new DateTime(2014, 6, 1);
			_worksheet.Cells[2, 4].Value = 1.0;
			_worksheet.Cells[3, 4].Value = 1.0;
			_worksheet.Cells[4, 4].Value = 1.0;
			_worksheet.Cells[5, 4].Value = 1.0;
			_worksheet.Cells[6, 4].Value = 1.0;
			_worksheet.Cells[8, 2].Value = new DateTime(2013, 1, 1);
			_worksheet.Cells[8, 3].Formula = "SUMIF(C2:C6,\"<\"&B8,D2:D6)";
			_worksheet.Calculate();
			Assert.AreEqual(3.0, _worksheet.Cells[8, 3].Value);
		}

		[TestMethod]
		public void SumIfSingleCellWithNoSumRange()
		{
			_worksheet.Cells[2, 2].Value = 1;
			_worksheet.Cells[3, 3].Formula = "SUMIF(B2,1)";
			_worksheet.Cells[3, 3].Calculate();
			Assert.AreEqual(1d, _worksheet.Cells[3, 3].Value);
		}

		[TestMethod]
		public void SumIfSingleCellWithSumRange()
		{
			_worksheet.Cells[2, 2].Value = "Value";
			_worksheet.Cells[2, 3].Value = 1;
			_worksheet.Cells[3, 3].Formula = "SUMIF(B2,\"Value\",C2)";
			_worksheet.Cells[3, 3].Calculate();
			Assert.AreEqual(1d, _worksheet.Cells[3, 3].Value);
		}

		[TestMethod]
		public void SumIfArrayComparisons()
		{
			_worksheet.Cells[2, 2].Value = 1;
			_worksheet.Cells[2, 3].Formula = "{1,2,3;4,5,6}";
			_worksheet.Cells[3, 3].Formula = "SUMIF(C2,{1},B2)";
			_worksheet.Cells[2, 4].Formula = "{1}";
			_worksheet.Cells[3, 4].Formula = "SUMIF(D2,{1,2,3},B2)";
			_worksheet.Cells[2, 5].Formula = "{1,2,3}";
			_worksheet.Cells[3, 5].Formula = "SUMIF(E2,\"{1,2,3}\",B2)";
			_worksheet.Calculate();
			Assert.AreEqual(1d, _worksheet.Cells[3, 3].Value);
			Assert.AreEqual(1d, _worksheet.Cells[3, 4].Value);
			Assert.AreEqual(0d, _worksheet.Cells[3, 5].Value);
		}

		[TestMethod]
		public void SumIfWithArraySingleCell()
		{
			_worksheet.Cells[2, 2].Value = 1;
			_worksheet.Cells[2, 3].Formula = "{1,2,3}";
			_worksheet.Cells[3, 3].Formula = "SUMIF(C2,{1,2,3},B2)";
			_worksheet.Cells[3, 3].Calculate();
			Assert.AreEqual(1d, _worksheet.Cells[3, 3].Value);
		}

		[TestMethod]
		public void SumIfWithArrayMultiCell()
		{
			_worksheet.Cells[2, 2].Value = 1;
			_worksheet.Cells[2, 3].Value = 1;
			_worksheet.Cells[2, 4].Value = 1;
			_worksheet.Cells[3, 2].Formula = "{1,2,3}";
			_worksheet.Cells[3, 3].Formula = "{1,2,3}";
			_worksheet.Cells[3, 4].Formula = "{1,2,3}";
			_worksheet.Cells[4, 4].Formula = "SUMIF(B3:D3,{1,2,3},B2:D2)";
			_worksheet.Cells[4, 4].Calculate();
			Assert.AreEqual(3d, _worksheet.Cells[4, 4].Value);
		}

		[TestMethod]
		public void SumIfWithErrorSingleCell()
		{
			_worksheet.Cells[2, 2].Value = "Value";
			_worksheet.Cells[3, 2].Value = ExcelErrorValue.Create(eErrorType.Value);
			_worksheet.Cells[4, 4].Formula = "SUMIF(B3,\"Value\")";
			_worksheet.Cells[5, 4].Formula = "SUMIF(B2,\"Value\",B3)";
			_worksheet.Calculate();
			Assert.AreEqual(0d, _worksheet.Cells[4, 4].Value);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)_worksheet.Cells[5, 4].Value).Type);
		}

		[TestMethod]
		public void SumIfWithErrorMultiCell()
		{
			_worksheet.Cells[2, 2].Value = "Value";
			_worksheet.Cells[2, 3].Value = "Value";
			_worksheet.Cells[2, 4].Value = "Value";
			_worksheet.Cells[3, 2].Value = ExcelErrorValue.Create(eErrorType.Value);
			_worksheet.Cells[3, 3].Value = ExcelErrorValue.Create(eErrorType.Value);
			_worksheet.Cells[3, 4].Value = ExcelErrorValue.Create(eErrorType.Value);
			_worksheet.Cells[4, 4].Formula = "SUMIF(B3:D3,\"Value\")";
			_worksheet.Cells[5, 4].Formula = "SUMIF(B2:D2,\"Value\",B3:D3)";
			_worksheet.Calculate();
			Assert.AreEqual(0d, _worksheet.Cells[4, 4].Value);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)_worksheet.Cells[5, 4].Value).Type);
		}

		[TestMethod]
		public void SumIfWithRangeWithOnlyNumbersReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 4;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Formula = "SUMIF(B1:B2, \"<>-10\")";
				worksheet.Calculate();
				Assert.AreEqual(6d, worksheet.Cells["B3"].Value);
			}
		}

		[TestMethod]
		public void SumIfWithRangeWithNumericStringReturnsCorrectValue()
		{
			using (var packge = new ExcelPackage())
			{
				var worksheet = packge.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "4";
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Formula = "SUMIF(B1:B2, \"<>-10\")";
				worksheet.Calculate();
				Assert.AreEqual(2d, worksheet.Cells["B3"].Value);
			}
		}

		[TestMethod]
		public void SumIfWithRangeWithNonNumericStringReturnsCorrctValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "word";
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Formula = "SUMIF(B1:B2, \"<>-10\")";
				worksheet.Calculate();
				Assert.AreEqual(2d, worksheet.Cells["B3"].Value);
			}
		}

		[TestMethod]
		public void SumIfWithRangeWithBooleanValueReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = true;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Formula = "SUMIF(B1:B2, \"<>-10\")";
				worksheet.Calculate();
				Assert.AreEqual(2d, worksheet.Cells["B3"].Value);
			}
		}

		[TestMethod]
		public void SumIfWithDateReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "DATE(2017, 6, 22)";
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Formula = "SUMIF(B1:B2, \"<>-10\")";
				worksheet.Calculate();
				Assert.AreEqual(42910d, worksheet.Cells["B3"].Value);
			}
		}

		[TestMethod]
		public void SumIfWithDateInStringReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "\"6/22/2017\"";
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Formula = "SUMIF(B1:B2, \"<>-10\")";
				worksheet.Calculate();
				Assert.AreEqual(2d, worksheet.Cells["B3"].Value);
			}
		}

		[TestMethod]
		public void SumIfWithEmptyCellReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 2;
				worksheet.Cells["B3"].Formula = "SUMIF(B1:B2, \"<>-10\")";
				worksheet.Calculate();
				Assert.AreEqual(2d, worksheet.Cells["B3"].Value);
			}
		}

		[TestMethod]
		public void SumIfWithErrorValueCellReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 2;
				worksheet.Cells["C1"].Formula = "notaformula";
				worksheet.Cells["B3"].Formula = "SUMIF(B1:C1, \"<>-10\")";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Name, ((ExcelErrorValue)worksheet.Cells["B3"].Value).Type);
			}
		}

		[TestMethod]
		public void SumIfWithRangeWithOnlyNumbersCriteriaReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 4;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 4;
				worksheet.Cells["B4"].Formula = "SUMIF(B2:B3, \"<>-10\", $B$1)";
				worksheet.Calculate();
				Assert.AreEqual(6d, worksheet.Cells["B4"].Value);
			}
		}

		[TestMethod]
		public void SumIfWithNumericStringAndNumericCriteriaReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 4;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = "\"4\"";
				worksheet.Cells["B4"].Value = 2;
				worksheet.Cells["B5"].Formula = "SUMIF(B3:B4, \">0\", $B$1:$B$2)";
				worksheet.Calculate();
				Assert.AreEqual(2d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void SumIfWithNumericStringNonNumericCriteriaReturnsCorrectResult()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 4;
				worksheet.Cells["C1"].Value = 2;
				worksheet.Cells["B2"].Value = "\"4\"";
				worksheet.Cells["C2"].Value = 2;
				worksheet.Cells["B4"].Formula = "SUMIF(B2:C2, \"<>-10\", B1:C1)";
				worksheet.Calculate();
				Assert.AreEqual(6d, worksheet.Cells["B4"].Value);
			}
		}

		[TestMethod]
		public void SumIfWithRangeWithNonNumericStringAndTextCriteriaReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 4;
				worksheet.Cells["B2"].Value = "word";
				worksheet.Cells["B3"].Value = 2;
				worksheet.Cells["B4"].Formula = "SUMIF(B2:B3, \"word\", $B$1)";
				worksheet.Calculate();
				Assert.AreEqual(4d, worksheet.Cells["B4"].Value);
			}
		}

		[TestMethod]
		public void SumIfWithRangeWithNonNumericStringAndNonNumericCriteritaReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 4;
				worksheet.Cells["C1"].Value = 2;
				worksheet.Cells["B4"].Value = "word";
				worksheet.Cells["C4"].Value = 2;
				worksheet.Cells["B5"].Formula = "SUMIF(B4:C4,\"<>-10\",B1:C1)";
				worksheet.Calculate();
				Assert.AreEqual(6d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void SumIfWithRangeWithBooleanValueAndBooleanCriteriaReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = true;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 4;
				worksheet.Cells["B4"].Formula = "SUMIF(B1:B2, TRUE, $B$3)";
				worksheet.Calculate();
				Assert.AreEqual(4d, worksheet.Cells["B4"].Value);
			}
		}

		[TestMethod]
		public void SumIfWithRangeWithBooleanValueAndNonNumericCriteriaReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["A2"].Value = false;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 4;
				worksheet.Cells["C3"].Value = 2;
				worksheet.Cells["B4"].Formula = "SUMIF(A2:B2, \"<>-10\", B3:C3)";
				worksheet.Calculate();
				Assert.AreEqual(6d, worksheet.Cells["B4"].Value);
			}
		}

		[TestMethod]
		public void SumIfWithRangeWithDateAndDateCriteriaReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 2;
				worksheet.Cells["C1"].Formula = "DATE(2017, 6, 22)";
				worksheet.Cells["B3"].Value = 4;
				worksheet.Cells["C3"].Value = 2;
				worksheet.Cells["B4"].Formula = "SUMIF(B1:C1, \"<6/23/2017\", B3:C3)";
				worksheet.Calculate();
				Assert.AreEqual(6d, worksheet.Cells["B4"].Value);
			}
		}

		[TestMethod]
		public void SumIfWihtDateInStringReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 2;
				worksheet.Cells["B2"].Value = "6/22/2017";
				worksheet.Cells["B3"].Value = 4;
				worksheet.Cells["B4"].Formula = "SUMIF(B1:B2, \"6/22/2017\", $B$3)";
			}
		}

		[TestMethod]
		public void SumIfWithRangeWithEmptyCellAndEmptyStringCriteriaReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Shet1");
				worksheet.Cells["B1"].Value = 4;
				worksheet.Cells["C1"].Value = 2;
				worksheet.Cells["A3"].Value = null;
				worksheet.Cells["B3"].Value = 2;
				worksheet.Cells["B4"].Formula = "SUMIF(A3:B3, \"\", B1:C1)";
				worksheet.Calculate();
				Assert.AreEqual(4d, worksheet.Cells["B4"].Value);
			}
		}

		[TestMethod]
		public void SumIfWithRangeWithEmptyCellAndNonNumericCriteriaReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Shet1");
				worksheet.Cells["B1"].Value = 4;
				worksheet.Cells["C1"].Value = 2;
				worksheet.Cells["A3"].Value = null;
				worksheet.Cells["B3"].Value = 2;
				worksheet.Cells["B4"].Formula = "SUMIF(A3:B3, \"<>-10\", B1:C1)";
				worksheet.Calculate();
				Assert.AreEqual(6d, worksheet.Cells["B4"].Value);
			}
		}

		[TestMethod]
		public void SumIfWithRangeWithErrorValueReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 4;
				worksheet.Cells["B2"].Formula = "notaformula";
				worksheet.Cells["B3"].Value = 2;
				worksheet.Cells["B4"].Formula = "SUMIF(B1:B3, \"<>-10\", $B$1)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Name, ((ExcelErrorValue)worksheet.Cells["B4"].Value).Type);
			}
		}

		[TestMethod]
		public void SumIfCriteriaTestsReturnTheCorrectValue()
		{
			//These tests all involve various inputs to the criteria portion of the function.
			//they involve a large set of data so they have been grouped into this one unit test.
			var package = this.CreateTestingPackage();
			var worksheet = package.Workbook.Worksheets["Sheet1"];

			worksheet.Cells["B2"].Formula = "SUMIF(C2:W2,5,C3:W3)";
			worksheet.Cells["B3"].Formula = "SUMIF(C2:W2,\"5\",C3:W3)";
			worksheet.Cells["B4"].Formula = "SUMIF(C2:W2,\"=5\",C3:W3)";
			worksheet.Cells["B5"].Formula = "SUMIF(C2:W2,3.5,C3:W3)";
			worksheet.Cells["B6"].Formula = "SUMIF(C2:W2,TRUE,C3:W3)";
			worksheet.Cells["B7"].Formula = "SUMIF(C2:W2,\"TRUE\",C3:W3)";
			worksheet.Cells["B8"].Formula = "SUMIF(C2:W2,\"6/23/2017\",C3:W3)";
			worksheet.Cells["B9"].Formula = "SUMIF(C2:W2,\"6:00 PM\",C3:W3)";
			worksheet.Cells["B10"].Formula = "SUMIF(C2:W2,\"T*sday\",C3:W3)";
			worksheet.Cells["B11"].Formula = "SUMIF(C2:W2,\">=1\",C3:W3)";
			worksheet.Cells["B12"].Formula = "SUMIF(C2:W2,\">6/22/2017\",C3:W3)";
			worksheet.Cells["B13"].Formula = "SUMIF(C11:E11,1,F11:H11)";
			worksheet.Calculate();
			Assert.AreEqual(17d, worksheet.Cells["B2"].Value);
			Assert.AreEqual(17d, worksheet.Cells["B3"].Value);
			Assert.AreEqual(17d, worksheet.Cells["B4"].Value);
			Assert.AreEqual(8d, worksheet.Cells["B5"].Value);
			Assert.AreEqual(12d, worksheet.Cells["B6"].Value);
			Assert.AreEqual(12d, worksheet.Cells["B7"].Value);
			Assert.AreEqual(16d, worksheet.Cells["B8"].Value);
			Assert.AreEqual(21d, worksheet.Cells["B9"].Value);
			Assert.AreEqual(12d, worksheet.Cells["B10"].Value);
			Assert.AreEqual(88d, worksheet.Cells["B11"].Value);
			Assert.AreEqual(33d, worksheet.Cells["B12"].Value);
			Assert.AreEqual(4d, worksheet.Cells["B13"].Value);
		}

		[TestMethod]
		public void SumIfAverageRangeWithNumbersReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 3;
				worksheet.Cells["B2"].Value = 1;
				worksheet.Cells["B3"].Value = 4;
				worksheet.Cells["B4"].Value = 2;
				worksheet.Cells["B5"].Formula = "SUMIF($B$1:$B$2, \">0\", $B$3:$B$4)";
				worksheet.Calculate();
				Assert.AreEqual(6d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void SumIfAverageRangeWithNumericStringReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 3;
				worksheet.Cells["B2"].Value = 1;
				worksheet.Cells["B3"].Value = 2;
				worksheet.Cells["B4"].Value = "4";
				worksheet.Cells["B5"].Formula = "SUMIF($B$1:$B$2, \">0\", $B$3:$B$4)";
				worksheet.Calculate();
				Assert.AreEqual(2d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void SumIfAverageRangeWithNonNumericStringsReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 3;
				worksheet.Cells["B2"].Value = 1;
				worksheet.Cells["B3"].Value = "word";
				worksheet.Cells["B4"].Value = 2;
				worksheet.Cells["B5"].Formula = "SUMIF($B$1:$B$2, \">0\", $B$3:$B$4)";
				worksheet.Calculate();
				Assert.AreEqual(2d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void SumIfAverageRangeWithBoolenReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 3;
				worksheet.Cells["B2"].Value = 1;
				worksheet.Cells["B3"].Value = true;
				worksheet.Cells["B4"].Value = 2;
				worksheet.Cells["B5"].Formula = "SUMIF($B$1:$B$2, \">0\", $B$3:$B$4)";
				worksheet.Calculate();
				Assert.AreEqual(2d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void SumIfAverageRangeWithDateReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 3;
				worksheet.Cells["B2"].Value = 1;
				worksheet.Cells["B3"].Formula = "DATE(2017, 6, 22)";
				worksheet.Cells["B4"].Value = 2;
				worksheet.Cells["B5"].Formula = "SUMIF($B$1:$B$2, \">0\", $B$3:$B$4)";
				worksheet.Calculate();
				Assert.AreEqual(42910d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void SumIfAverageRangeWithDateInStringReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 3;
				worksheet.Cells["B2"].Value = 1;
				worksheet.Cells["B3"].Value = "6/22/2017";
				worksheet.Cells["B4"].Value = 2;
				worksheet.Cells["B5"].Formula = "SUMIF($B$1:$B$2, \">0\", $B$3:$B$4)";
				worksheet.Calculate();
				Assert.AreEqual(2d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void SumIfAverageRangeWithEmptyCellReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 3;
				worksheet.Cells["B2"].Value = 1;
				worksheet.Cells["B4"].Value = 2;
				worksheet.Cells["B5"].Formula = "SUMIF($B$1:$B$2, \">0\", $B$3:$B$4)";
				worksheet.Calculate();
				Assert.AreEqual(2d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void SumIfAverageRangeWithErrorValueReturnsTheErrorValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["C2"].Value = 3;
				worksheet.Cells["D2"].Value = 1;
				worksheet.Cells["B10"].Formula = "AVERAGEIF(C2:D2,\">0\",C10:D10)";
				worksheet.Cells["C10"].Formula = "notAValidFormula"; // Evaluates to #NAME.
				worksheet.Cells["D10"].Value = 2;
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Name, ((ExcelErrorValue)worksheet.Cells["B10"].Value).Type);
			}
		}

		[TestMethod]
		public void SumIfAverageRangeWithLowerCaseCriteriaReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "aa";
				worksheet.Cells["B2"].Value = "ab";
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["B4"].Value = 1;
				worksheet.Cells["B5"].Formula = "SUMIF(B1:B2, \"?b\", B3:B4)";
				worksheet.Calculate();
				Assert.AreEqual(1d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void SumIfAverageRangeWithUpperCaseCriteriaReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "aa";
				worksheet.Cells["B2"].Value = "ab";
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["B4"].Value = 1;
				worksheet.Cells["B5"].Formula = "SUMIF(B1:B2, \"*B\", B3:B4)";
				worksheet.Calculate();
				Assert.AreEqual(1d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void SumIfCriteriaAsSingeRangeWithTwoArgumentsReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 3;
				worksheet.Cells["B2"].Value = 1;
				worksheet.Cells["B3"].Value = ">0";
				worksheet.Cells["B4"].Formula = "SUMIF(B1:B2, B3)";
				worksheet.Calculate();
				Assert.AreEqual(4d, worksheet.Cells["B4"].Value);
			}
		}

		[TestMethod]
		public void SumIfCriteriaAsSingleCellRangeWithThreeArgumentsReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "aa";
				worksheet.Cells["B2"].Value = "ab";
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["B4"].Value = 1;
				worksheet.Cells["B5"].Value = "*b";
				worksheet.Cells["B6"].Formula = "SUMIF(B1:B2, B5, B3:B4)";
				worksheet.Calculate();
				Assert.AreEqual(1d, worksheet.Cells["B6"].Value);
			}
		}

		[TestMethod]
		public void SumIfCriteriaAsMultiCellRangeWithTwoArgumentsReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 3;
				worksheet.Cells["B2"].Value = 1;
				worksheet.Cells["B3"].Formula = "SUMIF(B1:B2, B1:B2)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B3"].Value);
			}
		}

		[TestMethod]
		public void SumIfCriteriaAsMultiCellRangeWithThreeArgumentsReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 3;
				worksheet.Cells["B2"].Value = 1;
				worksheet.Cells["B3"].Value = "aa";
				worksheet.Cells["B4"].Value = "ab";
				worksheet.Cells["B5"].Formula = "SUMIF(B3:B4, B3:B4, B1:B2)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void SumIfCriteriaAsMultiCellRangeWithFunctionCallInDifferentLocationsReturnsCorrectValues()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["C1"].Value = 1;
				worksheet.Cells["C2"].Value = 2;
				worksheet.Cells["C3"].Value = 3;
				worksheet.Cells["D1"].Formula = "SUMIF(B1:B3, C1:C3)";
				worksheet.Cells["D2"].Formula = "SUMIF(B1:B3, C1:C3)";
				worksheet.Cells["D3"].Formula = "SUMIF(B1:B3, C1:C3)";
				worksheet.Cells["C4"].Formula = "SUMIF(B1:B3, C1:C3)";
				worksheet.Cells["B8"].Value = 1;
				worksheet.Cells["C8"].Value = 2;
				worksheet.Cells["D8"].Value = 3;
				worksheet.Cells["B9"].Value = 1;
				worksheet.Cells["C9"].Value = 2;
				worksheet.Cells["D9"].Value = 3;
				worksheet.Cells["E9"].Formula = "SUMIF(B8:D8, B9:D9)";
				worksheet.Cells["B10"].Formula = "SUMIF(B8:D8, B9:D9)";
				worksheet.Cells["C10"].Formula = "SUMIF(B8:D8, B9:D9)";
				worksheet.Cells["D10"].Formula = "SUMIF(B8:D8, B9:D9)";
				worksheet.Calculate();
				Assert.AreEqual(1d, worksheet.Cells["D1"].Value);
				Assert.AreEqual(2d, worksheet.Cells["D2"].Value);
				Assert.AreEqual(3d, worksheet.Cells["D3"].Value);
				Assert.AreEqual(0d, worksheet.Cells["C4"].Value);
				Assert.AreEqual(0d, worksheet.Cells["E9"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B10"].Value);
				Assert.AreEqual(2d, worksheet.Cells["C10"].Value);
				Assert.AreEqual(3d, worksheet.Cells["D10"].Value);
			}
		}

		[TestMethod]
		public void SumIfWithCriteriaAsArrayReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 1;
				worksheet.Cells["B3"].Value = 1;
				worksheet.Cells["B4"].Value = 1;
				worksheet.Cells["B5"].Value = 1;
				worksheet.Cells["B6"].Value = 1;
				worksheet.Cells["B7"].Formula = "SUMIF(B1:B3, {1,2,3}, B4:B6)";
				worksheet.Calculate();
				Assert.AreEqual(3d, worksheet.Cells["B7"].Value);
			}
		}

		[TestMethod]
		public void SumIfWithCriteraAsCellReferenceToAnArrayReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 1;
				worksheet.Cells["B3"].Value = 1;
				worksheet.Cells["B4"].Formula = "{1,2,3}";
				worksheet.Cells["B5"].Formula = "{1,2,3}";
				worksheet.Cells["B6"].Formula = "{1,2,3}";
				worksheet.Cells["B7"].Formula = "SUMIF(B4:B6, B4, B1:B3)";
				worksheet.Calculate();
				Assert.AreEqual(3d, worksheet.Cells["B7"].Value);
			}
		}

		[TestMethod]
		public void SumIfWithLessThanACriteriaReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = null;
				worksheet.Cells["B2"].Value = string.Empty;
				worksheet.Cells["B3"].Value = "notempty";
				worksheet.Cells["C1"].Value = 1;
				worksheet.Cells["C2"].Value = 3;
				worksheet.Cells["C3"].Value = 5;
				worksheet.Cells["D1"].Formula = "SUMIFS(C1:C3, B1:B3, \"<a\")";
				worksheet.Calculate();
				Assert.AreEqual(3d, worksheet.Cells["D1"].Value);
			}
		}

		[TestMethod]
		public void SumIfWithEmptyStringCriteriaReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = null;
				worksheet.Cells["B2"].Value = string.Empty;
				worksheet.Cells["B3"].Value = "notempty";
				worksheet.Cells["C1"].Value = 1;
				worksheet.Cells["C2"].Value = 3;
				worksheet.Cells["C3"].Value = 5;
				worksheet.Cells["D2"].Formula = "SUMIFS(C1:C3, B1:B3, \"\")";
				worksheet.Calculate();
				Assert.AreEqual(4d, worksheet.Cells["D2"].Value);
			}
		}

		[TestMethod]
		public void SumIfWithBooleanValueInputsAsCriteriaReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = true;
				worksheet.Cells["B2"].Value = 1;
				worksheet.Cells["B3"].Value = false;
				worksheet.Cells["B4"].Formula = "SUMIF(B1, \">TRUE\", B2)";
				worksheet.Cells["B5"].Formula = "SUMIF(B3, \"<TRUE\", B2)";
				worksheet.Cells["B6"].Formula = "SUMIF(B1, \">FALSE\", B2)";
				worksheet.Cells["B7"].Formula = "SUMIF(B3, \"<FALSE\", B2)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B4"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B5"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B6"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B7"].Value);
			}
		}

		[TestMethod]
		public void SumIfWithCriteriaWithExpressionCharacterReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheeet = package.Workbook.Worksheets.Add("Sheet1");
				worksheeet.Cells["B1"].Value = "=";
				worksheeet.Cells["B2"].Value = 1;
				worksheeet.Cells["B3"].Formula = "SUMIF(B1, \"=\", B2)";
				worksheeet.Calculate();
				Assert.AreEqual(0d, worksheeet.Cells["B3"].Value);
			}
		}

		[TestMethod]
		public void SumIfWithExpressionCharAndEmptyStringReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "";
				worksheet.Cells["B2"].Value = 1;
				worksheet.Cells["B3"].Formula = "SUMIF(B1, \"=\", B2)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B3"].Value);
			}
		}

		[TestMethod]
		public void SumIfWithExpressionCharAndEmptyCellReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = null;
				worksheet.Cells["B2"].Value = 1;
				worksheet.Cells["B3"].Formula = "SUMIF(B1, \"=\", B2)";
				worksheet.Calculate();
				Assert.AreEqual(1d, worksheet.Cells["B3"].Value);
			}
		}

		[TestMethod]
		public void SumIfWithStringComparisonsReturnsCorrectValues()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "ay";
				worksheet.Cells["B2"].Value = 1;
				worksheet.Cells["B3"].Formula = "SUMIF($B$1, \"<axz\", $B$2)";
				worksheet.Cells["B4"].Formula = "SUMIF($B$1, \"<aya\", $B$2)";
				worksheet.Cells["B5"].Formula = "SUMIF($B$1, \"<az\", $B$2)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B3"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B4"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void SumIfWithStringComparisonsWithWildcardCharacterReturnCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["C2"].Value = "ay";
				worksheet.Cells["C3"].Value = "Modday";
				worksheet.Cells["C4"].Value = "Monnnnday";
				worksheet.Cells["D2"].Value = 1;
				worksheet.Cells["D3"].Value = 3;
				worksheet.Cells["D4"].Value = 5;
				worksheet.Cells["B2"].Formula = "SUMIF(C2:C4,\"=Mo*day\",D2:D4)";
				worksheet.Cells["B3"].Formula = "SUMIF(C2:C4,\">Mo*day\",D2:D4)";
				worksheet.Cells["B4"].Formula = "SUMIF(C2:C4,\"<Mo*day\",D2:D4)";
				worksheet.Calculate();
				Assert.AreEqual(8d, worksheet.Cells["B2"].Value);
				Assert.AreEqual(8d, worksheet.Cells["B3"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B4"].Value);
			}
		}

		[TestMethod]
		public void SumIfWithStringComparisionsWithEscapedWildcardCharacterReturnCorrectValues()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "Mon?ay";
				worksheet.Cells["B2"].Value = "Monday";
				worksheet.Cells["B3"].Value = 1;
				worksheet.Cells["B4"].Value = 3;
				worksheet.Cells["C1"].Value = "Mon*ay";
				worksheet.Cells["C2"].Value = "Monddday";
				worksheet.Cells["C3"].Value = 5;
				worksheet.Cells["C4"].Value = 7;
				worksheet.Cells["D1"].Formula = "SUMIF(B1:B2, \"Mon?ay\", B3:B4)";
				worksheet.Cells["D2"].Formula = "SUMIF(B1:B2, \"Mon~?ay\", B3:B4)";
				worksheet.Cells["D3"].Formula = "SUMIF(C1:C2, \"Mon*ay\", C3:C4)";
				worksheet.Cells["D4"].Formula = "SUMIF(C1:C2, \"Mon~*ay\", C3:C4)";
				worksheet.Calculate();
				Assert.AreEqual(4d, worksheet.Cells["D1"].Value);
				Assert.AreEqual(1d, worksheet.Cells["D2"].Value);
				Assert.AreEqual(12d, worksheet.Cells["D3"].Value);
				Assert.AreEqual(5d, worksheet.Cells["D4"].Value);
			}
		}

		[TestMethod]
		public void SumIfWithStringComparisonWithEmptyCellReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B2"].Value = 1;
				worksheet.Cells["B3"].Formula = "SUMIF(B1, \">a\", B2)";
				worksheet.Cells["B4"].Formula = "SUMIF(B1, \"<a\", B2)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B3"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B4"].Value);
			}
		}

		[TestMethod]
		public void SumIfWithStringComparisonWithEmptyStringReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "";
				worksheet.Cells["B2"].Value = 1;
				worksheet.Cells["B3"].Formula = "SUMIF(B1, \">a\", B2)";
				worksheet.Cells["B4"].Formula = "SUMIF(B1, \"<a\", B2)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B3"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B4"].Value);
			}
		}

		[TestMethod]
		public void SumIfWithStringComparisonWithStringReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "zzz";
				worksheet.Cells["B2"].Value = 1;
				worksheet.Cells["B3"].Formula = "SUMIF(B1, \">a\", B2)";
				worksheet.Cells["B4"].Formula = "SUMIF(B1, \"<a\", B2)";
				worksheet.Calculate();
				Assert.AreEqual(1d, worksheet.Cells["B3"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B4"].Value);
			}
		}

		[TestMethod]
		public void SumIfWithStringComparisonWithNumberReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 1;
				worksheet.Cells["B3"].Formula = "SUMIF(B1, \">a\", B2)";
				worksheet.Cells["B4"].Formula = "SUMIF(B1, \"<a\", B2)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B3"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B4"].Value);
			}
		}

		[TestMethod]
		public void SumIfWithStringComparisonWithNumberInStringReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "1";
				worksheet.Cells["B2"].Value = 1;
				worksheet.Cells["B3"].Formula = "SUMIF(B1, \">a\", B2)";
				worksheet.Cells["B4"].Formula = "SUMIF(B1, \"<a\", B2)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B3"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B4"].Value);
			}
		}

		[TestMethod]
		public void SumIfWithStringComparisonWithBooleanValuesReturnsCorrectValues()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = true;
				worksheet.Cells["B2"].Value = 1;
				worksheet.Cells["B3"].Formula = "SUMIF(B1, \">a\", B2)";
				worksheet.Cells["B4"].Formula = "SUMIF(B1, \"<a\", B2)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B3"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B4"].Value);
			}

		}

		[TestMethod]
		public void SumIfWithRegexForExpressionCharactersReturnsCorrectValues()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 1;
				worksheet.Cells["B3"].Value = ">1";
				worksheet.Cells["B4"].Value = "<1";
				worksheet.Cells["B5"].Value = "=1";
				worksheet.Cells["C1"].Formula = "SUMIF(B1, \"=1\", B2)";
				worksheet.Cells["C2"].Formula = "SUMIF(B1, \">1\", B2)";
				worksheet.Cells["C3"].Formula = "SUMIF(B1, \"<1\", B2)";
				worksheet.Cells["C4"].Formula = "SUMIF(B1, \">=1\", B2)";
				worksheet.Cells["C5"].Formula = "SUMIF(B1, \"<=1\", B2)";
				worksheet.Cells["C6"].Formula = "SUMIF(B1, \"<>1\", B2)";
				worksheet.Cells["C7"].Formula = "SUMIF(B3, \"=>1\", B2)";
				worksheet.Cells["C8"].Formula = "SUMIF(B4, \"=<1\", B2)";
				worksheet.Cells["C9"].Formula = "SUMIF(B5, \"==1\", B2)";
				worksheet.Cells["C10"].Formula = "SUMIF(B1, \">>1\", B2)";
				worksheet.Cells["C11"].Formula = "SUMIF(B1, \"><1\", B2)";
				worksheet.Cells["C12"].Formula = "SUMIF(B1, \"<<1\", B2)";
				worksheet.Calculate();
				Assert.AreEqual(1d, worksheet.Cells["C1"].Value);
				Assert.AreEqual(0d, worksheet.Cells["C2"].Value);
				Assert.AreEqual(0d, worksheet.Cells["C3"].Value);
				Assert.AreEqual(1d, worksheet.Cells["C4"].Value);
				Assert.AreEqual(1d, worksheet.Cells["C5"].Value);
				Assert.AreEqual(0d, worksheet.Cells["C6"].Value);
				Assert.AreEqual(1d, worksheet.Cells["C7"].Value);
				Assert.AreEqual(1d, worksheet.Cells["C8"].Value);
				Assert.AreEqual(1d, worksheet.Cells["C9"].Value);
				Assert.AreEqual(0d, worksheet.Cells["C10"].Value);
				Assert.AreEqual(0d, worksheet.Cells["C11"].Value);
				Assert.AreEqual(0d, worksheet.Cells["C12"].Value);
			}
		}

		[TestMethod]
		public void SumIfWithNullCriteriaReturns0()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B2"].Formula = "SUMIF(C2:C3,1,D2:D3)";
				worksheet.Cells["B3"].Formula = "SUMIF(C4:C5,,D2:D3)";
				worksheet.Cells["B4"].Formula = "SUMIF(C4:C5,C5,D2:D3)";
				worksheet.Cells["C2"].Value = 1;
				worksheet.Cells["C3"].Value = 1;
				worksheet.Cells["C4"].Value = 0;
				worksheet.Cells["C5"].Value = null;
				worksheet.Cells["D2"].Value = 1.5;
				worksheet.Cells["D3"].Value = 2.5;
				worksheet.Calculate();
				Assert.AreEqual(4d, worksheet.Cells["B2"].Value);
				Assert.AreEqual(1.5, worksheet.Cells["B3"].Value);
				Assert.AreEqual(1.5, worksheet.Cells["B4"].Value);
			}
		}

		[TestMethod]
		public void SumIfWithUnsetEmptyCellsInCriteria()
		{
			// This test exists to ensure that cells that have never been set are still 
			// being compared against the criterion.
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B2"].Formula = "SUMIF(D2:D3,\"\",C2:C3)";
				worksheet.Cells["C2"].Value = 1;
				worksheet.Cells["C3"].Value = 2;
				worksheet.Calculate();
				Assert.AreEqual(3d, worksheet.Cells["B2"].Value);
			}
		}
		#endregion
	}
}
