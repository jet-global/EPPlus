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
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using static OfficeOpenXml.FormulaParsing.ExcelDataProvider;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math
{
	[TestClass]
	public class CountIfTests : MathFunctionsTestBase
	{
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

		#region CountIf Tests
		[TestMethod]
		public void CountIfWithVariedRangeValuesAndConstantCriteria()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B2"].Formula = "COUNTIF(C2:D2,\"<>-10\")";
				worksheet.Cells["C2"].Value = 4;
				worksheet.Cells["D2"].Value = 2;
				worksheet.Cells["B3"].Formula = "COUNTIF(C3:D3,\"<>-10\")";
				worksheet.Cells["C3"].Value = "4";
				worksheet.Cells["D3"].Value = 2;
				worksheet.Cells["B4"].Formula = "COUNTIF(C4:D4,\"<>-10\")";
				worksheet.Cells["C4"].Value = "word";
				worksheet.Cells["D4"].Value = 2;
				worksheet.Cells["B5"].Formula = "COUNTIF(C5:D5,\"<>-10\")";
				worksheet.Cells["C5"].Value = true;
				worksheet.Cells["D5"].Value = 2;
				worksheet.Cells["B6"].Formula = "COUNTIF(C6:D6,\"<>-10\")";
				worksheet.Cells["C6"].Value = (new System.DateTime(2017, 6, 22)).ToOADate();
				worksheet.Cells["D6"].Value = 2;
				worksheet.Cells["B7"].Formula = "COUNTIF(C7:D7,\"<>-10\")";
				worksheet.Cells["C7"].Value = "6/22/2017";
				worksheet.Cells["D7"].Value = 2;
				worksheet.Cells["B8"].Formula = "COUNTIF(C8:D8,\"<>-10\")";
				worksheet.Cells["C8"].Value = null;
				worksheet.Cells["D8"].Value = 2;
				worksheet.Cells["B9"].Formula = "COUNTIF(C9:D9,\"<>-10\")";
				worksheet.Cells["C9"].Formula = "notAValidFormula"; // Evaluates to #NAME.
				worksheet.Cells["D9"].Value = 2;
				worksheet.Calculate();
				Assert.AreEqual(2d, worksheet.Cells["B2"].Value);
				Assert.AreEqual(2d, worksheet.Cells["B3"].Value);
				Assert.AreEqual(2d, worksheet.Cells["B4"].Value);
				Assert.AreEqual(2d, worksheet.Cells["B5"].Value);
				Assert.AreEqual(2d, worksheet.Cells["B6"].Value);
				Assert.AreEqual(2d, worksheet.Cells["B7"].Value);
				Assert.AreEqual(2d, worksheet.Cells["B8"].Value);
				Assert.AreEqual(2d, worksheet.Cells["B9"].Value);
			}
		}
		
		[TestMethod]
		public void CountIfWithVariedCriteria()
		{
			using (var package = this.CreateTestingPackage())
			{
				var worksheet = package.Workbook.Worksheets["Sheet1"];
				worksheet.Cells["B2"].Formula = "COUNTIF(C2:W2,5)";
				worksheet.Cells["B3"].Formula = "COUNTIF(C2:W2,\"5\")";
				worksheet.Cells["B4"].Formula = "COUNTIF(C2:W2,\"=5\")";
				worksheet.Cells["B5"].Formula = "COUNTIF(C2:W2,3.5)";
				worksheet.Cells["B6"].Formula = "COUNTIF(C2:W2,TRUE)";
				worksheet.Cells["B7"].Formula = "COUNTIF(C2:W2,\"TRUE\")";
				worksheet.Cells["B8"].Formula = "COUNTIF(C2:W2,\"6/23/2017\")";
				worksheet.Cells["B9"].Formula = "COUNTIF(C2:W2,\"6:00 PM\")";
				worksheet.Cells["B10"].Formula = "COUNTIF(C2:W2,\"T*sday\")";
				worksheet.Cells["B11"].Formula = "COUNTIF(C2:W2,\">=1\")";
				worksheet.Cells["B12"].Formula = "COUNTIF(C2:W2,\">6/22/2017\")";
				worksheet.Cells["B13"].Formula = "COUNTIF(C11:E11,1)";
				worksheet.Calculate();
				Assert.AreEqual(2d, worksheet.Cells["B2"].Value);
				Assert.AreEqual(2d, worksheet.Cells["B3"].Value);
				Assert.AreEqual(2d, worksheet.Cells["B4"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B5"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B6"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B7"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B8"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B9"].Value);
				Assert.AreEqual(3d, worksheet.Cells["B10"].Value);
				Assert.AreEqual(8d, worksheet.Cells["B11"].Value);
				Assert.AreEqual(2d, worksheet.Cells["B12"].Value);
				Assert.AreEqual(2d, worksheet.Cells["B13"].Value);
			}
		}

		[TestMethod]
		public void CountIfCriteriaIgnoresCase()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["C2"].Value = "aa";
				worksheet.Cells["D2"].Value = "ab";
				worksheet.Cells["B2"].Formula = "COUNTIF(C2:D2,\"?b\")";
				worksheet.Cells["B3"].Formula = "COUNTIF(C2:D2,\"?B\")";
				worksheet.Calculate();
				Assert.AreEqual(1d, worksheet.Cells["B2"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B3"].Value);
			}
		}

		[TestMethod]
		public void CountIfWithCriteriaAsCellReference()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B2"].Formula = "COUNTIF(C3:D3,C2)";
				worksheet.Cells["C2"].Value = ">0";
				worksheet.Cells["C3"].Value = 3;
				worksheet.Cells["D3"].Value = 1;
				worksheet.Cells["B3"].Formula = "COUNTIF(C3:D3,C3:D3)";
				worksheet.Calculate();
				Assert.AreEqual(2d, worksheet.Cells["B2"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B3"].Value);
			}
		}

		[TestMethod]
		public void CountIfWithAbsoluteNamedRanges()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				package.Workbook.Names.Add("NamedRange", worksheet.Cells["$C$3:$D$3"]);
				worksheet.Cells["B2"].Formula = "COUNTIF(NamedRange,C2)";
				worksheet.Cells["C2"].Value = ">0";
				worksheet.Cells["C3"].Value = 3;
				worksheet.Cells["D3"].Value = 1;
				worksheet.Cells["B3"].Formula = "COUNTIF(NamedRange,NamedRange)";
				worksheet.Calculate();
				Assert.AreEqual(2d, worksheet.Cells["B2"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B3"].Value);
			}
		}

		[TestMethod]
		public void CountIfWithRelativeNamedRanges()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				package.Workbook.Names.Add("NamedRange", worksheet.Cells["C3:D3"]);
				worksheet.Cells["B2"].Formula = "COUNTIF(NamedRange,C2)";
				worksheet.Cells["C2"].Value = ">0";
				worksheet.Cells["D4"].Value = 3;
				worksheet.Cells["E4"].Value = 1;
				worksheet.Cells["B3"].Formula = "COUNTIF(NamedRange,NamedRange)";
				worksheet.Calculate();
				Assert.AreEqual(2d, worksheet.Cells["B2"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B3"].Value);
			}
		}

		[TestMethod]
		public void CountIfWithCriteriaAsArray()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["C2"].Formula = "{1,2,3}";
				worksheet.Cells["D2"].Formula = "{1,2,4}";
				worksheet.Cells["B2"].Formula = "COUNTIF(C2:D2,{1,2,3})";
				worksheet.Cells["B3"].Formula = "COUNTIF(C2:D2,C2)";
				worksheet.Calculate();
				Assert.AreEqual(2d, worksheet.Cells["B2"].Value);
				Assert.AreEqual(2d, worksheet.Cells["B3"].Value);
			}
		}

		[TestMethod]
		public void CountIfWithEmptyStringCriteria()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B2"].Formula = "COUNTIF(C2:C4,\"\")";
				worksheet.Cells["C2"].Value = null;
				worksheet.Cells["C3"].Value = string.Empty;
				worksheet.Cells["C4"].Value = "Not Empty";
				worksheet.Calculate();
				Assert.AreEqual(2d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void CountIfWithBooleanComparisons()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["C2"].Value = true;
				worksheet.Cells["C3"].Value = false;
				worksheet.Cells["B2"].Formula = "COUNTIF(C2,\">TRUE\")";
				worksheet.Cells["B3"].Formula = "COUNTIF(C3,\"<TRUE\")";
				worksheet.Cells["B4"].Formula = "COUNTIF(C2,\">FALSE\")";
				worksheet.Cells["B5"].Formula = "COUNTIF(C3,\"<FALSE\")";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B2"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B3"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B4"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void CountIfWithCriteriaAsExpressionCharacter()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B2"].Formula = "COUNTIF(C2,\"=\")";
				worksheet.Cells["C2"].Value = "=";
				worksheet.Cells["B3"].Formula = "COUNTIF(C3,\"=\")";
				worksheet.Cells["C3"].Value = string.Empty;
				worksheet.Cells["B4"].Formula = "COUNTIF(C4,\"=\")";
				worksheet.Cells["C4"].Value = null;
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B2"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B3"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B4"].Value);
			}
		}

		[TestMethod]
		public void CountIfWithStringValueComparisons()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["C2"].Value = "ay";
				worksheet.Cells["B2"].Formula = "COUNTIF(C2,\"<axz\")";
				worksheet.Cells["B3"].Formula = "COUNTIF(C2,\"<aya\")";
				worksheet.Cells["B4"].Formula = "COUNTIF(C2,\"<az\")";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B2"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B3"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B4"].Value);
			}
		}

		[TestMethod]
		public void CountIfWithStringComparisonsWithWildcardCharacter()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B2"].Formula = "COUNTIF(C2:C4,\"=Mo*day\")";
				worksheet.Cells["C2"].Value = "ay";
				worksheet.Cells["B3"].Formula = "COUNTIF(C2:C4,\">Mo*day\")";
				worksheet.Cells["C3"].Value = "Modday";
				worksheet.Cells["B4"].Formula = "COUNTIF(C2:C4,\"<Mo*day\")";
				worksheet.Cells["C4"].Value = "Monnnnday";
				worksheet.Calculate();
				Assert.AreEqual(2d, worksheet.Cells["B2"].Value);
				Assert.AreEqual(2d, worksheet.Cells["B3"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B4"].Value);
			}
		}

		[TestMethod]
		public void CountIfWithEscapedWildcardCriteria()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B2"].Formula = "COUNTIF(C2:C3,\"Mon?ay\")";
				worksheet.Cells["B3"].Formula = "COUNTIF(C2:C3,\"Mon~?ay\")";
				worksheet.Cells["C2"].Value = "Mon?ay";
				worksheet.Cells["C3"].Value = "Monday";
				worksheet.Cells["B4"].Formula = "COUNTIF(C4:C5,\"Mon*ay\")";
				worksheet.Cells["B5"].Formula = "COUNTIF(C4:C5,\"Mon~*ay\")";
				worksheet.Cells["C4"].Value = "Mon*ay";
				worksheet.Cells["C5"].Value = "Mondday";
				worksheet.Calculate();
				Assert.AreEqual(2d, worksheet.Cells["B2"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B3"].Value);
				Assert.AreEqual(2d, worksheet.Cells["B4"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void CountIfWithStringComparisons()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B2"].Formula = "COUNTIF(C2,\">a\")";
				worksheet.Cells["B3"].Formula = "COUNTIF(C2,\"<a\")";
				worksheet.Cells["C2"].Value = null;
				worksheet.Cells["B4"].Formula = "COUNTIF(C3,\">a\")";
				worksheet.Cells["B5"].Formula = "COUNTIF(C3,\"<a\")";
				worksheet.Cells["C3"].Value = string.Empty;
				worksheet.Cells["B6"].Formula = "COUNTIF(C4,\">a\")";
				worksheet.Cells["B7"].Formula = "COUNTIF(C4,\"<a\")";
				worksheet.Cells["C4"].Value = "zzz";
				worksheet.Cells["B8"].Formula = "COUNTIF(C5,\">a\")";
				worksheet.Cells["B9"].Formula = "COUNTIF(C5,\"<a\")";
				worksheet.Cells["C5"].Value = 1;
				worksheet.Cells["B10"].Formula = "COUNTIF(C6,\">a\")";
				worksheet.Cells["B11"].Formula = "COUNTIF(C6,\"<a\")";
				worksheet.Cells["C6"].Value = "1";
				worksheet.Cells["B12"].Formula = "COUNTIF(C7,\">a\")";
				worksheet.Cells["B13"].Formula = "COUNTIF(C7,\"<a\")";
				worksheet.Cells["C7"].Value = true;
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B2"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B3"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B4"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B5"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B6"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B7"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B8"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B9"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B10"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B11"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B12"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B13"].Value);
			}
		}

		[TestMethod]
		public void CountIfWithVariedExpressionCharacters()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["C2"].Value = 1;
				worksheet.Cells["C3"].Value = ">1";
				worksheet.Cells["C4"].Value = "<1";
				worksheet.Cells["C5"].Value = "=1";
				worksheet.Cells["B2"].Formula = "COUNTIF(C2,\"=1\")";
				worksheet.Cells["B3"].Formula = "COUNTIF(C2,\">1\")";
				worksheet.Cells["B4"].Formula = "COUNTIF(C2,\"<1\")";
				worksheet.Cells["B5"].Formula = "COUNTIF(C2,\">=1\")";
				worksheet.Cells["B6"].Formula = "COUNTIF(C2,\"<=1\")";
				worksheet.Cells["B7"].Formula = "COUNTIF(C2,\"<>1\")";
				worksheet.Cells["B8"].Formula = "COUNTIF(C3,\"=>1\")";
				worksheet.Cells["B9"].Formula = "COUNTIF(C4,\"=<1\")";
				worksheet.Cells["B10"].Formula = "COUNTIF(C5,\"==1\")";
				worksheet.Cells["B11"].Formula = "COUNTIF(C2,\">>1\")";
				worksheet.Cells["B12"].Formula = "COUNTIF(C2,\"><1\")";
				worksheet.Cells["B13"].Formula = "COUNTIF(C2,\"<<1\")";
				worksheet.Calculate();
				Assert.AreEqual(1d, worksheet.Cells["B2"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B3"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B4"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B5"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B6"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B7"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B8"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B9"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B10"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B11"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B12"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B13"].Value);
			}
		}

		[TestMethod]
		public void CountIfWithNullCriteriaReturns0()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B2"].Formula = "COUNTIF(C2:C5,1)";
				worksheet.Cells["B3"].Formula = "COUNTIF(C2:C5,)";
				worksheet.Cells["B4"].Formula = "COUNTIF(C2:C5,C5)";
				worksheet.Cells["C2"].Value = 1;
				worksheet.Cells["C3"].Value = 1;
				worksheet.Cells["C4"].Value = 0;
				worksheet.Cells["C5"].Value = null;
				worksheet.Calculate();
				Assert.AreEqual(2d, worksheet.Cells["B2"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B3"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B4"].Value);
			}
		}

		[TestMethod]
		public void CountIfWithNumbers()
		{
			_worksheet.Cells["A1"].Value = 1d;
			_worksheet.Cells["A2"].Value = 2d;
			_worksheet.Cells["A3"].Value = 3d;
			var func = new CountIf();
			IRangeInfo range = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
			var args = FunctionsHelper.CreateArgs(range, ">1");
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(2d, result.Result);
		}

		[TestMethod]
		public void CountIfWithNonNumericStrings()
		{
			_worksheet.Cells["A1"].Value = "Monday";
			_worksheet.Cells["A2"].Value = "Tuesday";
			_worksheet.Cells["A3"].Value = "Thursday";
			var func = new CountIf();
			IRangeInfo range = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
			var args = FunctionsHelper.CreateArgs(range, "T*day");
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(2d, result.Result);
		}

		[TestMethod]
		public void CountIfWithNumberAsCriteria()
		{
			_worksheet.Cells["A1"].Value = null;
			_worksheet.Cells["A2"].Value = 1d;
			_worksheet.Cells["A3"].Value = "Not Empty";
			var func = new CountIf();
			IRangeInfo range = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
			var args = FunctionsHelper.CreateArgs(range, 1d);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(1d, result.Result);
		}

		[TestMethod]
		public void CountIfWithCriteriaAsNotEqualToNull()
		{
			_worksheet.Cells["A1"].Value = null;
			_worksheet.Cells["A2"].Value = string.Empty;
			_worksheet.Cells["A3"].Value = "Not Empty";
			var func = new CountIf();
			IRangeInfo range = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
			var args = FunctionsHelper.CreateArgs(range, "<>");
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(2d, result.Result);
		}

		[TestMethod]
		public void CountIfWithCriteriaAsNotEqualToZero()
		{
			_worksheet.Cells["A1"].Value = null;
			_worksheet.Cells["A2"].Value = string.Empty;
			_worksheet.Cells["A3"].Value = 0d;
			var func = new CountIf();
			IRangeInfo range = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
			var args = FunctionsHelper.CreateArgs(range, "<>0");
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(2d, result.Result);
		}

		[TestMethod]
		public void CountIfWithArraySingleCell()
		{
			_worksheet.Cells[2, 3].Formula = "{1,2,3}";
			_worksheet.Cells[3, 3].Formula = "COUNTIF(C2,{1,2,3})";
			_worksheet.Cells[3, 3].Calculate();
			Assert.AreEqual(1d, _worksheet.Cells[3, 3].Value);
		}

		[TestMethod]
		public void CountIfWithArrayMultiCell()
		{
			_worksheet.Cells[2, 3].Formula = "{1,2,3}";
			_worksheet.Cells[2, 4].Formula = "{1,2,3}";
			_worksheet.Cells[2, 5].Formula = "{1,2,3}";
			_worksheet.Cells[3, 3].Formula = "COUNTIF(C2:E2,{1,2,3})";
			_worksheet.Cells[3, 3].Calculate();
			Assert.AreEqual(3d, _worksheet.Cells[3, 3].Value);
		}
		
		[TestMethod]
		public void CountIfWithInequalityDateCriterion()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B2"].Formula = "COUNTIF(C2:C3, \"7/1/2017\")";
				worksheet.Cells["B3"].Formula = "COUNTIF(C2:C3, \">=7/1/2017\")";
				worksheet.Cells["C2"].Value = (new System.DateTime(2017, 7, 1)).ToOADate();
				worksheet.Cells["C3"].Value = "7/1/2017";
				worksheet.Calculate();
				Assert.AreEqual(2d, worksheet.Cells["B2"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B3"].Value);
			}
		}

		[TestMethod]
		public void CountIfWithUnsetEmptyCellsInCriteria()
		{
			// This test exists to ensure that cells that have never been set are still 
			// being compared against the criterion.
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B2"].Formula = "COUNTIF(C2:C3,\"\")";
				worksheet.Calculate();
				Assert.AreEqual(2d, worksheet.Cells["B2"].Value);
			}
		}
		#endregion
	}
}
