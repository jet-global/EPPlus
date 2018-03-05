using System;
using EPPlusTest.FormulaParsing.Excel.Functions.Math;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest.FormulaParsing.Excel.Functions
{
	[TestClass]
	public class SumIfsTests : MathFunctionsTestBase
	{
		#region Calculate Tests
		[TestMethod]
		public void CalculateSingleCellRanges()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				sheet.Cells["A1"].Value = 1;

				sheet.Cells["B1"].Value = 1;

				sheet.Cells["E1"].Formula = "SUMIFS(A1:A4,B1:B4,1)";
				sheet.Calculate();

				Assert.AreEqual(1d, sheet.Cells["E1"].Value);
			}
		}

		[TestMethod]
		public void CalculateSingleCriteriaMatchedDimensionsMatchedRoot()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				sheet.Cells["A1"].Value = 1;
				sheet.Cells["A2"].Value = 2;
				sheet.Cells["A3"].Value = 3;
				sheet.Cells["A4"].Value = 4;

				sheet.Cells["B1"].Value = 1;
				sheet.Cells["B2"].Value = 2;
				sheet.Cells["B3"].Value = 3;
				sheet.Cells["B4"].Value = 4;

				sheet.Cells["E1"].Formula = "SUMIFS(A1:A4,B1:B4,\">2\")";
				sheet.Calculate();

				Assert.AreEqual(7d, sheet.Cells["E1"].Value);
			}
		}

		[TestMethod]
		public void CalculateSingleCriteriaMatchedDimensionsMatchedRootCriteriaReference()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				sheet.Cells["A1"].Value = 1;
				sheet.Cells["A2"].Value = 2;
				sheet.Cells["A3"].Value = 3;
				sheet.Cells["A4"].Value = 4;

				sheet.Cells["B1"].Value = 1;
				sheet.Cells["B2"].Value = 2;
				sheet.Cells["B3"].Value = 3;
				sheet.Cells["B4"].Value = 4;

				sheet.Cells["C1"].Value = ">2";

				sheet.Cells["E1"].Formula = "SUMIFS(A1:A4,B1:B4,C1)";
				sheet.Calculate();

				Assert.AreEqual(7d, sheet.Cells["E1"].Value);
			}
		}

		[TestMethod]
		public void CalculateSingleCriteriaMatchedDimensionsMatchedRootCriteriaMultiRangeReferenceTakesFirstCell()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				sheet.Cells["A1"].Value = 1;
				sheet.Cells["A2"].Value = 2;
				sheet.Cells["A3"].Value = 3;
				sheet.Cells["A4"].Value = 4;

				sheet.Cells["B1"].Value = 1;
				sheet.Cells["B2"].Value = 2;
				sheet.Cells["B3"].Value = 3;
				sheet.Cells["B4"].Value = 4;

				sheet.Cells["C1"].Value = ">2";

				sheet.Cells["E1"].Formula = "SUMIFS(A1:A4,B1:B4,C1:C2)";
				sheet.Calculate();

				Assert.AreEqual(7d, sheet.Cells["E1"].Value);
			}
		}

		[TestMethod]
		public void CalculateSingleCriteriaMatchedDimensionsMatchedRootEmptyCriteriaReferenceMatches0()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				sheet.Cells["A1"].Value = 1;
				sheet.Cells["A2"].Value = 2;
				sheet.Cells["A3"].Value = 3;
				sheet.Cells["A4"].Value = 4;

				sheet.Cells["B1"].Value = 0;
				sheet.Cells["B2"].Value = 2;
				sheet.Cells["B3"].Value = 3;
				sheet.Cells["B4"].Value = 4;

				sheet.Cells["E1"].Formula = "SUMIFS(A1:A4,B1:B4,C1)";
				sheet.Calculate();

				Assert.AreEqual(1d, sheet.Cells["E1"].Value);
			}
		}

		[TestMethod]
		public void CalculateSingleCriteriaMatchedDimensionsOffsetRoot()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				sheet.Cells["A1"].Value = 1;
				sheet.Cells["A2"].Value = 2;
				sheet.Cells["A3"].Value = 3;
				sheet.Cells["A4"].Value = 4;

				sheet.Cells["B3"].Value = 1;
				sheet.Cells["B4"].Value = 2;
				sheet.Cells["B5"].Value = 3;
				sheet.Cells["B6"].Value = 4;

				sheet.Cells["E1"].Formula = "SUMIFS(A1:A4,B3:B6,\">2\")";
				sheet.Calculate();

				Assert.AreEqual(7d, sheet.Cells["E1"].Value);
			}
		}

		[TestMethod]
		public void CalculateSingleCriteriaMatchedDimensionsFullColumn()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				sheet.Cells["A1"].Value = 1;
				sheet.Cells["A2"].Value = 2;
				sheet.Cells["A3"].Value = 3;
				sheet.Cells["A4"].Value = 4;

				sheet.Cells["B1"].Value = 1;
				sheet.Cells["B2"].Value = 2;
				sheet.Cells["B3"].Value = 3;
				sheet.Cells["B4"].Value = 4;

				sheet.Cells["E1"].Formula = "SUMIFS(A:A,B:B,\">2\")";
				sheet.Calculate();

				Assert.AreEqual(7d, sheet.Cells["E1"].Value);
			}
		}

		[TestMethod]
		public void CalculateSingleCriteriaMatchedDimensionsFullColumnMissingValuesSumRange()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				sheet.Cells["A1"].Value = 1;
				sheet.Cells["A2"].Value = 2;
				sheet.Cells["A4"].Value = 4;

				sheet.Cells["B1"].Value = 1;
				sheet.Cells["B2"].Value = 2;
				sheet.Cells["B3"].Value = 3;
				sheet.Cells["B4"].Value = 4;

				sheet.Cells["E1"].Formula = "SUMIFS(A:A,B:B, \">2\")";
				sheet.Calculate();

				Assert.AreEqual(4d, sheet.Cells["E1"].Value);
			}
		}

		[TestMethod]
		public void CalculateSingleCriteriaMatchedDimensionsFullColumnMissingValuesCriteriaRange()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				sheet.Cells["A1"].Value = 1;
				sheet.Cells["A2"].Value = 2;
				sheet.Cells["A3"].Value = 3;
				sheet.Cells["A4"].Value = 4;

				sheet.Cells["B1"].Value = 1;
				sheet.Cells["B2"].Value = 2;
				sheet.Cells["B4"].Value = 4;

				sheet.Cells["E1"].Formula = "SUMIFS(A:A,B:B,\">2\")";
				sheet.Calculate();

				Assert.AreEqual(4d, sheet.Cells["E1"].Value);
			}
		}

		[TestMethod]
		public void CalculateSingleCriteriaMatchedMultiDimensionsMatchedRoot()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				sheet.Cells["A1"].Value = 1;
				sheet.Cells["A2"].Value = 2;
				sheet.Cells["A3"].Value = 3;
				sheet.Cells["A4"].Value = 4;
				sheet.Cells["B1"].Value = 10;
				sheet.Cells["B2"].Value = 20;
				sheet.Cells["B3"].Value = 30;
				sheet.Cells["B4"].Value = 40;

				sheet.Cells["C1"].Value = 1;
				sheet.Cells["C2"].Value = 2;
				sheet.Cells["C3"].Value = 3;
				sheet.Cells["C4"].Value = 4;
				sheet.Cells["D1"].Value = 1;
				sheet.Cells["D2"].Value = 2;
				sheet.Cells["D3"].Value = 3;
				sheet.Cells["D4"].Value = 4;

				sheet.Cells["E1"].Formula = "SUMIFS(A1:B4,C1:D4,\">2\")";
				sheet.Calculate();

				Assert.AreEqual(77d, sheet.Cells["E1"].Value);
			}
		}

		[TestMethod]
		public void CalculateSingleCriteriaMatchedMultiDimensionsOffsetRoot()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				sheet.Cells["A1"].Value = 1;
				sheet.Cells["A2"].Value = 2;
				sheet.Cells["A3"].Value = 3;
				sheet.Cells["A4"].Value = 4;
				sheet.Cells["B1"].Value = 10;
				sheet.Cells["B2"].Value = 20;
				sheet.Cells["B3"].Value = 30;
				sheet.Cells["B4"].Value = 40;

				sheet.Cells["C3"].Value = 1;
				sheet.Cells["C4"].Value = 2;
				sheet.Cells["C5"].Value = 3;
				sheet.Cells["C6"].Value = 4;
				sheet.Cells["D3"].Value = 1;
				sheet.Cells["D4"].Value = 2;
				sheet.Cells["D5"].Value = 3;
				sheet.Cells["D6"].Value = 4;

				sheet.Cells["E1"].Formula = "SUMIFS(A1:B4,C3:D6,\">2\")";
				sheet.Calculate();

				Assert.AreEqual(77d, sheet.Cells["E1"].Value);
			}
		}

		[TestMethod]
		public void CalculateSingleCriteriaMisMatchedDimensionsCriteriaTooWide()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				sheet.Cells["E1"].Formula = "SUMIFS(A1:A4,B1:C4,\">2\")";
				sheet.Calculate();

				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["E1"].Value);
			}
		}

		[TestMethod]
		public void CalculateSingleCriteriaMisMatchedDimensionsCriteriaTooThin()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				sheet.Cells["E1"].Formula = "SUMIFS(A1:B4,C1:C4,\">2\")";
				sheet.Calculate();

				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["E1"].Value);
			}
		}

		[TestMethod]
		public void CalculateSingleCriteriaMisMatchedDimensionsCriteriaTooTall()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				sheet.Cells["E1"].Formula = "SUMIFS(A1:A4,C1:C5,\">2\")";
				sheet.Calculate();

				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["E1"].Value);
			}
		}

		[TestMethod]
		public void CalculateSingleCriteriaMisMatchedDimensionsCriteriaTooShort()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				sheet.Cells["E1"].Formula = "SUMIFS(A1:A4,C1:C5,\">2\")";
				sheet.Calculate();

				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["E1"].Value);
			}
		}

		[TestMethod]
		public void CalculateMatchedErrorResultsInErrorPoundValueExample()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				sheet.Cells["A1"].Value = 1;
				sheet.Cells["A2"].Value = 2;
				sheet.Cells["A3"].Value = ExcelErrorValue.Create(eErrorType.Value);
				sheet.Cells["A4"].Value = 4;

				sheet.Cells["B1"].Value = 1;
				sheet.Cells["B2"].Value = 2;
				sheet.Cells["B3"].Value = 3;
				sheet.Cells["B4"].Value = 4;

				sheet.Cells["E1"].Formula = "SUMIFS(A1:A4,B1:B4,\">2\")";
				sheet.Calculate();

				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["E1"].Value);
			}
		}

		[TestMethod]
		public void CalculateMatchedErrorResultsInErrorDiv0Example()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				sheet.Cells["A1"].Value = 1;
				sheet.Cells["A2"].Value = 2;
				sheet.Cells["A3"].Value = ExcelErrorValue.Create(eErrorType.Div0);
				sheet.Cells["A4"].Value = 4;

				sheet.Cells["B1"].Value = 1;
				sheet.Cells["B2"].Value = 2;
				sheet.Cells["B3"].Value = 3;
				sheet.Cells["B4"].Value = 4;

				sheet.Cells["E1"].Formula = "SUMIFS(A1:A4,B1:B4,\">2\")";
				sheet.Calculate();

				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Div0), sheet.Cells["E1"].Value);
			}
		}

		[TestMethod]
		public void CalculateUnMatchedErrorIsIgnoredPoundValueExample()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				sheet.Cells["A1"].Value = 1;
				sheet.Cells["A2"].Value = ExcelErrorValue.Create(eErrorType.Value);
				sheet.Cells["A3"].Value = 3;
				sheet.Cells["A4"].Value = 4;

				sheet.Cells["B1"].Value = 1;
				sheet.Cells["B2"].Value = 2;
				sheet.Cells["B3"].Value = 3;
				sheet.Cells["B4"].Value = 4;

				sheet.Cells["E1"].Formula = "SUMIFS(A1:A4,B1:B4,\">2\")";
				sheet.Calculate();

				Assert.AreEqual(7d, sheet.Cells["E1"].Value);
			}
		}

		[TestMethod]
		public void CalculateUnMatchedErrorIsIgnoredDiv0Example()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				sheet.Cells["A1"].Value = 1;
				sheet.Cells["A2"].Value = ExcelErrorValue.Create(eErrorType.Div0);
				sheet.Cells["A3"].Value = 3;
				sheet.Cells["A4"].Value = 4;

				sheet.Cells["B1"].Value = 1;
				sheet.Cells["B2"].Value = 2;
				sheet.Cells["B3"].Value = 3;
				sheet.Cells["B4"].Value = 4;

				sheet.Cells["E1"].Formula = "SUMIFS(A1:A4,B1:B4,\">2\")";
				sheet.Calculate();

				Assert.AreEqual(7d, sheet.Cells["E1"].Value);
			}
		}

		[TestMethod]
		public void CalculateCriteriaErrorIsIgnoredPoundValueExample()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				sheet.Cells["A1"].Value = 1;
				sheet.Cells["A2"].Value = 2;
				sheet.Cells["A3"].Value = 3;
				sheet.Cells["A4"].Value = 4;

				sheet.Cells["B1"].Value = 1;
				sheet.Cells["B2"].Value = 2;
				sheet.Cells["B3"].Value = ExcelErrorValue.Create(eErrorType.Value);
				sheet.Cells["B4"].Value = 4;

				sheet.Cells["E1"].Formula = "SUMIFS(A1:A4,B1:B4,\">2\")";
				sheet.Calculate();

				Assert.AreEqual(4d, sheet.Cells["E1"].Value);
			}
		}

		[TestMethod]
		public void CalculateCriteriaErrorIsIgnoredDiv0Example()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				sheet.Cells["A1"].Value = 1;
				sheet.Cells["A2"].Value = 2;
				sheet.Cells["A3"].Value = 3;
				sheet.Cells["A4"].Value = 4;

				sheet.Cells["B1"].Value = 1;
				sheet.Cells["B2"].Value = 2;
				sheet.Cells["B3"].Value = ExcelErrorValue.Create(eErrorType.Div0);
				sheet.Cells["B4"].Value = 4;

				sheet.Cells["E1"].Formula = "SUMIFS(A1:A4,B1:B4,\">2\")";
				sheet.Calculate();

				Assert.AreEqual(4d, sheet.Cells["E1"].Value);
			}
		}

		[TestMethod]
		public void CalculateMultipleCriteriaMatchedDimensionsMatchedRoot()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				sheet.Cells["A1"].Value = 1;
				sheet.Cells["A2"].Value = 2;
				sheet.Cells["A3"].Value = 3;
				sheet.Cells["A4"].Value = 4;

				sheet.Cells["B1"].Value = 1;
				sheet.Cells["B2"].Value = 2;
				sheet.Cells["B3"].Value = 3;
				sheet.Cells["B4"].Value = 4;

				sheet.Cells["C1"].Value = 1;
				sheet.Cells["C2"].Value = 2;
				sheet.Cells["C3"].Value = 3;
				sheet.Cells["C4"].Value = 4;

				sheet.Cells["E1"].Formula = "SUMIFS(A1:A4,B1:B4,\">2\",C1:C4,\">3\")";
				sheet.Calculate();

				Assert.AreEqual(4d, sheet.Cells["E1"].Value);
			}
		}

		[TestMethod]
		public void CalculateMultipleCriteriaMatchedDimensionsOffsetRoot()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				sheet.Cells["A1"].Value = 1;
				sheet.Cells["A2"].Value = 2;
				sheet.Cells["A3"].Value = 3;
				sheet.Cells["A4"].Value = 4;

				sheet.Cells["B1"].Value = 1;
				sheet.Cells["B2"].Value = 2;
				sheet.Cells["B3"].Value = 3;
				sheet.Cells["B4"].Value = 4;

				sheet.Cells["C3"].Value = 1;
				sheet.Cells["C4"].Value = 2;
				sheet.Cells["C5"].Value = 3;
				sheet.Cells["C6"].Value = 4;

				sheet.Cells["E1"].Formula = "SUMIFS(A1:A4,B1:B4,\">2\",C3:C6,\">3\")";
				sheet.Calculate();

				Assert.AreEqual(4d, sheet.Cells["E1"].Value);
			}
		}

		[TestMethod]
		public void CalculateMultipleCriteriaMisMatchedDimensionsSecondCriteriaTooWide()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				sheet.Cells["E1"].Formula = "SUMIFS(A1:A4,A5:A8,\">2\",B1:C4,\">2\")";
				sheet.Calculate();

				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["E1"].Value);
			}
		}

		[TestMethod]
		public void CalculateMultipleCriteriaMisMatchedDimensionsSecondCriteriaTooThin()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				sheet.Cells["E1"].Formula = "SUMIFS(A1:B4,A5:B8,\">2\",C1:C4,\">2\")";
				sheet.Calculate();

				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["E1"].Value);
			}
		}

		[TestMethod]
		public void CalculateMultipleCriteriaMisMatchedDimensionsSecondCriteriaTooTall()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				sheet.Cells["E1"].Formula = "SUMIFS(A1:A4,A5:A8,\">2\",C1:C5,\">2\")";
				sheet.Calculate();

				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["E1"].Value);
			}
		}

		[TestMethod]
		public void CalculateMultipleCriteriaMisMatchedDimensionsSecondCriteriaTooShort()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				sheet.Cells["E1"].Formula = "SUMIFS(A1:A4,A5:A8,\">2\",C1:C5,\">2\")";
				sheet.Calculate();

				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["E1"].Value);
			}
		}

		[TestMethod]
		public void CalculateMatchedTextIsIgnored()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				sheet.Cells["A1"].Value = 1;
				sheet.Cells["A2"].Value = 2;
				sheet.Cells["A3"].Value = "a";
				sheet.Cells["A4"].Value = 4;

				sheet.Cells["B1"].Value = 1;
				sheet.Cells["B2"].Value = 2;
				sheet.Cells["B3"].Value = 3;
				sheet.Cells["B4"].Value = 4;

				sheet.Cells["E1"].Formula = "SUMIFS(A1:A4,B1:B4,\">2\")";
				sheet.Calculate();

				Assert.AreEqual(4d, sheet.Cells["E1"].Value);
			}
		}

		[TestMethod]
		public void CalculateMatchedDateIsConverted()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				sheet.Cells["A1"].Value = 1;
				sheet.Cells["A2"].Value = 2;
				sheet.Cells["A3"].Value = DateTime.Parse("1/1/2013");
				sheet.Cells["A4"].Value = 4;

				sheet.Cells["B1"].Value = 1;
				sheet.Cells["B2"].Value = 2;
				sheet.Cells["B3"].Value = 3;
				sheet.Cells["B4"].Value = 4;

				sheet.Cells["E1"].Formula = "SUMIFS(A1:A4,B1:B4,\">2\")";
				sheet.Calculate();

				Assert.AreEqual(41279d, sheet.Cells["E1"].Value);
			}
		}

		[TestMethod]
		public void CalculateMatchImplicitEqualityNumericType()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				sheet.Cells["A1"].Value = 1;
				sheet.Cells["A2"].Value = 2;
				sheet.Cells["A3"].Value = 3;
				sheet.Cells["A4"].Value = 4;

				sheet.Cells["B1"].Value = 1;
				sheet.Cells["B2"].Value = 2;
				sheet.Cells["B3"].Value = 3;
				sheet.Cells["B4"].Value = 4;

				sheet.Cells["E1"].Formula = "SUMIFS(A1:A4,B1:B4,2)";
				sheet.Calculate();

				Assert.AreEqual(2d, sheet.Cells["E1"].Value);
			}
		}

		[TestMethod]
		public void CalculateMatchImplicitEqualityTextType()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				sheet.Cells["A1"].Value = 1;
				sheet.Cells["A2"].Value = 2;
				sheet.Cells["A3"].Value = 3;
				sheet.Cells["A4"].Value = 4;

				sheet.Cells["B1"].Value = 1;
				sheet.Cells["B2"].Value = 2;
				sheet.Cells["B3"].Value = 3;
				sheet.Cells["B4"].Value = 4;

				sheet.Cells["E1"].Formula = "SUMIFS(A1:A4,B1:B4,\"2\")";
				sheet.Calculate();

				Assert.AreEqual(2d, sheet.Cells["E1"].Value);
			}
		}

		[TestMethod]
		public void CalculateMatchExplicitEquality()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				sheet.Cells["A1"].Value = 1;
				sheet.Cells["A2"].Value = 2;
				sheet.Cells["A3"].Value = 3;
				sheet.Cells["A4"].Value = 4;

				sheet.Cells["B1"].Value = 1;
				sheet.Cells["B2"].Value = 2;
				sheet.Cells["B3"].Value = 3;
				sheet.Cells["B4"].Value = 4;

				sheet.Cells["E1"].Formula = "SUMIFS(A1:A4,B1:B4,\"=2\")";
				sheet.Calculate();

				Assert.AreEqual(2d, sheet.Cells["E1"].Value);
			}
		}

		[TestMethod]
		public void CalculateTooFewArguments()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				sheet.Cells["E1"].Formula = "SUMIFS(A1:A4,B1:C4)";
				sheet.Calculate();

				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["E1"].Value);
			}
		}

		[TestMethod]
		public void CalculateCriteriaRangeMissingMatchCriteria()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				sheet.Cells["E1"].Formula = "SUMIFS(A1:A4,B1:C4,\">2\",C1:C4)";
				sheet.Calculate();

				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["E1"].Value);
			}
		}

		// Below are new tests 
		// The following tests test the Range parameter
		[TestMethod]
		public void SumIfsRangeWithOnlyNumbersReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 4;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 4;
				worksheet.Cells["B4"].Value = 2;
				worksheet.Cells["B5"].Formula = "SUMIFS($B$1:$B$2, B3:B4, \"<>-10\")";
				worksheet.Calculate();
				Assert.AreEqual(6d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void SumIfsRangeWithNumericStringAndNumericCriteriaReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 4;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = "4";
				worksheet.Cells["B4"].Value = 2;
				worksheet.Cells["B5"].Formula = "SUMIFS($B$1:$B$2, B3:B4, \">0\")";
				worksheet.Calculate();
				Assert.AreEqual(2d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void SumIfsRangeWithNumericStringNonNumericCriteriaReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 4;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = "4";
				worksheet.Cells["B4"].Value = 2;
				worksheet.Cells["B5"].Formula = "SUMIFS($B$1:$B$2, B3:B4, \"<>-10\")";
				worksheet.Calculate();
				Assert.AreEqual(6d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void SumIfsRangeWithNonNumericStringAndTextCriteriaReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 4;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = "word";
				worksheet.Cells["B4"].Value = 2;
				worksheet.Cells["B5"].Formula = "SUMIFS($B$1:$B$2, B3:B4, \"word\")";
				worksheet.Calculate();
				Assert.AreEqual(4d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void SumIfsRangeWithNonNumericStringAndNonNumericCriteriaReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 4;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = "word";
				worksheet.Cells["B4"].Value = 2;
				worksheet.Cells["B5"].Formula = "SUMIFS($B$1:$B$2, B3:B4, \"<>-10\")";
				worksheet.Calculate();
				Assert.AreEqual(6d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void SumIfsRangeWithBooleanValueAndBooleanCriteriaReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 4;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = true;
				worksheet.Cells["B4"].Value = 2;
				worksheet.Cells["B5"].Formula = "SUMIFS($B$1:$B$2, B3:B4, \"TRUE\")";
				worksheet.Calculate();
				Assert.AreEqual(4d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void SumIfsRangeWithBooleanValueAndNonNumericCriteriaReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 4;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = false;
				worksheet.Cells["B4"].Value = 2;
				worksheet.Cells["B5"].Formula = "SUMIFS($B$1:$B$2, B3:B4, \"<>-10\")";
				worksheet.Calculate();
				Assert.AreEqual(6d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void SumIfsRangeWithDateAndDateCriteriaReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 4;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Formula = "DATE(2017, 6, 22)";
				worksheet.Cells["B4"].Value = 2;
				worksheet.Cells["B5"].Formula = "SUMIFS($B$1:$B$2, B3:B4, \"<6/23/2017\")";
				worksheet.Calculate();
				Assert.AreEqual(6d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void SumIfsRangeWithDateAndNonNumericCriteriaReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 4;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Formula = "DATE(2017, 6, 22)";
				worksheet.Cells["B4"].Value = 2;
				worksheet.Cells["B5"].Formula = "SUMIFS($B$1:$B$2, B3:B4, \"<>-10\")";
				worksheet.Calculate();
				Assert.AreEqual(6d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void SumIfsWithRangeWithDateInStringReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 4;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = "6/22/2017";
				worksheet.Cells["B4"].Value = 2;
				worksheet.Cells["B5"].Formula = "SUMIFS($B$1:$B$2, B3:B4, \"6/22/2017\")";
				worksheet.Calculate();
				Assert.AreEqual(4d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void SumIfsWithRangeWithEmptyCellAndEmptyStringCriteriaReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 4;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = null;
				worksheet.Cells["B4"].Value = 2;
				worksheet.Cells["B5"].Formula = "SUMIFS($B$1:$B$2, B3:B4, \"\")";
				worksheet.Calculate();
				Assert.AreEqual(4d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void SumIfsWithRangeWithEmptyCellAndNonNumericCriteraReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 4;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = null;
				worksheet.Cells["B4"].Value = 2;
				worksheet.Cells["B5"].Formula = "SUMIFS($B$1:$B$2, B3:B4, \"<>-10\")";
				worksheet.Calculate();
				Assert.AreEqual(6d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void SumIfsWithRangeWithErrorValueReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 4;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Formula = "notaformula";
				worksheet.Cells["B4"].Value = 2;
				worksheet.Cells["B5"].Formula = "SUMIFS($B$1:$B$2, B3:B4, \"<>-10\")";
				worksheet.Calculate();
				Assert.AreEqual(6d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void SumIfsCriteriaTestsReturnsCorrectValues()
		{
			//These tests are in one because they all use the same Excel Package of data
			var package = this.CreateTestingPackage();
			var worksheet = package.Workbook.Worksheets["Sheet1"];

			worksheet.Cells["B2"].Formula = "SUMIFS(C3:W3,C2:W2,5)";
			worksheet.Cells["B3"].Formula = "SUMIFS(C3:W3,C2:W2,\"5\")";
			worksheet.Cells["B4"].Formula = "SUMIFS(C3:W3,C2:W2,\"=5\")";
			worksheet.Cells["B5"].Formula = "SUMIFS(C3:W3,C2:W2,3.5)";
			worksheet.Cells["B6"].Formula = "SUMIFS(C3:W3,C2:W2,TRUE)";
			worksheet.Cells["B7"].Formula = "SUMIFS(C3:W3,C2:W2,\"TRUE\")";
			worksheet.Cells["B8"].Formula = "SUMIFS(C3:W3,C2:W2,\"6/23/2017\")";
			worksheet.Cells["B9"].Formula = "SUMIFS(C3:W3,C2:W2,\"6:00 PM\")";
			worksheet.Cells["B10"].Formula = "SUMIFS(C3:W3,C2:W2,\"T*sday\")";
			worksheet.Cells["B11"].Formula = "SUMIFS(C3:W3,C2:W2,\">=1\")";
			worksheet.Cells["B12"].Formula = "SUMIFS(C3:W3,C2:W2,\">6/22/2017\")";
			worksheet.Cells["B13"].Formula = "SUMIFS(F11:H11,C11:E11,1)";
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
		public void SumIfsAverageRangeTestsReturnsCorrectValues()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["C2"].Value = 3;
				worksheet.Cells["D2"].Value = 1;
				worksheet.Cells["B3"].Formula = "SUMIFS(C3:D3,C2:D2,\">0\")";
				worksheet.Cells["C3"].Value = 4;
				worksheet.Cells["D3"].Value = 2;
				worksheet.Cells["B4"].Formula = "SUMIFS(C4:D4,C2:D2,\">0\")";
				worksheet.Cells["C4"].Value = "4";
				worksheet.Cells["D4"].Value = 2;
				worksheet.Cells["B5"].Formula = "SUMIFS(C5:D5,C2:D2,\">0\")";
				worksheet.Cells["C5"].Value = "word";
				worksheet.Cells["D5"].Value = 2;
				worksheet.Cells["B6"].Formula = "SUMIFS(C6:D6,C2:D2,\">0\")";
				worksheet.Cells["C6"].Value = true;
				worksheet.Cells["D6"].Value = 2;
				worksheet.Cells["B7"].Formula = "SUMIFS(C7:D7,C2:D2,\">0\")";
				worksheet.Cells["C7"].Value = (new DateTime(2017, 6, 22)).ToOADate();
				worksheet.Cells["D7"].Value = 2;
				worksheet.Cells["B8"].Formula = "SUMIFS(C8:D8,C2:D2,\">0\")";
				worksheet.Cells["C8"].Value = "6/22/2017";
				worksheet.Cells["D8"].Value = 2;
				worksheet.Cells["B9"].Formula = "SUMIFS(C9:D9,C2:D2,\">0\")";
				worksheet.Cells["C9"].Value = null;
				worksheet.Cells["D9"].Value = 2;
				worksheet.Cells["B10"].Formula = "SUMIFS(C10:D10,C2:D2,\">0\")";
				worksheet.Cells["C10"].Formula = "notAValidFormula"; // Evaluates to #NAME.
				worksheet.Cells["D10"].Value = 2;
				worksheet.Calculate();
				Assert.AreEqual(6d, worksheet.Cells["B3"].Value);
				Assert.AreEqual(2d, worksheet.Cells["B4"].Value);
				Assert.AreEqual(2d, worksheet.Cells["B5"].Value);
				Assert.AreEqual(2d, worksheet.Cells["B6"].Value);
				Assert.AreEqual(42910d, worksheet.Cells["B7"].Value);
				Assert.AreEqual(2d, worksheet.Cells["B8"].Value);
				Assert.AreEqual(2d, worksheet.Cells["B9"].Value);
				Assert.AreEqual(eErrorType.Name, ((ExcelErrorValue)worksheet.Cells["B10"].Value).Type);
			}
		}

		[TestMethod]
		public void SumIfsCriteriaCaseSensitivityReturnsCorrectValues()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "aa";
				worksheet.Cells["B2"].Value = "ab";
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["B4"].Value = 1;
				worksheet.Cells["B5"].Formula = "SUMIFS(B3:B4, B1:B2, \"?b\")";
				worksheet.Cells["B6"].Formula = "SUMIFS(B3:B4, B1:B2, \"*B\")";
				worksheet.Calculate();
				Assert.AreEqual(1d, worksheet.Cells["B5"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B6"].Value);
			}
		}

		[TestMethod]
		public void SumIfsCriteriaCellRangesReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "aa";
				worksheet.Cells["B2"].Value = "ab";
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["B4"].Value = 1;
				worksheet.Cells["B5"].Value = "*b";
				worksheet.Cells["B6"].Formula = "SUMIFS(B3:B4, $B$1:$B$2, B5)";
				worksheet.Cells["B7"].Formula = "SUMIFS(B3:B4, $B$1:$B$2, B1:B2)";
				worksheet.Calculate();
				Assert.AreEqual(1d, worksheet.Cells["B6"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B7"].Value);
			}
		}

		[TestMethod]
		public void SumIfsCriteriaAsArraysReturnsCorrectValues()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 3;
				worksheet.Cells["B3"].Value = 5;
				worksheet.Cells["B4"].Formula = "{1,2,3}";
				worksheet.Cells["B5"].Formula = "{1,2,3}";
				worksheet.Cells["B6"].Formula = "{1,2,3}";
				worksheet.Cells["B7"].Formula = "SUMIFS(B1:B3, B4:B6, {1,2,3})";
				worksheet.Cells["B8"].Formula = "SUMIFS(B1:B3, B4:B6, B4)";
				worksheet.Calculate();
				Assert.AreEqual(9d, worksheet.Cells["B7"].Value);
				Assert.AreEqual(9d, worksheet.Cells["B8"].Value);
			}
		}

		[TestMethod]
		public void SumIfsWithStringCriteriasReturnCorrectValues()
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
				worksheet.Cells["D2"].Formula = "SUMIFS(C1:C3, B1:B3, \"\")";
				worksheet.Calculate();
				Assert.AreEqual(3d, worksheet.Cells["D1"].Value);
				Assert.AreEqual(4d, worksheet.Cells["D2"].Value);
			}
		}

		[TestMethod]
		public void SumIfsWithBooleanComparisonsReturnCorrectValues()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = true;
				worksheet.Cells["B2"].Value = 1;
				worksheet.Cells["B3"].Value = false;
				worksheet.Cells["B4"].Formula = "SUMIFS(B2, B1, \">TRUE\")";
				worksheet.Cells["B5"].Formula = "SUMIFS(B2, B3, \"<TRUE\")";
				worksheet.Cells["B6"].Formula = "SUMIFS(B2, B1, \">FALSE\")";
				worksheet.Cells["B7"].Formula = "SUMIFS(B2, B3, \"<FALSE\")";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B4"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B5"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B6"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B7"].Value);
			}
		}

		[TestMethod]
		public void SumIfsWithExpressionCharacterCriteriaReturnCorrectValues()
		{

				using (var package = new ExcelPackage())
				{
					var worksheet = package.Workbook.Worksheets.Add("Sheet1");
					worksheet.Cells["C2"].Value = "=";
					worksheet.Cells["C3"].Value = "";
					worksheet.Cells["C4"].Value = null;
					worksheet.Cells["D2"].Value = 1;
					worksheet.Cells["B2"].Formula = "SUMIFS(D2,C2,\"=\")";
					worksheet.Cells["B3"].Formula = "SUMIFS(D2,C3,\"=\")";
					worksheet.Cells["B4"].Formula = "SUMIFS(D2,C4,\"=\")";
					worksheet.Calculate();
					Assert.AreEqual(0d, worksheet.Cells["B2"].Value);
					Assert.AreEqual(0d, worksheet.Cells["B3"].Value);
					Assert.AreEqual(1d, worksheet.Cells["B4"].Value);
				}
		}

		[TestMethod]
		public void SumIfsWithStringComparisonsReturnCorrectValues()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "ay";
				worksheet.Cells["B2"].Value = 1;
				worksheet.Cells["B3"].Formula = "SUMIFS($B$2, $B$1, \"<axz\")";
				worksheet.Cells["B4"].Formula = "SUMIFS($B$2, $B$1, \"<aya\")";
				worksheet.Cells["B5"].Formula = "SUMIFS($B$2, $B$1, \"<az\")";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B3"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B4"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B5"].Value);
			}
		}

		[TestMethod]
		public void SumIfsWithComparisonsWithWildcardCharacterReturnCorrectValues()
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
				worksheet.Cells["B2"].Formula = "SUMIFS(D2:D4,C2:C4,\"=Mo*day\")";
				worksheet.Cells["B3"].Formula = "SUMIFS(D2:D4,C2:C4,\">Mo*day\")";
				worksheet.Cells["B4"].Formula = "SUMIFS(D2:D4,C2:C4,\"<Mo*day\")";
				worksheet.Calculate();
				Assert.AreEqual(8d, worksheet.Cells["B2"].Value);
				Assert.AreEqual(8d, worksheet.Cells["B3"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B4"].Value);
			}
		}

		[TestMethod]
		public void SumIfsWithEscapedWildcardCharacterReturnCorrectValues()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = "Mon?ay";
				worksheet.Cells["B2"].Value = "Monday";
				worksheet.Cells["B3"].Value = "Mon*ay";
				worksheet.Cells["B4"].Value = "Monddday";
				worksheet.Cells["C1"].Value = 1;
				worksheet.Cells["C2"].Value = 3;
				worksheet.Cells["C3"].Value = 5;
				worksheet.Cells["C4"].Value = 7;
				worksheet.Cells["D1"].Formula = "SUMIFS(C1:C2, B1:B2, \"Mon?ay\")";
				worksheet.Cells["D2"].Formula = "SUMIFS(C1:C2, B1:B2, \"Mon~?ay\")";
				worksheet.Cells["D3"].Formula = "SUMIFS(C3:C4, B3:B4, \"Mon*ay\")";
				worksheet.Cells["D4"].Formula = "SUMIFS(C3:C4, B3:B4, \"Mon~*ay\")";
				worksheet.Calculate();
				Assert.AreEqual(4d, worksheet.Cells["D1"].Value);
				Assert.AreEqual(1d, worksheet.Cells["D2"].Value);
				Assert.AreEqual(12d, worksheet.Cells["D3"].Value);
				Assert.AreEqual(5d, worksheet.Cells["D4"].Value);
			}
		}

		[TestMethod]
		public void SumIfsWithStringComparisonsCellReferencesReturnCorrectValues()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = null;
				worksheet.Cells["B2"].Value = 1;
				worksheet.Cells["B3"].Formula = "SUMIFS(B2, B1, \">a\")";
				worksheet.Cells["B4"].Formula = "SUMIFS(B2, B1, \"<a\")";
				worksheet.Cells["B5"].Value = "";
				worksheet.Cells["B6"].Formula = "SUMIFS(B2, B5, \">a\")";
				worksheet.Cells["B7"].Formula = "SUMIFS(B2, B5, \"<a\")";
				worksheet.Cells["B8"].Value = "zzz";
				worksheet.Cells["B9"].Formula = "SUMIFS(B2, B8, \">a\")";
				worksheet.Cells["B10"].Formula = "SUMIFS(B2, B8, \"<a\")";
				worksheet.Cells["B11"].Value = 1;
				worksheet.Cells["B12"].Formula = "SUMIFS(B2, B11, \">a\")";
				worksheet.Cells["B13"].Formula = "SUMIFS(B2, B11, \"<a\")";
				worksheet.Cells["B14"].Value = "1";
				worksheet.Cells["B15"].Formula = "SUMIFS(B2, B14, \">a\")";
				worksheet.Cells["B16"].Formula = "SUMIFS(B2, B14, \"<a\")";
				worksheet.Cells["B17"].Value = true;
				worksheet.Cells["B18"].Formula = "SUMIFS(B2, B17, \">a\")";
				worksheet.Cells["B19"].Formula = "SUMIFS(B2, B17, \"<a\")";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["B3"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B4"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B6"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B7"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B9"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B10"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B12"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B13"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B15"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B16"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B18"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B19"].Value);
			}
		}

		[TestMethod]
		public void SumIfsWithRegexExpressionCharactersRetunCorrectValues()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 1;
				worksheet.Cells["B3"].Value = ">1";
				worksheet.Cells["B4"].Value = "<1";
				worksheet.Cells["B5"].Value = "=1";
				worksheet.Cells["C1"].Formula = "SUMIFS(B1, B2, \"=1\")";
				worksheet.Cells["C2"].Formula = "SUMIFS(B1, B2, \">1\")";
				worksheet.Cells["C3"].Formula = "SUMIFS(B1, B2, \"<1\")";
				worksheet.Cells["C4"].Formula = "SUMIFS(B1, B2, \">=1\")";
				worksheet.Cells["C5"].Formula = "SUMIFS(B1, B2, \"<=1\")";
				worksheet.Cells["C6"].Formula = "SUMIFS(B1, B2, \"<>1\")";
				worksheet.Cells["C7"].Formula = "SUMIFS(B1, B3, \"=>1\")";
				worksheet.Cells["C8"].Formula = "SUMIFS(B1, B4, \"=<1\")";
				worksheet.Cells["C9"].Formula = "SUMIFS(B1, B5, \"==1\")";
				worksheet.Cells["C10"].Formula = "SUMIFS(B1, B2, \">>1\")";
				worksheet.Cells["C11"].Formula = "SUMIFS(B1, B2, \"><1\")";
				worksheet.Cells["C12"].Formula = "SUMIFS(B1, B2, \"<<1\")";
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
		public void SumIfsWithRangeAndAverageRangeDifferentSizesReturnsPoundValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["B4"].Value = 4;
				worksheet.Cells["B5"].Value = 5;
				worksheet.Cells["B6"].Value = 6;
				worksheet.Cells["B7"].Value = 1;
				worksheet.Cells["B8"].Formula = "SUMIFS(B7, B1:B6, \">0\")";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)worksheet.Cells["B8"].Value).Type);
			}
		}

		[TestMethod]
		public void SumIfsWithNullCriteriaReturns0()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B2"].Formula = "SUMIFS(C2:C3,D2:D3,1)";
				worksheet.Cells["B3"].Formula = "SUMIFS(C2:C3,D2:D3,1,E2:E3,)";
				worksheet.Cells["B4"].Formula = "SUMIFS(C2:C3,D2:D3,1,E2:E3,E3)";
				worksheet.Cells["C2"].Value = 1.5;
				worksheet.Cells["C3"].Value = 2.5;
				worksheet.Cells["D2"].Value = 1;
				worksheet.Cells["D3"].Value = 1;
				worksheet.Cells["E2"].Value = null;
				worksheet.Cells["E3"].Value = null;
				worksheet.Calculate();
				Assert.AreEqual(4d, worksheet.Cells["B2"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B3"].Value);
				Assert.AreEqual(0d, worksheet.Cells["B4"].Value);
			}
		}

		[TestMethod]
		public void SumIfsWithEmptySumRangeReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells[2, 2].Formula = "SUMIFS(C3:C6, D3:D6, E5)";
				worksheet.Cells[3, 4].Value = 1;
				worksheet.Cells[4, 4].Value = 2;
				worksheet.Cells[5, 4].Value = 3;
				worksheet.Cells[6, 4].Value = 4;
				worksheet.Cells[5, 5].Value = 2;
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells[2, 2].Value);
			}
		}

		[TestMethod]
		public void SumIfsWithEmptyMatchValueReturnsZero()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells[2, 2].Formula = "SUMIFS(C3:C6, D3:D6, E5)";

				worksheet.Cells[3, 3].Value = 1;
				worksheet.Cells[4, 3].Value = 2;
				// worksheet.Cells[5, 3] is left empty.
				worksheet.Cells[6, 3].Value = 4;

				worksheet.Cells[3, 4].Value = 1;
				worksheet.Cells[4, 4].Value = 2;
				worksheet.Cells[5, 4].Value = 3;
				worksheet.Cells[6, 4].Value = 4;

				worksheet.Cells[5, 5].Value = 3;
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells[2, 2].Value);
			}
		}

		[TestMethod]
		public void SumIfsWithUnsetEmptyCellsInCriteria()
		{
			// This test exists to ensure that cells that have never been set are still 
			// being compared against the criterion.
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B2"].Formula = "SUMIFS(C2:C3,D2:D3,\"\")";
				worksheet.Cells["C2"].Value = 1;
				worksheet.Cells["C3"].Value = 2;
				worksheet.Calculate();
				Assert.AreEqual(3d, worksheet.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void SumIfsErrorPropogation()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B2"].Formula = @"SUMIFS(OFFSET(A1,0,MATCH(""A"",A1,0)),D2:D3,"""")";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)worksheet.Cells["B2"].Value).Type);
			}
		}
		#endregion
	}
}
