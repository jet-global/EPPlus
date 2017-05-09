using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest.FormulaParsing.Excel.Functions
{
	[TestClass]
	public class SumIfsTests
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

				sheet.Cells["E1"].Formula = "SUMIFS(A:A,B:B,\">2\")";
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
		#endregion
	}
}
