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
	public class AverageIfTests
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

		#region AverageIf Tests
		[TestMethod]
		public void AverageIfWithOnlyRangeAndCriteriaWorksAsExpected()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["C2"].Value = 1;
				worksheet.Cells["D2"].Value = 3;
				worksheet.Cells["B2"].Formula = "AVERAGEIF(C2:D2,\">0\")";
				worksheet.Cells["C3"].Value = 1;
				worksheet.Cells["D3"].Value = 2;
				worksheet.Cells["B3"].Formula = "AVERAGEIF(C3:D3,1)";
				worksheet.Cells["C4"].Value = null;
				worksheet.Cells["D4"].Value = 1;
				worksheet.Cells["B4"].Formula = "AVERAGEIF(C4:D4,\"<10\")";
				worksheet.Cells["C5"].Value = "2";
				worksheet.Cells["D5"].Value = 1;
				worksheet.Cells["B5"].Formula = "AVERAGEIF(C5:D5,\">0\")";
				worksheet.Cells["C6"].Value = "word";
				worksheet.Cells["D6"].Value = 1;
				worksheet.Cells["B6"].Formula = "AVERAGEIF(C6:D6,\">0\")";
				worksheet.Cells["C7"].Value = true;
				worksheet.Cells["D7"].Value = 1;
				worksheet.Cells["B7"].Formula = "AVERAGEIF(C7:D7,\">0\")";
				worksheet.Cells["C8"].Value = "6/16/2017";
				worksheet.Cells["D8"].Value = 1;
				worksheet.Cells["B8"].Formula = "AVERAGEIF(C8:D8,\">0\")";
				worksheet.Calculate();
				Assert.AreEqual(2d, worksheet.Cells["B2"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B3"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B4"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B5"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B6"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B7"].Value);
				Assert.AreEqual(1d, worksheet.Cells["B8"].Value);
			}
		}

		[TestMethod]
		public void AverageIfWithRangeAndCriteriaAndAverageRangeWorksAsExpected()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["C2"].Value = "aa";
				worksheet.Cells["D2"].Value = "ab";
				worksheet.Cells["E2"].Value = "ac";
				worksheet.Cells["F2"].Value = "ba";
				worksheet.Cells["G2"].Value = "bb";
				worksheet.Cells["H2"].Value = "bc";

				worksheet.Cells["C3"].Value = 1.5;
				worksheet.Cells["D3"].Value = 2;
				worksheet.Cells["E3"].Value = 3.5;
				worksheet.Cells["F3"].Value = 7;
				worksheet.Cells["G3"].Value = 1;
				worksheet.Cells["H3"].Value = 3;
				worksheet.Cells["B3"].Formula = "AVERAGEIF(B2:G2,\"*\",C3:H3)";
				worksheet.Cells["C4"].Value = 1.5;
				worksheet.Cells["D4"].Value = 2;
				worksheet.Cells["E4"].Value = "3.5";
				worksheet.Cells["F4"].Value = 7;
				worksheet.Cells["G4"].Value = 1;
				worksheet.Cells["H4"].Value = 3;
				worksheet.Cells["B4"].Formula = "AVERAGEIF(B2:G2,\"*\",C4:H4)";
				worksheet.Cells["C5"].Value = 1.5;
				worksheet.Cells["D5"].Value = 2;
				worksheet.Cells["E5"].Value = "word";
				worksheet.Cells["F5"].Value = 7;
				worksheet.Cells["G5"].Value = 1;
				worksheet.Cells["H5"].Value = 3;
				worksheet.Cells["B5"].Formula = "AVERAGEIF(B2:G2,\"*\",C5:H5)";
				worksheet.Cells["C6"].Value = 1.5;
				worksheet.Cells["D6"].Value = 2;
				worksheet.Cells["E6"].Value = true;
				worksheet.Cells["F6"].Value = 7;
				worksheet.Cells["G6"].Value = 1;
				worksheet.Cells["H6"].Value = 3;
				worksheet.Cells["B6"].Formula = "AVERAGEIF(B2:G2,\"*\",C6:H6)";
				worksheet.Cells["C7"].Value = 1.5;
				worksheet.Cells["D7"].Value = 2;
				worksheet.Cells["E7"].Value = null;
				worksheet.Cells["F7"].Value = 7;
				worksheet.Cells["G7"].Value = 1;
				worksheet.Cells["H7"].Value = 3;
				worksheet.Cells["B7"].Formula = "AVERAGEIF(B2:G2,\"*\",C7:H7)";
				worksheet.Cells["C8"].Value = 1.5;
				worksheet.Cells["D8"].Value = 2;
				worksheet.Cells["E8"].Value = "6/16/2017";
				worksheet.Cells["F8"].Value = 7;
				worksheet.Cells["G8"].Value = 1;
				worksheet.Cells["H8"].Value = 3;
				worksheet.Cells["B8"].Formula = "AVERAGEIF(B2:G2,\"*\",C8:H8)";
				worksheet.Calculate();
				Assert.AreEqual(3d, worksheet.Cells["B3"].Value);
				Assert.AreEqual(2.9, worksheet.Cells["B4"].Value);
				Assert.AreEqual(2.9, worksheet.Cells["B5"].Value);
				Assert.AreEqual(2.9, worksheet.Cells["B6"].Value);
				Assert.AreEqual(2.9, worksheet.Cells["B7"].Value);
				Assert.AreEqual(2.9, worksheet.Cells["B8"].Value);
			}
		}

		[TestMethod]
		public void AverageIfWithCriteriaFilteringWorksAsExpected()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["C2"].Value = "aa";
				worksheet.Cells["D2"].Value = "ab";
				worksheet.Cells["E2"].Value = "ac";
				worksheet.Cells["F2"].Value = "ba";
				worksheet.Cells["G2"].Value = "bb";
				worksheet.Cells["H2"].Value = "bc";

				worksheet.Cells["C3"].Value = 1.5;
				worksheet.Cells["D3"].Value = 2;
				worksheet.Cells["E3"].Value = 3.5;
				worksheet.Cells["F3"].Value = 7;
				worksheet.Cells["G3"].Value = 1;
				worksheet.Cells["H3"].Value = 3;
				worksheet.Cells["B3"].Formula = "AVERAGEIF(C2:H2,\"*b*\",C3:H3)";
				worksheet.Cells["B4"].Formula = "AVERAGEIF(C2:H2,\"*B*\",C3:H3)";
				
				worksheet.Calculate();

				Assert.AreEqual(3.25, worksheet.Cells["B3"].Value);
				Assert.AreEqual(3.25, worksheet.Cells["B4"].Value);
			}
		}

		//[TestMethod]
		//public void AverageIfWith()
		//{
		//	using (var package = new ExcelPackage())
		//	{
		//		var worksheet = package.Workbook.Worksheets.Add("Sheet1");

		//		worksheet.Cells["C"].Value = "";
		//		worksheet.Cells["B"].Formula = "AVERAGEIF()";

		//		worksheet.Calculate();

		//		Assert.AreEqual(, worksheet.Cells["B"].Value);
		//	}
		//}

		[TestMethod]
		public void AverageIfNumeric()
		{
			_worksheet.Cells["A1"].Value = 1d;
			_worksheet.Cells["A2"].Value = 2d;
			_worksheet.Cells["A3"].Value = 3d;
			_worksheet.Cells["B1"].Value = 1d;
			_worksheet.Cells["B2"].Value = 3d;
			_worksheet.Cells["B3"].Value = 5d;
			var func = new AverageIf();
			IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
			IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
			var args = FunctionsHelper.CreateArgs(range1, ">1", range2);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(4d, result.Result);
		}

		[TestMethod]
		public void AverageIfNonNumeric()
		{
			_worksheet.Cells["A1"].Value = "Monday";
			_worksheet.Cells["A2"].Value = "Tuesday";
			_worksheet.Cells["A3"].Value = "Thursday";
			_worksheet.Cells["B1"].Value = 1d;
			_worksheet.Cells["B2"].Value = 3d;
			_worksheet.Cells["B3"].Value = 5d;
			var func = new AverageIf();
			IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
			IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
			var args = FunctionsHelper.CreateArgs(range1, "T*day", range2);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(4d, result.Result);
		}

		[TestMethod]
		public void AverageIfNumericExpression()
		{
			_worksheet.Cells["A1"].Value = null;
			_worksheet.Cells["A2"].Value = 1d;
			_worksheet.Cells["A3"].Value = "Not Empty";
			var func = new AverageIf();
			IRangeInfo range = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
			var args = FunctionsHelper.CreateArgs(range, 1d);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(1d, result.Result);
		}

		[TestMethod]
		public void AverageIfEqualToEmptyString()
		{
			_worksheet.Cells["A1"].Value = null;
			_worksheet.Cells["A2"].Value = string.Empty;
			_worksheet.Cells["A3"].Value = "Not Empty";
			_worksheet.Cells["B1"].Value = 1d;
			_worksheet.Cells["B2"].Value = 3d;
			_worksheet.Cells["B3"].Value = 5d;
			var func = new AverageIf();
			IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
			IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
			var args = FunctionsHelper.CreateArgs(range1, "", range2);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(1d, result.Result);
		}

		[TestMethod]
		public void AverageIfNotEqualToNull()
		{
			_worksheet.Cells["A1"].Value = null;
			_worksheet.Cells["A2"].Value = string.Empty;
			_worksheet.Cells["A3"].Value = "Not Empty";
			_worksheet.Cells["B1"].Value = 1d;
			_worksheet.Cells["B2"].Value = 3d;
			_worksheet.Cells["B3"].Value = 5d;
			var func = new AverageIf();
			IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
			IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
			var args = FunctionsHelper.CreateArgs(range1, "<>", range2);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(4d, result.Result);
		}

		[TestMethod]
		public void AverageIfEqualToZero()
		{
			_worksheet.Cells["A1"].Value = null;
			_worksheet.Cells["A2"].Value = string.Empty;
			_worksheet.Cells["A3"].Value = 0d;
			_worksheet.Cells["B1"].Value = 1d;
			_worksheet.Cells["B2"].Value = 3d;
			_worksheet.Cells["B3"].Value = 5d;
			var func = new AverageIf();
			IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
			IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
			var args = FunctionsHelper.CreateArgs(range1, "0", range2);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(5d, result.Result);
		}

		[TestMethod]
		public void AverageIfNotEqualToZero()
		{
			_worksheet.Cells["A1"].Value = null;
			_worksheet.Cells["A2"].Value = string.Empty;
			_worksheet.Cells["A3"].Value = 0d;
			_worksheet.Cells["B1"].Value = 1d;
			_worksheet.Cells["B2"].Value = 3d;
			_worksheet.Cells["B3"].Value = 5d;
			var func = new AverageIf();
			IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
			IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
			var args = FunctionsHelper.CreateArgs(range1, "<>0", range2);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(2d, result.Result);
		}

		[TestMethod]
		public void AverageIfGreaterThanZero()
		{
			_worksheet.Cells["A1"].Value = null;
			_worksheet.Cells["A2"].Value = string.Empty;
			_worksheet.Cells["A3"].Value = 1d;
			_worksheet.Cells["B1"].Value = 1d;
			_worksheet.Cells["B2"].Value = 3d;
			_worksheet.Cells["B3"].Value = 5d;
			var func = new AverageIf();
			IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
			IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
			var args = FunctionsHelper.CreateArgs(range1, ">0", range2);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(5d, result.Result);
		}

		[TestMethod]
		public void AverageIfGreaterThanOrEqualToZero()
		{
			_worksheet.Cells["A1"].Value = null;
			_worksheet.Cells["A2"].Value = string.Empty;
			_worksheet.Cells["A3"].Value = 1d;
			_worksheet.Cells["B1"].Value = 1d;
			_worksheet.Cells["B2"].Value = 3d;
			_worksheet.Cells["B3"].Value = 5d;
			var func = new AverageIf();
			IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
			IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
			var args = FunctionsHelper.CreateArgs(range1, ">=0", range2);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(5d, result.Result);
		}

		[TestMethod]
		public void AverageIfLessThanZero()
		{
			_worksheet.Cells["A1"].Value = null;
			_worksheet.Cells["A2"].Value = string.Empty;
			_worksheet.Cells["A3"].Value = -1d;
			_worksheet.Cells["B1"].Value = 1d;
			_worksheet.Cells["B2"].Value = 3d;
			_worksheet.Cells["B3"].Value = 5d;
			var func = new AverageIf();
			IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
			IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
			var args = FunctionsHelper.CreateArgs(range1, "<0", range2);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(5d, result.Result);
		}

		[TestMethod]
		public void AverageIfLessThanOrEqualToZero()
		{
			_worksheet.Cells["A1"].Value = null;
			_worksheet.Cells["A2"].Value = string.Empty;
			_worksheet.Cells["A3"].Value = -1d;
			_worksheet.Cells["B1"].Value = 1d;
			_worksheet.Cells["B2"].Value = 3d;
			_worksheet.Cells["B3"].Value = 5d;
			var func = new AverageIf();
			IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
			IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
			var args = FunctionsHelper.CreateArgs(range1, "<=0", range2);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(5d, result.Result);
		}

		[TestMethod]
		public void AverageIfLessThanCharacter()
		{
			_worksheet.Cells["A1"].Value = null;
			_worksheet.Cells["A2"].Value = string.Empty;
			_worksheet.Cells["A3"].Value = "Not Empty";
			_worksheet.Cells["B1"].Value = 1d;
			_worksheet.Cells["B2"].Value = 3d;
			_worksheet.Cells["B3"].Value = 5d;
			var func = new AverageIf();
			IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
			IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
			var args = FunctionsHelper.CreateArgs(range1, "<a", range2);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(3d, result.Result);
		}

		[TestMethod]
		public void AverageIfLessThanOrEqualToCharacter()
		{
			_worksheet.Cells["A1"].Value = null;
			_worksheet.Cells["A2"].Value = string.Empty;
			_worksheet.Cells["A3"].Value = "Not Empty";
			_worksheet.Cells["B1"].Value = 1d;
			_worksheet.Cells["B2"].Value = 3d;
			_worksheet.Cells["B3"].Value = 5d;
			var func = new AverageIf();
			IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
			IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
			var args = FunctionsHelper.CreateArgs(range1, "<=a", range2);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(3d, result.Result);
		}

		[TestMethod]
		public void AverageIfGreaterThanCharacter()
		{
			_worksheet.Cells["A1"].Value = null;
			_worksheet.Cells["A2"].Value = string.Empty;
			_worksheet.Cells["A3"].Value = "Not Empty";
			_worksheet.Cells["B1"].Value = 1d;
			_worksheet.Cells["B2"].Value = 3d;
			_worksheet.Cells["B3"].Value = 5d;
			var func = new AverageIf();
			IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
			IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
			var args = FunctionsHelper.CreateArgs(range1, ">a", range2);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(5d, result.Result);
		}

		[TestMethod]
		public void AverageIfGreaterThanOrEqualToCharacter()
		{
			_worksheet.Cells["A1"].Value = null;
			_worksheet.Cells["A2"].Value = string.Empty;
			_worksheet.Cells["A3"].Value = "Not Empty";
			_worksheet.Cells["B1"].Value = 1d;
			_worksheet.Cells["B2"].Value = 3d;
			_worksheet.Cells["B3"].Value = 5d;
			var func = new AverageIf();
			IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
			IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
			var args = FunctionsHelper.CreateArgs(range1, ">=a", range2);
			var result = func.Execute(args, _parsingContext);
			Assert.AreEqual(5d, result.Result);
		}

		[TestMethod]
		public void AverageIfWithArraySingleCell()
		{
			_worksheet.Cells[2, 2].Value = 1;
			_worksheet.Cells[2, 3].Formula = "{1,2,3}";
			_worksheet.Cells[3, 3].Formula = "AVERAGEIF(C2,{1,2,3},B2)";
			_worksheet.Cells[3, 3].Calculate();
			Assert.AreEqual(1d, _worksheet.Cells[3, 3].Value);
		}

		[TestMethod]
		public void AverageIfWithArrayMultiCell()
		{
			_worksheet.Cells[2, 2].Value = 1;
			_worksheet.Cells[2, 3].Value = 1;
			_worksheet.Cells[2, 4].Value = 1;
			_worksheet.Cells[3, 2].Formula = "{1,2,3}";
			_worksheet.Cells[3, 3].Formula = "{1,2,3}";
			_worksheet.Cells[3, 4].Formula = "{1,2,3}";
			_worksheet.Cells[4, 4].Formula = "AVERAGEIF(B3:D3,{1,2,3},B2:D2)";
			_worksheet.Cells[4, 4].Calculate();
			Assert.AreEqual(1d, _worksheet.Cells[4, 4].Value);
		}
		#endregion
	}
}
