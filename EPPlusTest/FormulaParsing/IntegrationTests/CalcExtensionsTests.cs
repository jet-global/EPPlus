using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;

namespace EPPlusTest.FormulaParsing.IntegrationTests
{
	[TestClass]
	public class CalcExtensionsTests
	{
		#region Calculate Test Methods
		[TestMethod]
		public void ShouldCalculateChainTest()
		{
			var package = new ExcelPackage(new FileInfo("c:\\temp\\chaintest.xlsx"));
			package.Workbook.Calculate();
		}

		[TestMethod]
		public void CalculateTest()
		{
			var pck = new ExcelPackage();
			var ws = pck.Workbook.Worksheets.Add("Calc1");

			ws.SetValue("A1", (short)1);
			var v = pck.Workbook.FormulaParserManager.Parse("2.5-Calc1!A1+ABS(-3.0)-SIN(3)*abs(5)");
			Assert.AreEqual(3.79439996, Math.Round((double)v, 9));
		}

		[TestMethod]
		public void CalculateTest2()
		{
			var pck = new ExcelPackage();
			var ws = pck.Workbook.Worksheets.Add("Calc1");

			ws.SetValue("A1", (short)1);
			var v = pck.Workbook.FormulaParserManager.Parse("3*(2+5.5*2)+2*0.5+3");
			Assert.AreEqual(43, Math.Round((double)v, 9));
		}

		[TestMethod]
		public void CalculateWithSetStyle()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				var calcOption = new ExcelCalculationOption();
				var range = worksheet.Cells[2, 2];
				range.Formula = "TODAY()";
				range.Calculate(calcOption, true);
				Assert.AreEqual(14, range.Style.Numberformat.NumFmtID);

				range.Formula = "NOW()";
				range.Calculate(calcOption, true);
				Assert.AreEqual(14, range.Style.Numberformat.NumFmtID);

				range.Formula = "TODAY() + 5";
				range.Calculate(calcOption, true);
				Assert.AreEqual(14, range.Style.Numberformat.NumFmtID);
				Assert.AreEqual(DateTime.Today.AddDays(5), range.Value);

				range.Formula = "TIME(14, 23, 15)";
				range.Calculate(calcOption, true);
				Assert.AreEqual(21, range.Style.Numberformat.NumFmtID);

				range.Formula = "1 + 2";
				range.Calculate(calcOption, true);
				Assert.AreEqual(1, range.Style.Numberformat.NumFmtID);

				range.Formula = "1.5 * 2.3";
				range.Calculate(calcOption, true);
				Assert.AreEqual(2, range.Style.Numberformat.NumFmtID);

				range.Formula = @"""some""&"" text""";
				range.Calculate(calcOption, true);
				Assert.AreEqual(2, range.Style.Numberformat.NumFmtID);
			}
		}
		#endregion
	}
}
