using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using Rhino.Mocks;
using static OfficeOpenXml.FormulaParsing.EpplusExcelDataProvider;

namespace EPPlusTest.FormulaParsing.ExpressionGraph.FunctionCompilers
{
	[TestClass]
	public class LookupFunctionCompilerTest
	{
		[TestMethod]
		public void VLookupCompiler()
		{
			using (var excelPackage = new ExcelPackage())
			{
				var worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["C7"].Value = 1;
				worksheet.Cells["C8"].Value = 2;
				worksheet.Cells["A1"].Value = 1;
				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["A2"].Value = 2;
				worksheet.Cells["B2"].Value = 5;
				worksheet.Cells["C3"].Formula = "VLOOKUP(C8, A1:B2, C7 + 1)";
				worksheet.Cells["C3"].Calculate();
				Assert.AreEqual(5, worksheet.Cells["C3"].Value);
			}
		}
	}
}
