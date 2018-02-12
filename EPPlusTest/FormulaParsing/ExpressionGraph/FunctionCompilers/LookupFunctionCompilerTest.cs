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
			var parsingContext = ParsingContext.Create();
			parsingContext.Scopes.NewScope(RangeAddress.Empty);
			using (var excelPackage = new ExcelPackage())
			{
				var worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["C7"].Value = 1;
				worksheet.Cells["C8"].Value = 2;
				var provider = MockRepository.GenerateStub<EpplusExcelDataProvider>(excelPackage);
				provider.Stub(x => x.GetCellValue(null, 1, 1)).Return(1);
				provider.Stub(x => x.GetCellValue(null, 1, 2)).Return(1);
				provider.Stub(x => x.GetCellValue(null, 2, 1)).Return(2);
				provider.Stub(x => x.GetCellValue(null, 2, 2)).Return(5);
				provider.Stub(x => x.GetRange(null, 0, 0, "C8")).Return(new RangeInfo(worksheet, 8, 3, 8, 3));
				provider.Stub(x => x.GetRange(null, 0, 0, "C7")).Return(new RangeInfo(worksheet, 7, 3, 7, 3));
				parsingContext.ExcelDataProvider = provider;
				var tokens = SourceCodeTokenizer.Default.Tokenize("VLOOKUP(C8, A1:B2, C7 + 1)", worksheet.Name);
				var graph = new ExpressionGraphBuilder(provider, parsingContext).Build(tokens);
				var compileResult = new ExpressionCompiler().Compile(graph.Expressions);
				Assert.AreEqual(5, compileResult.ResultValue);
			}
		}
	}
}
