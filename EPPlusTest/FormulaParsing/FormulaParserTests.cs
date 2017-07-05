using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using Rhino.Mocks;
using ExGraph = OfficeOpenXml.FormulaParsing.ExpressionGraph.ExpressionGraph;

namespace EPPlusTest.FormulaParsing
{
	[TestClass]
	public class FormulaParserTests
	{
		private FormulaParser _parser;

		[TestInitialize]
		public void Setup()
		{
			var provider = MockRepository.GenerateStub<ExcelDataProvider>();
			_parser = new FormulaParser(provider);

		}

		[TestCleanup]
		public void Cleanup()
		{

		}

		[TestMethod]
		public void ParserShouldCallLexer()
		{
			var lexer = MockRepository.GenerateStub<ILexer>();
			lexer.Stub(x => x.Tokenize("ABC")).Return(Enumerable.Empty<Token>());
			_parser.Configure(x => x.SetLexer(lexer));

			_parser.Parse("ABC");

			lexer.AssertWasCalled(x => x.Tokenize("ABC"));
		}

		[TestMethod]
		public void ParserShouldCallGraphBuilder()
		{
			var lexer = MockRepository.GenerateStub<ILexer>();
			var tokens = new List<Token>();
			lexer.Stub(x => x.Tokenize("ABC")).Return(tokens);
			var graphBuilder = MockRepository.GenerateStub<IExpressionGraphBuilder>();
			graphBuilder.Stub(x => x.Build(tokens)).Return(new ExGraph());

			_parser.Configure(config =>
				 {
					 config
							  .SetLexer(lexer)
							  .SetGraphBuilder(graphBuilder);
				 });

			_parser.Parse("ABC");

			graphBuilder.AssertWasCalled(x => x.Build(tokens));
		}

		[TestMethod]
		public void ParserShouldCallCompiler()
		{
			var lexer = MockRepository.GenerateStub<ILexer>();
			var tokens = new List<Token>();
			lexer.Stub(x => x.Tokenize("ABC")).Return(tokens);
			var expectedGraph = new ExGraph();
			expectedGraph.Add(new StringExpression("asdf"));
			var graphBuilder = MockRepository.GenerateStub<IExpressionGraphBuilder>();
			graphBuilder.Stub(x => x.Build(tokens)).Return(expectedGraph);
			var compiler = MockRepository.GenerateStub<IExpressionCompiler>();
			compiler.Stub(x => x.Compile(expectedGraph.Expressions)).Return(new CompileResult(0, DataType.Integer));

			_parser.Configure(config =>
			{
				config
						 .SetLexer(lexer)
						 .SetGraphBuilder(graphBuilder)
						 .SetExpresionCompiler(compiler);
			});

			_parser.Parse("ABC");

			compiler.AssertWasCalled(x => x.Compile(expectedGraph.Expressions));
		}

		[TestMethod]
		public void ParseAtShouldCallExcelDataProvider()
		{
			var excelDataProvider = MockRepository.GenerateStub<ExcelDataProvider>();
			excelDataProvider
				 .Stub(x => x.GetRangeFormula(string.Empty, 1, 1))
				 .Return("Sum(1,2)");
			var parser = new FormulaParser(excelDataProvider);
			var result = parser.ParseAt("A1");
			Assert.AreEqual(3d, result);
		}

		[TestMethod, ExpectedException(typeof(ArgumentException))]
		public void ParseAtShouldThrowIfAddressIsNull()
		{
			_parser.ParseAt(null);
		}

		//Tests involving formulas with cell ranges, not cell references.

		[TestMethod]
		public void ParseWithCellRangeReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["D2"].Formula = "B1:B3";
				worksheet.Cells["AA2"].Formula = "B1:B3";
				worksheet.Calculate();
				Assert.AreEqual(2, worksheet.Cells["D2"].Value);
				Assert.AreEqual(2, worksheet.Cells["AA2"].Value);
			}
		}

		[TestMethod]
		public void ParseWithSingleCellReferenceReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["D2"].Formula = "B1";
				worksheet.Calculate();
				Assert.AreEqual(1, worksheet.Cells["D2"].Value);
			}
		}

		[TestMethod]
		public void ParseWithCellReferenceNotInRowReturnsPoundValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["D3"].Formula = "B1:B2";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)worksheet.Cells["D3"].Value).Type);
			}
		}

		[TestMethod]
		public void ParseWithCellReferenceInRowReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["D1"].Formula = "B1:B2";
				worksheet.Cells["D2"].Formula = "B1:B2";
				worksheet.Cells["D3"].Formula = "B1:B3";
				worksheet.Calculate();
				Assert.AreEqual(1, worksheet.Cells["D1"].Value);
				Assert.AreEqual(2, worksheet.Cells["D2"].Value);
				Assert.AreEqual(3, worksheet.Cells["D3"].Value);
			}
		}

		[TestMethod]
		public void ParseWithSameRowDifferentColumnsReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["J6"].Value = 1;
				worksheet.Cells["K6"].Value = 2;
				worksheet.Cells["L6"].Value = 3;
				worksheet.Cells["K8"].Formula = "J6:L6";
				worksheet.Cells["J8"].Formula = "J6:L6";
				worksheet.Cells["L8"].Formula = "J6:L6";
				worksheet.Calculate();
				Assert.AreEqual(1, worksheet.Cells["J8"].Value);
				Assert.AreEqual(2, worksheet.Cells["K8"].Value);
				Assert.AreEqual(3, worksheet.Cells["L8"].Value);
			}
		}
	}
}
