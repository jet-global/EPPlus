using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace EPPlusTest.FormulaParsing.ExpressionGraph
{
	[TestClass]
	public class StructuredReferenceExpressionTest
	{
		#region Constants
		private const string TableName = "MyTable";
		private const string Header1 = "Header1";
		private const string Header2 = "Header2";
		private const string Header3 = "Header3";
		private const string Header4 = "Header4";
		#endregion

		#region Class Variables
		private ExcelPackage myPackage;
		private EpplusExcelDataProvider myDataProvider;
		private RangeAddressFactory myAddressFactory;
		#endregion

		#region Properties
		private ExcelPackage Package
		{
			get
			{
				if (myPackage == null)
				{
					myPackage = new ExcelPackage();
					var worksheet = myPackage.Workbook.Worksheets.Add("Sheet1");
					worksheet.Cells[3, 3].Value = StructuredReferenceExpressionTest.Header1;
					worksheet.Cells[3, 4].Value = StructuredReferenceExpressionTest.Header2;
					worksheet.Cells[3, 5].Value = StructuredReferenceExpressionTest.Header3;
					worksheet.Cells[3, 6].Value = StructuredReferenceExpressionTest.Header4;
					worksheet.Cells[4, 3].Value = "h1_r1";
					worksheet.Cells[4, 4].Value = "h2_r1";
					worksheet.Cells[4, 5].Value = "h3_r1";
					worksheet.Cells[4, 6].Value = "h4_r1";
					worksheet.Cells[5, 3].Value = "h1_r2";
					worksheet.Cells[5, 4].Value = "h2_r2";
					worksheet.Cells[5, 5].Value = "h3_r2";
					worksheet.Cells[5, 6].Value = "h4_r2";
					worksheet.Cells[6, 3].Value = "h1_r3";
					worksheet.Cells[6, 4].Value = "h2_r3";
					worksheet.Cells[6, 5].Value = "h3_r3";
					worksheet.Cells[6, 6].Value = "h4_r3";
					worksheet.Cells[7, 3].Value = "h1_r4";
					worksheet.Cells[7, 4].Value = "h2_r4";
					worksheet.Cells[7, 5].Value = "h3_r4";
					worksheet.Cells[7, 6].Value = "h4_r4";
					worksheet.Cells[8, 3].Value = 1;
					worksheet.Cells[8, 4].Value = 2;
					worksheet.Cells[8, 5].Value = 3;
					worksheet.Cells[8, 6].Value = 4;
					worksheet.Cells[9, 3].Value = "h1_r6";
					worksheet.Cells[9, 4].Value = "h2_r6";
					worksheet.Cells[9, 5].Value = "h3_r6";
					worksheet.Cells[9, 6].Value = "h4_r6";
					var table = worksheet.Tables.Add(new ExcelAddress("Sheet1", 3, 3, 9, 6), StructuredReferenceExpressionTest.TableName);
				}
				return myPackage;
			}
		}

		private EpplusExcelDataProvider DataProvider
		{
			get
			{
				if (myDataProvider == null)
					myDataProvider = new EpplusExcelDataProvider(this.Package);
				return myDataProvider;
			}
		}

		private RangeAddressFactory AddressFactory
		{
			get
			{
				if (myAddressFactory == null)
					myAddressFactory = new RangeAddressFactory(this.DataProvider);
				return myAddressFactory;
			}
		}
		#endregion

		#region Compile Tests
		[TestMethod]
		public void CompileSingleCellReferenceDoNotNegate()
		{
			var structuredReference = $"{StructuredReferenceExpressionTest.TableName}[[#This Row],[{StructuredReferenceExpressionTest.Header1}]]";
			// NOTE:: RangeAddressFactory.Create() takes column first, then row
			var origin = this.AddressFactory.Create("Sheet1", 7, 8);
			var parsingContext = ParsingContext.Create();
			using (var scope = parsingContext.Scopes.NewScope(origin))
			{
				var expression = new StructuredReferenceExpression(structuredReference, this.DataProvider, parsingContext, false);
				var result = expression.Compile();
				Assert.AreEqual(1, result.Result);
			}
		}

		[TestMethod]
		public void CompileSingleCellReferenceNegate()
		{
			var structuredReference = $"{StructuredReferenceExpressionTest.TableName}[[#This Row],[{StructuredReferenceExpressionTest.Header1}]]";
			// NOTE:: RangeAddressFactory.Create() takes column first, then row
			var origin = this.AddressFactory.Create("Sheet1", 7, 8);
			var parsingContext = ParsingContext.Create();
			using (var scope = parsingContext.Scopes.NewScope(origin))
			{
				var expression = new StructuredReferenceExpression(structuredReference, this.DataProvider, parsingContext, true);
				var result = expression.Compile();
				Assert.AreEqual(-1d, result.Result);
			}
		}

		[TestMethod]
		public void CompileSingleCellReferenceNegateNonNumericIsValueError()
		{
			var structuredReference = $"{StructuredReferenceExpressionTest.TableName}[[#This Row],[{StructuredReferenceExpressionTest.Header1}]]";
			// NOTE:: RangeAddressFactory.Create() takes column first, then row
			var origin = this.AddressFactory.Create("Sheet1", 7, 7);
			var parsingContext = ParsingContext.Create();
			using (var scope = parsingContext.Scopes.NewScope(origin))
			{
				var expression = new StructuredReferenceExpression(structuredReference, this.DataProvider, parsingContext, true);
				var result = expression.Compile();
				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), result.Result);
			}
		}

		[TestMethod]
		public void CompileSingleCellReferenceSingleCellResolveAsRange()
		{
			var structuredReference = $"{StructuredReferenceExpressionTest.TableName}[[#This Row],[{StructuredReferenceExpressionTest.Header1}]]";
			// NOTE:: RangeAddressFactory.Create() takes column first, then row
			var origin = this.AddressFactory.Create("Sheet1", 7, 8);
			var parsingContext = ParsingContext.Create();
			using (var scope = parsingContext.Scopes.NewScope(origin))
			{
				var expression = new StructuredReferenceExpression(structuredReference, this.DataProvider, parsingContext, false);
				expression.ResolveAsRange = true;
				var result = expression.Compile();
				var resultRange = result.Result as ExcelDataProvider.IRangeInfo;
				Assert.AreEqual(8, resultRange.Address.Start.Row);
				Assert.AreEqual(3, resultRange.Address.Start.Column);
				Assert.AreEqual(8, resultRange.Address.End.Row);
				Assert.AreEqual(3, resultRange.Address.End.Column);
				Assert.AreEqual(1, resultRange.GetOffset(0, 0));
			}
		}

		[TestMethod]
		public void CompileSingleCellReferenceMultiCellReferenceOutsideOfRow()
		{
			var structuredReference = $"{StructuredReferenceExpressionTest.TableName}[[#Data],[{StructuredReferenceExpressionTest.Header3}]:[{StructuredReferenceExpressionTest.Header4}]]";
			// NOTE:: RangeAddressFactory.Create() takes column first, then row
			var origin = this.AddressFactory.Create("Sheet1", 7, 15);
			var parsingContext = ParsingContext.Create();
			using (var scope = parsingContext.Scopes.NewScope(origin))
			{
				var expression = new StructuredReferenceExpression(structuredReference, this.DataProvider, parsingContext, false);
				var result = expression.Compile();
				var resultRange = result.Result as ExcelDataProvider.IRangeInfo;
				Assert.AreEqual(4, resultRange.Address.Start.Row);
				Assert.AreEqual(5, resultRange.Address.Start.Column);
				Assert.AreEqual(9, resultRange.Address.End.Row);
				Assert.AreEqual(6, resultRange.Address.End.Column);
				Assert.AreEqual("h3_r1", resultRange.GetOffset(0, 0));
				Assert.AreEqual("h3_r2", resultRange.GetOffset(1, 0));
				Assert.AreEqual("h3_r3", resultRange.GetOffset(2, 0));
				Assert.AreEqual("h3_r4", resultRange.GetOffset(3, 0));
				Assert.AreEqual(3, resultRange.GetOffset(4, 0));
				Assert.AreEqual("h3_r6", resultRange.GetOffset(5, 0));
				Assert.AreEqual("h4_r1", resultRange.GetOffset(0, 1));
				Assert.AreEqual("h4_r2", resultRange.GetOffset(1, 1));
				Assert.AreEqual("h4_r3", resultRange.GetOffset(2, 1));
				Assert.AreEqual("h4_r4", resultRange.GetOffset(3, 1));
				Assert.AreEqual(4, resultRange.GetOffset(4, 1));
				Assert.AreEqual("h4_r6", resultRange.GetOffset(5, 1));
			}
		}
		#endregion
	}
}
