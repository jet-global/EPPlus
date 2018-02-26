using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using Rhino.Mocks;

namespace EPPlusTest.FormulaParsing.LexicalAnalysis
{
	[TestClass]
	public class TokenFactoryTests
	{
		#region Class Variables
		private ITokenFactory _tokenFactory;
		private INameValueProvider _nameValueProvider;
		#endregion

		#region Test Setup
		[TestInitialize]
		public void Setup()
		{
			var context = ParsingContext.Create();
			var excelDataProvider = MockRepository.GenerateStub<ExcelDataProvider>();
			_nameValueProvider = MockRepository.GenerateStub<INameValueProvider>();
			_tokenFactory = new TokenFactory(context.Configuration.FunctionRepository, _nameValueProvider);
		}
		#endregion

		#region Test Methods
		[TestMethod]
		public void ShouldCreateAStringToken()
		{
			var input = "\"";
			var token = _tokenFactory.Create(Enumerable.Empty<Token>(), input);

			Assert.AreEqual("\"", token.Value);
			Assert.AreEqual(TokenType.String, token.TokenType);
		}

		[TestMethod]
		public void ShouldCreatePlusAsOperatorToken()
		{
			var input = "+";
			var token = _tokenFactory.Create(Enumerable.Empty<Token>(), input);

			Assert.AreEqual("+", token.Value);
			Assert.AreEqual(TokenType.Operator, token.TokenType);
		}

		[TestMethod]
		public void ShouldCreateMinusAsOperatorToken()
		{
			var input = "-";
			var token = _tokenFactory.Create(Enumerable.Empty<Token>(), input);

			Assert.AreEqual("-", token.Value);
			Assert.AreEqual(TokenType.Operator, token.TokenType);
		}

		[TestMethod]
		public void ShouldCreateMultiplyAsOperatorToken()
		{
			var input = "*";
			var token = _tokenFactory.Create(Enumerable.Empty<Token>(), input);

			Assert.AreEqual("*", token.Value);
			Assert.AreEqual(TokenType.Operator, token.TokenType);
		}

		[TestMethod]
		public void ShouldCreateDivideAsOperatorToken()
		{
			var input = "/";
			var token = _tokenFactory.Create(Enumerable.Empty<Token>(), input);

			Assert.AreEqual("/", token.Value);
			Assert.AreEqual(TokenType.Operator, token.TokenType);
		}

		[TestMethod]
		public void ShouldCreateEqualsAsOperatorToken()
		{
			var input = "=";
			var token = _tokenFactory.Create(Enumerable.Empty<Token>(), input);

			Assert.AreEqual("=", token.Value);
			Assert.AreEqual(TokenType.Operator, token.TokenType);
		}

		[TestMethod]
		public void ShouldCreateIntegerAsIntegerToken()
		{
			var input = "23";
			var token = _tokenFactory.Create(Enumerable.Empty<Token>(), input);

			Assert.AreEqual("23", token.Value);
			Assert.AreEqual(TokenType.Integer, token.TokenType);
		}

		[TestMethod]
		public void ShouldCreateBooleanAsBooleanToken()
		{
			var input = "true";
			var token = _tokenFactory.Create(Enumerable.Empty<Token>(), input);

			Assert.AreEqual("true", token.Value);
			Assert.AreEqual(TokenType.Boolean, token.TokenType);
		}

		[TestMethod]
		public void ShouldCreateDecimalAsDecimalToken()
		{
			var input = "23.3";
			var token = _tokenFactory.Create(Enumerable.Empty<Token>(), input);

			Assert.AreEqual("23.3", token.Value);
			Assert.AreEqual(TokenType.Decimal, token.TokenType);
		}

		[TestMethod]
		public void CreateShouldReadFunctionsFromFuncRepository()
		{
			var input = "Text";
			var token = _tokenFactory.Create(Enumerable.Empty<Token>(), input);
			Assert.AreEqual(TokenType.Function, token.TokenType);
			Assert.AreEqual("Text", token.Value);
		}

		[TestMethod]
		public void CreateShouldCreateExcelAddressAsExcelAddressToken()
		{
			var input = "A1";
			var token = _tokenFactory.Create(Enumerable.Empty<Token>(), input);
			Assert.AreEqual(TokenType.ExcelAddress, token.TokenType);
			Assert.AreEqual("A1", token.Value);
		}

		[TestMethod]
		public void CreateShouldCreateExcelRangeAsExcelAddressToken()
		{
			var input = "A1:B15";
			var token = _tokenFactory.Create(Enumerable.Empty<Token>(), input);
			Assert.AreEqual(TokenType.ExcelAddress, token.TokenType);
			Assert.AreEqual("A1:B15", token.Value);
		}

		[TestMethod]
		public void CreateShouldCreateExcelRangeOnOtherSheetAsExcelAddressToken()
		{
			var input = "ws!A1:B15";
			var token = _tokenFactory.Create(Enumerable.Empty<Token>(), input);
			Assert.AreEqual(TokenType.ExcelAddress, token.TokenType);
			Assert.AreEqual("ws!A1:B15", token.Value);
		}

		[TestMethod]
		public void CreateShouldCreateNamedValueAsExcelAddressToken()
		{
			var input = "NamedValue";
			_nameValueProvider.Stub(x => x.IsNamedValue("NamedValue", "")).Return(true);
			_nameValueProvider.Stub(x => x.IsNamedValue("NamedValue", null)).Return(true);
			var token = _tokenFactory.Create(Enumerable.Empty<Token>(), input);
			Assert.AreEqual(TokenType.NameValue, token.TokenType);
			Assert.AreEqual("NamedValue", token.Value);
		}

		[TestMethod]
		public void CreateShouldCreateExternalWorkbookReferenceAsInvalidReference()
		{
			var input = "[1]ws!A1:B15";
			var token = _tokenFactory.Create(Enumerable.Empty<Token>(), input);
			Assert.AreEqual(TokenType.InvalidReference, token.TokenType);
			Assert.AreEqual("[1]ws!A1:B15", token.Value);
		}

		[TestMethod]
		public void ShouldCreateStructuredReferenceTokens()
		{
			var structuredReferences = new[]
			{
				"MyTable[[#All],[MyColumn]]", // Casing is unimportant for item specifiers
				"MyTable[[#ALL],[MyColumn]]",
				"MyTable[[#Data],[MyColumn]]",
				"MyTable[[#DATA],[MyColumn]]",
				"MyTable[[#Headers],[MyColumn]]",
				"MyTable[[#HEADERS],[MyColumn]]",
				"MyTable[[#Totals],[MyColumn]]",
				"MyTable[[#TOTALS],[MyColumn]]",
				"MyTable[[#This Row],[MyColumn]]",
				"MyTable[[#THIS ROW],[MyColumn]]",
				"MyTable[[#Headers],[#Data],[MyColumn]]",
				"MyTable[[#All],[#Totals],[#This Row],[#Headers],[#Data],[MyColumn]]", // multiple item specifiers
				"MyTable[[#Headers],[MyStartColumn]:[MyEndColumn]]", // multi-column selector
				"MyTable[MyColumn]",
				@"\MyTable[MyColumn]", // Tables can begin with \
				"_MyTable[MyColumn]", // Tables can begin with _
				"My.Table[MyColumn]", // Tables can contain .
				"MyTable[[MyColumn]]", // Columns can be double bracketed
				"MyTable[[My \t Column]]", // Column names with \t MUST be double bracketed
				"MyTable[[My \n Column]]", // Column names with \n MUST be double bracketed
				"MyTable[[My \r Column]]", // Column names with \r MUST be double bracketed
				"MyTable[[My , Column]]", // Column names with , MUST be double bracketed
				"MyTable[[My : Column]]", // Column names with : MUST be double bracketed
				"MyTable[[My . Column]]", // Column names with . MUST be double bracketed
				"MyTable[[My '[ Column]]", // Column names with [ MUST be double bracketed AND [ must be escaped with '
				"MyTable[[My '] Column]]", // Column names with ] MUST be double bracketed AND ] must be escaped with '
				"MyTable[[My '# Column]]", // Column names with # MUST be double bracketed AND # must be escaped with '
				"MyTable[['# MyColumn]]", // Column names with # MUST be double bracketed AND # must be escaped with '
				"MyTable[[My '' Column]]", // Column names with ' MUST be double bracketed AND ' must be escaped with '
				"MyTable[[My \" Column]]", // Column names with ' MUST be double bracketed
				"MyTable[[My { Column]]", // Column names with { MUST be double bracketed
				"MyTable[[My } Column]]", // Column names with } MUST be double bracketed
				"MyTable[[My $ Column]]", // Column names with $ MUST be double bracketed
				"MyTable[[My ^ Column]]", // Column names with ^ MUST be double bracketed
				"MyTable[[My & Column]]", // Column names with & MUST be double bracketed
				"MyTable[[My * Column]]", // Column names with * MUST be double bracketed
				"MyTable[[My + Column]]", // Column names with + MUST be double bracketed
				"MyTable[[My = Column]]", // Column names with = MUST be double bracketed
				"MyTable[[My - Column]]", // Column names with - MUST be double bracketed
				"MyTable[[My > Column]]", // Column names with > MUST be double bracketed
				"MyTable[[My < Column]]", // Column names with < MUST be double bracketed
				"MyTable[[My / Column]]", // Column names with / MUST be double bracketed
				"MyTable[   [MyColumn]   ]", // whitespace can generally be ignored
			};
			foreach (var reference in structuredReferences)
			{
				var token = _tokenFactory.Create(Enumerable.Empty<Token>(), reference);
				Assert.AreEqual(TokenType.StructuredReference, token.TokenType, $"Reference: {reference} did not tokenize correctly.");
			}
		}
		#endregion
	}
}
