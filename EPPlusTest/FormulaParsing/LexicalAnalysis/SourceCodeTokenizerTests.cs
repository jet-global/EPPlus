using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace EPPlusTest.FormulaParsing.LexicalAnalysis
{
	[TestClass]
	public class SourceCodeTokenizerTests
	{
		#region Class Variables
		private SourceCodeTokenizer _tokenizer;
		#endregion

		#region Test Setup
		[TestInitialize]
		public void Setup()
		{
			var context = ParsingContext.Create();
			_tokenizer = new SourceCodeTokenizer(context.Configuration.FunctionRepository, null);
		}
		#endregion
		
		#region Tokenize Tests
		[TestMethod]
		public void ShouldCreateTokensForStringCorrectly()
		{
			var input = "\"abc123\"";
			var tokens = _tokenizer.Tokenize(input);

			Assert.AreEqual(3, tokens.Count());
			Assert.AreEqual(TokenType.String, tokens.First().TokenType);
			Assert.AreEqual(TokenType.StringContent, tokens.ElementAt(1).TokenType);
			Assert.AreEqual(TokenType.String, tokens.Last().TokenType);
		}

		[TestMethod]
		public void ShouldTokenizeStringCorrectly()
		{
			var input = "\"ab(c)d\"";
			var tokens = _tokenizer.Tokenize(input);

			Assert.AreEqual(3, tokens.Count());
		}

		[TestMethod]
		public void ShouldHandleWhitespaceCorrectly()
		{
			var input = @"""          """;
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(3, tokens.Count());
			Assert.AreEqual(TokenType.StringContent, tokens.ElementAt(1).TokenType);
			Assert.AreEqual(10, tokens.ElementAt(1).Value.Length);
		}

		[TestMethod]
		public void ShouldCreateTokensForFunctionCorrectly()
		{
			var input = "Text(2)";
			var tokens = _tokenizer.Tokenize(input);

			Assert.AreEqual(4, tokens.Count());
			Assert.AreEqual(TokenType.Function, tokens.First().TokenType);
			Assert.AreEqual(TokenType.OpeningParenthesis, tokens.ElementAt(1).TokenType);
			Assert.AreEqual(TokenType.Integer, tokens.ElementAt(2).TokenType);
			Assert.AreEqual("2", tokens.ElementAt(2).Value);
			Assert.AreEqual(TokenType.ClosingParenthesis, tokens.Last().TokenType);
		}

		[TestMethod]
		public void ShouldHandleMultipleCharOperatorCorrectly()
		{
			var input = "1 <= 2";
			var tokens = _tokenizer.Tokenize(input);

			Assert.AreEqual(3, tokens.Count());
			Assert.AreEqual("<=", tokens.ElementAt(1).Value);
			Assert.AreEqual(TokenType.Operator, tokens.ElementAt(1).TokenType);
		}

		[TestMethod]
		public void ShouldCreateTokensForEnumerableCorrectly()
		{
			var input = "Text({1;2})";
			var tokens = _tokenizer.Tokenize(input);

			Assert.AreEqual(8, tokens.Count());
			Assert.AreEqual(TokenType.OpeningEnumerable, tokens.ElementAt(2).TokenType);
			Assert.AreEqual(TokenType.ClosingEnumerable, tokens.ElementAt(6).TokenType);
		}

		[TestMethod]
		public void ShouldCreateTokensForExcelAddressCorrectly()
		{
			var input = "Text(A1)";
			var tokens = _tokenizer.Tokenize(input);

			Assert.AreEqual(TokenType.ExcelAddress, tokens.ElementAt(2).TokenType);
		}

		[TestMethod]
		public void ShouldCreateTokenForPercentAfterDecimal()
		{
			var input = "1,23%";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(TokenType.Percent, tokens.Last().TokenType);
		}

		[TestMethod]
		public void ShouldIgnoreTwoSubsequentStringIdentifyers()
		{
			var input = "\"hello\"\"world\"";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(3, tokens.Count());
			Assert.AreEqual("hello\"world", tokens.ElementAt(1).Value);
		}

		[TestMethod]
		public void ShouldIgnoreTwoSubsequentStringIdentifyers2()
		{
			var input = "\"\"\"\"\"\"";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(TokenType.StringContent, tokens.ElementAt(1).TokenType);
		}

		[TestMethod]
		public void TokenizerShouldIgnoreOperatorInString()
		{
			var input = "\"*\"";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(TokenType.StringContent, tokens.ElementAt(1).TokenType);
		}

		[TestMethod]
		public void TokenizerShouldHandleWorksheetNameWithMinus()
		{
			var input = "'A-B'!A1";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(1, tokens.Count());
			Assert.AreEqual(TokenType.ExcelAddress, tokens.ElementAt(0).TokenType);
		}

		[TestMethod]
		public void TokenizeStripsLeadingPlusSign()
		{
			var input = @"+3-3";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(3, tokens.Count());
			Assert.AreEqual(TokenType.Integer, tokens.ElementAt(0).TokenType);
			Assert.AreEqual(TokenType.Operator, tokens.ElementAt(1).TokenType);
			Assert.AreEqual(TokenType.Integer, tokens.ElementAt(2).TokenType);
		}

		[TestMethod]
		public void TokenizeStripsLeadingDoubleNegator()
		{
			var input = @"--3-3";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(3, tokens.Count());
			Assert.AreEqual(TokenType.Integer, tokens.ElementAt(0).TokenType);
			Assert.AreEqual(TokenType.Operator, tokens.ElementAt(1).TokenType);
			Assert.AreEqual(TokenType.Integer, tokens.ElementAt(2).TokenType);
		}

		[TestMethod]
		public void TokenizeHandlesPositiveNegator()
		{
			var input = @"+-3-3";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(4, tokens.Count());
			Assert.AreEqual(TokenType.Negator, tokens.ElementAt(0).TokenType);
			Assert.AreEqual(TokenType.Integer, tokens.ElementAt(1).TokenType);
			Assert.AreEqual(TokenType.Operator, tokens.ElementAt(2).TokenType);
			Assert.AreEqual(TokenType.Integer, tokens.ElementAt(3).TokenType);
		}

		[TestMethod]
		public void TokenizeHandlesNegatorPositive()
		{
			var input = @"-+3-3";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(4, tokens.Count());
			Assert.AreEqual(TokenType.Negator, tokens.ElementAt(0).TokenType);
			Assert.AreEqual(TokenType.Integer, tokens.ElementAt(1).TokenType);
			Assert.AreEqual(TokenType.Operator, tokens.ElementAt(2).TokenType);
			Assert.AreEqual(TokenType.Integer, tokens.ElementAt(3).TokenType);
		}

		[TestMethod]
		public void TokenizeStripsLeadingPlusSignFromFirstFunctionArgument()
		{
			var input = @"SUM(+3-3,5)";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(8, tokens.Count());
			Assert.AreEqual(TokenType.Function, tokens.ElementAt(0).TokenType);
			Assert.AreEqual(TokenType.OpeningParenthesis, tokens.ElementAt(1).TokenType);
			Assert.AreEqual(TokenType.Integer, tokens.ElementAt(2).TokenType);
			Assert.AreEqual(TokenType.Operator, tokens.ElementAt(3).TokenType);
			Assert.AreEqual(TokenType.Integer, tokens.ElementAt(4).TokenType);
			Assert.AreEqual(TokenType.Comma, tokens.ElementAt(5).TokenType);
			Assert.AreEqual(TokenType.Integer, tokens.ElementAt(6).TokenType);
			Assert.AreEqual(TokenType.ClosingParenthesis, tokens.ElementAt(7).TokenType);
		}

		[TestMethod]
		public void TokenizeStripsLeadingPlusSignFromSecondFunctionArgument()
		{
			var input = @"SUM(5,+3-3)";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(8, tokens.Count());
			Assert.AreEqual(TokenType.Function, tokens.ElementAt(0).TokenType);
			Assert.AreEqual(TokenType.OpeningParenthesis, tokens.ElementAt(1).TokenType);
			Assert.AreEqual(TokenType.Integer, tokens.ElementAt(2).TokenType);
			Assert.AreEqual(TokenType.Comma, tokens.ElementAt(3).TokenType);
			Assert.AreEqual(TokenType.Integer, tokens.ElementAt(4).TokenType);
			Assert.AreEqual(TokenType.Operator, tokens.ElementAt(5).TokenType);
			Assert.AreEqual(TokenType.Integer, tokens.ElementAt(6).TokenType);
			Assert.AreEqual(TokenType.ClosingParenthesis, tokens.ElementAt(7).TokenType);
		}

		[TestMethod]
		public void TokenizeStripsLeadingDoubleNegatorFromFirstFunctionArgument()
		{
			var input = @"SUM(--3-3,5)";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(8, tokens.Count());
			Assert.AreEqual(TokenType.Function, tokens.ElementAt(0).TokenType);
			Assert.AreEqual(TokenType.OpeningParenthesis, tokens.ElementAt(1).TokenType);
			Assert.AreEqual(TokenType.Integer, tokens.ElementAt(2).TokenType);
			Assert.AreEqual(TokenType.Operator, tokens.ElementAt(3).TokenType);
			Assert.AreEqual(TokenType.Integer, tokens.ElementAt(4).TokenType);
			Assert.AreEqual(TokenType.Comma, tokens.ElementAt(5).TokenType);
			Assert.AreEqual(TokenType.Integer, tokens.ElementAt(6).TokenType);
			Assert.AreEqual(TokenType.ClosingParenthesis, tokens.ElementAt(7).TokenType);
		}

		[TestMethod]
		public void TokenizeStripsLeadingDoubleNegatorFromSecondFunctionArgument()
		{
			var input = @"SUM(5,--3-3)";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(8, tokens.Count());
			Assert.AreEqual(TokenType.Function, tokens.ElementAt(0).TokenType);
			Assert.AreEqual(TokenType.OpeningParenthesis, tokens.ElementAt(1).TokenType);
			Assert.AreEqual(TokenType.Integer, tokens.ElementAt(2).TokenType);
			Assert.AreEqual(TokenType.Comma, tokens.ElementAt(3).TokenType);
			Assert.AreEqual(TokenType.Integer, tokens.ElementAt(4).TokenType);
			Assert.AreEqual(TokenType.Operator, tokens.ElementAt(5).TokenType);
			Assert.AreEqual(TokenType.Integer, tokens.ElementAt(6).TokenType);
			Assert.AreEqual(TokenType.ClosingParenthesis, tokens.ElementAt(7).TokenType);
		}

		[TestMethod]
		public void TokenizeHandlesPositiveNegatorAsFirstFunctionArgument()
		{
			var input = @"SUM(+-3-3,5)";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(9, tokens.Count());
			Assert.AreEqual(TokenType.Function, tokens.ElementAt(0).TokenType);
			Assert.AreEqual(TokenType.OpeningParenthesis, tokens.ElementAt(1).TokenType);
			Assert.AreEqual(TokenType.Negator, tokens.ElementAt(2).TokenType);
			Assert.AreEqual(TokenType.Integer, tokens.ElementAt(3).TokenType);
			Assert.AreEqual(TokenType.Operator, tokens.ElementAt(4).TokenType);
			Assert.AreEqual(TokenType.Integer, tokens.ElementAt(5).TokenType);
			Assert.AreEqual(TokenType.Comma, tokens.ElementAt(6).TokenType);
			Assert.AreEqual(TokenType.Integer, tokens.ElementAt(7).TokenType);
			Assert.AreEqual(TokenType.ClosingParenthesis, tokens.ElementAt(8).TokenType);
		}

		[TestMethod]
		public void TokenizeHandlesNegatorPositiveAsFirstFunctionArgument()
		{
			var input = @"SUM(-+3-3,5)";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(9, tokens.Count());
			Assert.AreEqual(TokenType.Function, tokens.ElementAt(0).TokenType);
			Assert.AreEqual(TokenType.OpeningParenthesis, tokens.ElementAt(1).TokenType);
			Assert.AreEqual(TokenType.Negator, tokens.ElementAt(2).TokenType);
			Assert.AreEqual(TokenType.Integer, tokens.ElementAt(3).TokenType);
			Assert.AreEqual(TokenType.Operator, tokens.ElementAt(4).TokenType);
			Assert.AreEqual(TokenType.Integer, tokens.ElementAt(5).TokenType);
			Assert.AreEqual(TokenType.Comma, tokens.ElementAt(6).TokenType);
			Assert.AreEqual(TokenType.Integer, tokens.ElementAt(7).TokenType);
			Assert.AreEqual(TokenType.ClosingParenthesis, tokens.ElementAt(8).TokenType);
		}

		[TestMethod]
		public void TokenizeHandlesPositiveNegatorAsSecondFunctionArgument()
		{
			var input = @"SUM(5,+-3-3)";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(9, tokens.Count());
			Assert.AreEqual(TokenType.Function, tokens.ElementAt(0).TokenType);
			Assert.AreEqual(TokenType.OpeningParenthesis, tokens.ElementAt(1).TokenType);
			Assert.AreEqual(TokenType.Integer, tokens.ElementAt(2).TokenType);
			Assert.AreEqual(TokenType.Comma, tokens.ElementAt(3).TokenType);
			Assert.AreEqual(TokenType.Negator, tokens.ElementAt(4).TokenType);
			Assert.AreEqual(TokenType.Integer, tokens.ElementAt(5).TokenType);
			Assert.AreEqual(TokenType.Operator, tokens.ElementAt(6).TokenType);
			Assert.AreEqual(TokenType.Integer, tokens.ElementAt(7).TokenType);
			Assert.AreEqual(TokenType.ClosingParenthesis, tokens.ElementAt(8).TokenType);
		}

		[TestMethod]
		public void TokenizeHandlesNegatorPositiveAsSecondFunctionArgument()
		{
			var input = @"SUM(5,-+3-3)";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(9, tokens.Count());
			Assert.AreEqual(TokenType.Function, tokens.ElementAt(0).TokenType);
			Assert.AreEqual(TokenType.OpeningParenthesis, tokens.ElementAt(1).TokenType);
			Assert.AreEqual(TokenType.Integer, tokens.ElementAt(2).TokenType);
			Assert.AreEqual(TokenType.Comma, tokens.ElementAt(3).TokenType);
			Assert.AreEqual(TokenType.Negator, tokens.ElementAt(4).TokenType);
			Assert.AreEqual(TokenType.Integer, tokens.ElementAt(5).TokenType);
			Assert.AreEqual(TokenType.Operator, tokens.ElementAt(6).TokenType);
			Assert.AreEqual(TokenType.Integer, tokens.ElementAt(7).TokenType);
			Assert.AreEqual(TokenType.ClosingParenthesis, tokens.ElementAt(8).TokenType);
		}

		[TestMethod]
		public void TokenizeHandlesSingleStructuredReferences()
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
				var tokens = _tokenizer.Tokenize(reference);
				Assert.AreEqual(TokenType.StructuredReference, tokens.ElementAt(0).TokenType, $"Reference: {reference} did not tokenize correctly.");
			}
		}

		[TestMethod]
		public void TokenizeHandlesOperatorDelimitedStructuredReferences()
		{
			var formula = "MyTable[[#All],[MyColumn]]*MyTable[[#This Row],[MyColumn]]";
			var tokens = _tokenizer.Tokenize(formula);
			Assert.AreEqual(3, tokens.Count());
			Assert.AreEqual("MyTable[[#All],[MyColumn]]", tokens.ElementAt(0).Value);
			Assert.AreEqual(TokenType.StructuredReference, tokens.ElementAt(0).TokenType);
			Assert.AreEqual("*", tokens.ElementAt(1).Value);
			Assert.AreEqual(TokenType.Operator, tokens.ElementAt(1).TokenType);
			Assert.AreEqual("MyTable[[#This Row],[MyColumn]]", tokens.ElementAt(2).Value);
			Assert.AreEqual(TokenType.StructuredReference, tokens.ElementAt(2).TokenType);
		}

		[TestMethod]
		public void TokenizeHandlesStructuredReferencesAsFunctionArguments()
		{
			var formula = "SUM(MyTable[[#All],[MyColumn]],MyTable[[#This Row],[MyColumn]])";
			var tokens = _tokenizer.Tokenize(formula);
			Assert.AreEqual(6, tokens.Count());
			Assert.AreEqual("SUM", tokens.ElementAt(0).Value);
			Assert.AreEqual(TokenType.Function, tokens.ElementAt(0).TokenType);
			Assert.AreEqual("(", tokens.ElementAt(1).Value);
			Assert.AreEqual(TokenType.OpeningParenthesis, tokens.ElementAt(1).TokenType);
			Assert.AreEqual("MyTable[[#All],[MyColumn]]", tokens.ElementAt(2).Value);
			Assert.AreEqual(TokenType.StructuredReference, tokens.ElementAt(2).TokenType);
			Assert.AreEqual(",", tokens.ElementAt(3).Value);
			Assert.AreEqual(TokenType.Comma, tokens.ElementAt(3).TokenType);
			Assert.AreEqual("MyTable[[#This Row],[MyColumn]]", tokens.ElementAt(4).Value);
			Assert.AreEqual(TokenType.StructuredReference, tokens.ElementAt(4).TokenType);
			Assert.AreEqual(")", tokens.ElementAt(5).Value);
			Assert.AreEqual(TokenType.ClosingParenthesis, tokens.ElementAt(5).TokenType);
		}

		[TestMethod]
		public void TokenizeHandlesNegatedStructuredReferences()
		{
			var formula = "-MyTable[[#This Row],[MyColumn]]";
			var tokens = _tokenizer.Tokenize(formula);
			Assert.AreEqual(2, tokens.Count());
			Assert.AreEqual("-", tokens.ElementAt(0).Value);
			Assert.AreEqual(TokenType.Negator, tokens.ElementAt(0).TokenType);
			Assert.AreEqual("MyTable[[#This Row],[MyColumn]]", tokens.ElementAt(1).Value);
			Assert.AreEqual(TokenType.StructuredReference, tokens.ElementAt(1).TokenType);
		}

		#region Error Type Tests
		[TestMethod]
		public void TokenizeHandlesNotApplicableError()
		{
			string input = "#N/A";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(1, tokens.Count());
			var token = tokens.ElementAt(0);
			Assert.AreEqual(TokenType.NotApplicableError, token.TokenType);
			Assert.AreEqual("#N/A", token.Value);
		}

		[TestMethod]
		public void TokenizeHandlesNameError()
		{
			string input = "#NAME?";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(1, tokens.Count());
			var token = tokens.ElementAt(0);
			Assert.AreEqual(TokenType.NameError, token.TokenType);
			Assert.AreEqual("#NAME?", token.Value);
		}

		[TestMethod]
		public void TokenizeHandlesDivZeroError()
		{
			string input = "#DIV/0!";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(1, tokens.Count());
			var token = tokens.ElementAt(0);
			Assert.AreEqual(TokenType.DivideByZeroError, token.TokenType);
			Assert.AreEqual("#DIV/0!", token.Value);
		}

		[TestMethod]
		public void TokenizeHandlesNullError()
		{
			string input = "#NULL!";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(1, tokens.Count());
			var token = tokens.ElementAt(0);
			Assert.AreEqual(TokenType.Null, token.TokenType);
			Assert.AreEqual("#NULL!", token.Value);
		}

		[TestMethod]
		public void TokenizeHandlesValueError()
		{
			var input = "#VALUE!";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(1, tokens.Count());
			var token = tokens.ElementAt(0);
			Assert.AreEqual(TokenType.ValueDataTypeError, token.TokenType);
			Assert.AreEqual("#VALUE!", token.Value);
		}

		[TestMethod]
		public void TokenizeHandlesNumError()
		{
			var input = "#NUM!";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(1, tokens.Count());
			var token = tokens.ElementAt(0);
			Assert.AreEqual(TokenType.NumericError, token.TokenType);
			Assert.AreEqual("#NUM!", token.Value);
		}

		#region #REF! Address Tokenization Tests
		[TestMethod]
		public void TokenizeHandlesInvalidReferenceError()
		{
			string input = "#REF!";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(1, tokens.Count());
			var token = tokens.ElementAt(0);
			Assert.AreEqual(TokenType.InvalidReference, token.TokenType);
			Assert.AreEqual("#REF!", token.Value);
		}

		[TestMethod]
		public void TokenizeHandlesInvalidCellReferenceWithSheetName()
		{
			string input = "Sheet1!#REF!";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(1, tokens.Count());
			var token = tokens.ElementAt(0);
			Assert.AreEqual(TokenType.InvalidReference, token.TokenType);
			Assert.AreEqual("Sheet1!#REF!", token.Value);
		}

		[TestMethod]
		public void TokenizeHandlesInvalidCellReferenceWithQuotedSheetName()
		{
			string input = "'Sheet1'!#REF!";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(1, tokens.Count());
			var token = tokens.ElementAt(0);
			Assert.AreEqual(TokenType.InvalidReference, token.TokenType);
			Assert.AreEqual("'Sheet1'!#REF!", token.Value);
		}

		[TestMethod]
		public void TokenizeHandlesInvalidSheetReference()
		{
			string input = "#REF!C5";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(1, tokens.Count());
			var token = tokens.ElementAt(0);
			Assert.AreEqual(TokenType.InvalidReference, token.TokenType);
			Assert.AreEqual("#REF!C5", token.Value);
		}

		[TestMethod]
		public void TokenizeHandlesInvalidSheetReferenceRange()
		{
			string input = "#REF!C5:#REF!C6";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(1, tokens.Count());
			var token = tokens.ElementAt(0);
			Assert.AreEqual(TokenType.InvalidReference, token.TokenType);
			Assert.AreEqual("#REF!C5:#REF!C6", token.Value);
		}

		[TestMethod]
		public void TokenizeHandlesInvalidStartSheetReferenceRange()
		{
			string input = "#REF!C5:Sheet1!C6";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(1, tokens.Count());
			var token = tokens.ElementAt(0);
			Assert.AreEqual(TokenType.InvalidReference, token.TokenType);
			Assert.AreEqual("#REF!C5:Sheet1!C6", token.Value);
		}

		[TestMethod]
		public void TokenizeHandlesInvalidEndSheetReferenceRange()
		{
			string input = "Sheet1!C5:#REF!C6";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(1, tokens.Count());
			var token = tokens.ElementAt(0);
			Assert.AreEqual(TokenType.InvalidReference, token.TokenType);
			Assert.AreEqual("Sheet1!C5:#REF!C6", token.Value);
		}

		[TestMethod]
		public void TokenizeHandlesValidRefNamedSheetReference()
		{
			// A sheet named "#REF", which is valid.
			string input = "'#REF'!C5";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(1, tokens.Count());
			var token = tokens.ElementAt(0);
			Assert.AreEqual(TokenType.ExcelAddress, token.TokenType);
			Assert.AreEqual("'#REF'!C5", token.Value);
		}

		[TestMethod]
		public void TokenizeHandlesValidRefBangNamedSheetReference()
		{
			// A sheet named "#REF!", which is valid.
			string input = "'#REF!'!C5";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(1, tokens.Count());
			var token = tokens.ElementAt(0);
			Assert.AreEqual(TokenType.ExcelAddress, token.TokenType);
			Assert.AreEqual("'#REF!'!C5", token.Value);
		}

		[TestMethod]
		public void TokenizeHandlesStartInvalidCellReferenceRangeWithoutSheetName()
		{
			string input = "#REF!:C5";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(1, tokens.Count());
			var token = tokens.ElementAt(0);
			Assert.AreEqual(TokenType.InvalidReference, token.TokenType);
			Assert.AreEqual("#REF!:C5", token.Value);
		}

		[TestMethod]
		public void TokenizeHandlesEndInvalidCellReferenceRangeWithoutSheetName()
		{
			string input = "C5:#REF!";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(1, tokens.Count());
			var token = tokens.ElementAt(0);
			Assert.AreEqual(TokenType.InvalidReference, token.TokenType);
			Assert.AreEqual("C5:#REF!", token.Value);
		}

		[TestMethod]
		public void TokenizeHandlesBothInvalidCellReferenceRangeWithoutSheetName()
		{
			string input = "#REF!:#REF!";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(1, tokens.Count());
			var token = tokens.ElementAt(0);
			Assert.AreEqual(TokenType.InvalidReference, token.TokenType);
			Assert.AreEqual("#REF!:#REF!", token.Value);
		}

		[TestMethod]
		public void TokenizeHandlesStartInvalidCellReferenceRangeWithSheetName()
		{
			string input = "Sheet1!#REF!:Sheet1!C5";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(1, tokens.Count());
			var token = tokens.ElementAt(0);
			Assert.AreEqual(TokenType.InvalidReference, token.TokenType);
			Assert.AreEqual("Sheet1!#REF!:Sheet1!C5", token.Value);
		}

		[TestMethod]
		public void TokenizeHandlesEndInvalidCellReferenceRangeWithSheetName()
		{
			string input = "Sheet1!C5:Sheet1!#REF!";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(1, tokens.Count());
			var token = tokens.ElementAt(0);
			Assert.AreEqual(TokenType.InvalidReference, token.TokenType);
			Assert.AreEqual("Sheet1!C5:Sheet1!#REF!", token.Value);
		}

		[TestMethod]
		public void TokenizeHandlesBothInvalidCellReferenceRangeWithSheetName()
		{
			string input = "Sheet1!#REF!:Sheet1!#REF!";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(1, tokens.Count());
			var token = tokens.ElementAt(0);
			Assert.AreEqual(TokenType.InvalidReference, token.TokenType);
			Assert.AreEqual("Sheet1!#REF!:Sheet1!#REF!", token.Value);
		}

		[TestMethod]
		public void TokenizeHandlesStartInvalidCellReferenceRangeWithQuotedSheetName()
		{
			string input = "'Sheet1'!#REF!:'Sheet1'!C5";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(1, tokens.Count());
			var token = tokens.ElementAt(0);
			Assert.AreEqual(TokenType.InvalidReference, token.TokenType);
			Assert.AreEqual("'Sheet1'!#REF!:'Sheet1'!C5", token.Value);
		}

		[TestMethod]
		public void TokenizeHandlesEndInvalidCellReferenceRangeWithQuotedSheetName()
		{
			string input = "'Sheet1'!C5:'Sheet1'!#REF!";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(1, tokens.Count());
			var token = tokens.ElementAt(0);
			Assert.AreEqual(TokenType.InvalidReference, token.TokenType);
			Assert.AreEqual("'Sheet1'!C5:'Sheet1'!#REF!", token.Value);
		}

		[TestMethod]
		public void TokenizeHandlesBothInvalidCellReferenceRangeWithQuotedSheetName()
		{
			string input = "'Sheet1'!#REF!:'Sheet1'!#REF!";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(1, tokens.Count());
			var token = tokens.ElementAt(0);
			Assert.AreEqual(TokenType.InvalidReference, token.TokenType);
			Assert.AreEqual("'Sheet1'!#REF!:'Sheet1'!#REF!", token.Value);
		}

		[TestMethod]
		public void TokenizeHandlesStartInvalidCellReferenceRangeWithStartQuotedSheetName()
		{
			string input = "'Sheet1'!#REF!:Sheet1!C5";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(1, tokens.Count());
			var token = tokens.ElementAt(0);
			Assert.AreEqual(TokenType.InvalidReference, token.TokenType);
			Assert.AreEqual("'Sheet1'!#REF!:Sheet1!C5", token.Value);
		}

		[TestMethod]
		public void TokenizeHandlesEndInvalidCellReferenceRangeWithStartQuotedSheetName()
		{
			string input = "'Sheet1'!C5:Sheet1!#REF!";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(1, tokens.Count());
			var token = tokens.ElementAt(0);
			Assert.AreEqual(TokenType.InvalidReference, token.TokenType);
			Assert.AreEqual("'Sheet1'!C5:Sheet1!#REF!", token.Value);
		}

		[TestMethod]
		public void TokenizeHandlesBothInvalidCellReferenceRangeWithStartQuotedSheetName()
		{
			string input = "'Sheet1'!#REF!:Sheet1!#REF!";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(1, tokens.Count());
			var token = tokens.ElementAt(0);
			Assert.AreEqual(TokenType.InvalidReference, token.TokenType);
			Assert.AreEqual("'Sheet1'!#REF!:Sheet1!#REF!", token.Value);
		}

		[TestMethod]
		public void TokenizeHandlesStartInvalidCellReferenceRangeWithEndQuotedSheetName()
		{
			string input = "Sheet1!#REF!:'Sheet1'!C5";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(1, tokens.Count());
			var token = tokens.ElementAt(0);
			Assert.AreEqual(TokenType.InvalidReference, token.TokenType);
			Assert.AreEqual("Sheet1!#REF!:'Sheet1'!C5", token.Value);
		}

		[TestMethod]
		public void TokenizeHandlesEndInvalidCellReferenceRangeWithEndQuotedSheetName()
		{
			string input = "Sheet1!C5:'Sheet1'!#REF!";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(1, tokens.Count());
			var token = tokens.ElementAt(0);
			Assert.AreEqual(TokenType.InvalidReference, token.TokenType);
			Assert.AreEqual("Sheet1!C5:'Sheet1'!#REF!", token.Value);
		}

		[TestMethod]
		public void TokenizeHandlesBothInvalidCellReferenceRangeWithEndQuotedSheetName()
		{
			string input = "Sheet1!#REF!:'Sheet1'!#REF!";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(1, tokens.Count());
			var token = tokens.ElementAt(0);
			Assert.AreEqual(TokenType.InvalidReference, token.TokenType);
			Assert.AreEqual("Sheet1!#REF!:'Sheet1'!#REF!", token.Value);
		}

		[TestMethod]
		public void TokenizeHandlesInvalidExternalReferenceWithQuotedSheetName()
		{
			string input = "'[1]Some external reference'!#REF!";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(1, tokens.Count());
			var token = tokens.ElementAt(0);
			Assert.AreEqual(TokenType.InvalidReference, token.TokenType);
			Assert.AreEqual("'[1]Some external reference'!#REF!", token.Value);
		}

		[TestMethod]
		public void TokenizeHandlesInvalidExternalReferenceWithSheetName()
		{
			string input = "[1]ExternalReference!#REF!";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(1, tokens.Count());
			var token = tokens.ElementAt(0);
			Assert.AreEqual(TokenType.InvalidReference, token.TokenType);
			Assert.AreEqual("[1]ExternalReference!#REF!", token.Value);
		}

		[TestMethod]
		public void TokenizeHandlesInvalidExternalReferenceWithSheetNameRange()
		{
			string input = "'[1]ExternalReference'!#REF!:C3";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(1, tokens.Count());
			var token = tokens.ElementAt(0);
			Assert.AreEqual(TokenType.InvalidReference, token.TokenType);
			Assert.AreEqual("'[1]ExternalReference'!#REF!:C3", token.Value);
		}

		[TestMethod]
		public void TokenizeHandlesInvalidStartExternalReferenceWithSheetNameRange()
		{
			string input = "'[1]ExternalReference'!#REF!:'[1]ExternalReference'!C3";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(1, tokens.Count());
			var token = tokens.ElementAt(0);
			Assert.AreEqual(TokenType.InvalidReference, token.TokenType);
			Assert.AreEqual("'[1]ExternalReference'!#REF!:'[1]ExternalReference'!C3", token.Value);
		}

		[TestMethod]
		public void TokenizeHandlesInvalidStartSheetExternalReferenceWithSheetNameRange()
		{
			string input = "[1]#REF!C2:C3";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(1, tokens.Count());
			var token = tokens.ElementAt(0);
			Assert.AreEqual(TokenType.InvalidReference, token.TokenType);
			Assert.AreEqual("[1]#REF!C2:C3", token.Value);
		}
		#endregion
		#endregion

		#region Address Tokenization Tests
		[TestMethod]
		public void TokenizeHandlesReference()
		{
			string input = "C3";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(1, tokens.Count());
			var token = tokens.ElementAt(0);
			Assert.AreEqual(TokenType.ExcelAddress, token.TokenType);
			Assert.AreEqual("C3", token.Value);
		}

		[TestMethod]
		public void TokenizeHandlesCellReferenceWithSheetName()
		{
			string input = "Sheet1!C3";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(1, tokens.Count());
			var token = tokens.ElementAt(0);
			Assert.AreEqual(TokenType.ExcelAddress, token.TokenType);
			Assert.AreEqual("Sheet1!C3", token.Value);
		}

		[TestMethod]
		public void TokenizeHandlesCellReferenceWithQuotedSheetName()
		{
			string input = "'Sheet1'!C2";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(1, tokens.Count());
			var token = tokens.ElementAt(0);
			Assert.AreEqual(TokenType.ExcelAddress, token.TokenType);
			Assert.AreEqual("'Sheet1'!C2", token.Value);
		}

		[TestMethod]
		public void TokenizeHandlesCellReferenceWithQuotedSheetNameContainingTicks()
		{
			string input = "'She''et1'!C2";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(1, tokens.Count());
			var token = tokens.ElementAt(0);
			Assert.AreEqual(TokenType.ExcelAddress, token.TokenType);
			Assert.AreEqual("'She''et1'!C2", token.Value);
		}

		[TestMethod]
		public void TokenizeHandlesCellReferenceWithQuotedSheetNameContainingHashtag()
		{
			string input = "'She#et1'!C2";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(1, tokens.Count());
			var token = tokens.ElementAt(0);
			Assert.AreEqual(TokenType.ExcelAddress, token.TokenType);
			Assert.AreEqual("'She#et1'!C2", token.Value);
		}

		[TestMethod]
		public void TokenizeHandlesCellReferenceWithQuotedSheetNameContainingSpecialCharacters()
		{
			string input = "'ab#k.!.2'!A1:A2";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(1, tokens.Count());
			var token = tokens.ElementAt(0);
			Assert.AreEqual(TokenType.ExcelAddress, token.TokenType);
			Assert.AreEqual("'ab#k.!.2'!A1:A2", token.Value);
		}

		[TestMethod]
		public void TokenizeHandlesReferenceRangeWithSheetName()
		{
			string input = "Sheet1!C5:Sheet1!C6";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(1, tokens.Count());
			var token = tokens.ElementAt(0);
			Assert.AreEqual(TokenType.ExcelAddress, token.TokenType);
			Assert.AreEqual("Sheet1!C5:Sheet1!C6", token.Value);
		}

		[TestMethod]
		public void TokenizeHandlesRangeWithQuotedSheetName()
		{
			string input = "'Sheet1'!C2:'Sheet1'!C5";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(1, tokens.Count());
			var token = tokens.ElementAt(0);
			Assert.AreEqual(TokenType.ExcelAddress, token.TokenType);
			Assert.AreEqual("'Sheet1'!C2:'Sheet1'!C5", token.Value);
		}

		[TestMethod]
		public void TokenizeHandlesRangeWithoutSheetName()
		{
			string input = "C2:C5";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(1, tokens.Count());
			var token = tokens.ElementAt(0);
			Assert.AreEqual(TokenType.ExcelAddress, token.TokenType);
			Assert.AreEqual("C2:C5", token.Value);
		}
		
		[TestMethod]
		public void TokenizeHandlesRangeWithStartQuotedSheetName()
		{
			string input = "'Sheet1'!C2:Sheet1!C5";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(1, tokens.Count());
			var token = tokens.ElementAt(0);
			Assert.AreEqual(TokenType.ExcelAddress, token.TokenType);
			Assert.AreEqual("'Sheet1'!C2:Sheet1!C5", token.Value);
		}

		[TestMethod]
		public void TokenizeHandlesRangeWithEndQuotedSheetName()
		{
			string input = "Sheet1!C2:'Sheet1'!C5";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(1, tokens.Count());
			var token = tokens.ElementAt(0);
			Assert.AreEqual(TokenType.ExcelAddress, token.TokenType);
			Assert.AreEqual("Sheet1!C2:'Sheet1'!C5", token.Value);
		}
		
		[TestMethod]
		public void TokenizeHandlesExternalReferenceWithQuotedSheetName()
		{
			string input = "'[1]Some external reference'!C2";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(1, tokens.Count());
			var token = tokens.ElementAt(0);
			Assert.AreEqual(TokenType.InvalidReference, token.TokenType);
			Assert.AreEqual("'[1]Some external reference'!C2", token.Value);
		}

		[TestMethod]
		public void TokenizeHandlesExternalReferenceWithSheetName()
		{
			string input = "[1]ExternalReference!C2";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(1, tokens.Count());
			var token = tokens.ElementAt(0);
			// External references are considered invalid.
			Assert.AreEqual(TokenType.InvalidReference, token.TokenType);
			Assert.AreEqual("[1]ExternalReference!C2", token.Value);
		}

		[TestMethod]
		public void TokenizeHandlesExternalReferenceWithSheetNameRange()
		{
			string input = "'[1]ExternalReference'!C2:C3";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(1, tokens.Count());
			var token = tokens.ElementAt(0);
			Assert.AreEqual(TokenType.InvalidReference, token.TokenType);
			Assert.AreEqual("'[1]ExternalReference'!C2:C3", token.Value);
		}

		[TestMethod]
		public void TokenizeHandlesStartExternalReferenceWithSheetNameRange()
		{
			string input = "'[1]ExternalReference'!C2:'[1]ExternalReference'!C3";
			var tokens = _tokenizer.Tokenize(input);
			Assert.AreEqual(1, tokens.Count());
			var token = tokens.ElementAt(0);
			Assert.AreEqual(TokenType.InvalidReference, token.TokenType);
			Assert.AreEqual("'[1]ExternalReference'!C2:'[1]ExternalReference'!C3", token.Value);
		}
		#endregion
		#endregion

		#region Integration Tokenizer Tests
		[TestMethod]
		public void TestBug9_12_14()
		{
			//(( W60 -(- W63 )-( W29 + W30 + W31 ))/( W23 + W28 + W42 - W51 )* W4 )
			using (var pck = new ExcelPackage())
			{
				var ws1 = pck.Workbook.Worksheets.Add("test");
				for (var x = 1; x <= 10; x++)
				{
					ws1.Cells[x, 1].Value = x;
				}

				ws1.Cells["A11"].Formula = "(( A1 -(- A2 )-( A3 + A4 + A5 ))/( A6 + A7 + A8 - A9 )* A5 )";
				//ws1.Cells["A11"].Formula = "(-A2 + 1 )";
				ws1.Calculate();
				var result = ws1.Cells["A11"].Value;
				Assert.AreEqual(-3.75, result);
			}
		}
		#endregion
	}
}
