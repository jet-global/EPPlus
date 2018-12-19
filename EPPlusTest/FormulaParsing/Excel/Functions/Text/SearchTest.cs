using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Text
{
	[TestClass]
	public class SearchTest
	{
		#region Properties
		private ParsingContext ParsingContext { get; } = ParsingContext.Create();
		#endregion

		#region Test Methods
		[TestMethod]
		public void SearchShouldReturnCorrectValue()
		{
			var function = new Search();
			var args = FunctionsHelper.CreateArgs("is", "This is a test case.");
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(3, result.Result);
			args = FunctionsHelper.CreateArgs("of", "Testing THE CasIng OF The SeaRch FUnctIoN");
			result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(20, result.Result);
		}

		[TestMethod]
		public void SearchShouldReturnCorrectValueWithStartIndex()
		{
			var func = new Search();
			var args = FunctionsHelper.CreateArgs("is", "This is a test case.", 4);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(6, result.Result);
			args = FunctionsHelper.CreateArgs("a", "apple pie", 1);
			result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(1, result.Result);
			args = FunctionsHelper.CreateArgs("a", "banana", 2);
			result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(2, result.Result);
		}

		[TestMethod]
		public void SearchShouldReturnPoundValueErrorIfPhraseNotFound()
		{
			var func = new Search();
			var args = FunctionsHelper.CreateArgs("abc", "This is a test case.");
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SearchWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Search();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SearchWithPrependingQuestionMarkWildcard()
		{
			var function = new Search();
			var args = FunctionsHelper.CreateArgs("?e", "Hello there!");
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(1, result.Result);
			args = FunctionsHelper.CreateArgs("?e", "Hello there!", 5);
			result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(8, result.Result);
		}

		[TestMethod]
		public void SearchWithAppendingQuestionMarkWildcard()
		{
			var function = new Search();
			var args = FunctionsHelper.CreateArgs("e?r", "Trying to test the search function.");
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(21, result.Result);
			args = FunctionsHelper.CreateArgs("e?r", "Trying to test the search function.", 25);
			result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SearchWithWildcardQuestionMarkCharacter()
		{
			var function = new Search();
			var args = FunctionsHelper.CreateArgs("~?", "?Are you there?");
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(1, result.Result);
			args = FunctionsHelper.CreateArgs("~?", "?Are you there?", 5);
			result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(15, result.Result);
			args = FunctionsHelper.CreateArgs("?A", "American Wood Exports");
			result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(6, result.Result);
		}

		[TestMethod]
		public void SearchWithPrependingAsteriskWildcard()
		{
			var function = new Search();
			var args = FunctionsHelper.CreateArgs("*N", "yes no yes no");
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(1, result.Result);
			args = FunctionsHelper.CreateArgs("*N", "yes no yes no", 5);
			result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(5, result.Result);
			args = FunctionsHelper.CreateArgs("*P", "yes no yes no");
			result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SearchWithAsteriskWildcard()
		{
			var function = new Search();
			var args = FunctionsHelper.CreateArgs("m*k", "this is a marker and use it to mark the paper.");
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(11, result.Result);
			args = FunctionsHelper.CreateArgs("*?k", "this is a marker and use it to mark the paper.");
			result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(1, result.Result);
			args = FunctionsHelper.CreateArgs("m*k", "this is a marker and use it to mark the paper.", 16);
			result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(32, result.Result);
			args = FunctionsHelper.CreateArgs("m*k", "this is a marker and use it to park the paper.", 16);
			result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SearchWithAppendingAsteriskWildcard()
		{
			var function = new Search();
			var args = FunctionsHelper.CreateArgs("P*", "Lots of people here.");
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(9, result.Result);
			args = FunctionsHelper.CreateArgs("P*", "Lots of people here.", 15);
			result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SearchWithWildcarAsteriskCharacter()
		{
			var function = new Search();
			var args = FunctionsHelper.CreateArgs("~*", "Where are the ******?");
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(15, result.Result);
			args = FunctionsHelper.CreateArgs("~*", "Where are the ******?", 20);
			result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(20, result.Result);
		}

		[TestMethod]
		public void SearchWithTildaWildcard()
		{
			var function = new Search();
			var args = FunctionsHelper.CreateArgs("~~", "yes~no~yes");
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(4, result.Result);
			args = FunctionsHelper.CreateArgs(" ~", "yes~no~yes");
			result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
			args = FunctionsHelper.CreateArgs("~~", "yes~no~yes", 5);
			result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(7, result.Result);
		}

		[TestMethod]
		public void SearchWithMultipleWildcards()
		{
			var function = new Search();
			var args = FunctionsHelper.CreateArgs("~~*", "American Wo~od Exports");
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(12, result.Result);
			args = FunctionsHelper.CreateArgs("~?~~*", "yes no ?~ no yes");
			result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(8, result.Result);
			args = FunctionsHelper.CreateArgs("~**~?", "Where is the !~ star * now?");
			result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(22, result.Result);
		}
		#endregion
	}
}