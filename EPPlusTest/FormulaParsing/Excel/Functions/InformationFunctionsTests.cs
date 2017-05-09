using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;

namespace EPPlusTest.Excel.Functions
{
	[TestClass]
	public class InformationFunctionsTests
	{
		private ParsingContext _context;

		[TestInitialize]
		public void Setup()
		{
			_context = ParsingContext.Create();
		}

		[TestMethod]
		public void IsBlankShouldReturnTrueIfFirstArgIsNull()
		{
			var func = new IsBlank();
			var args = FunctionsHelper.CreateArgs(new object[] { null });
			var result = func.Execute(args, _context);
			Assert.IsTrue((bool)result.Result);
		}

		[TestMethod]
		public void IsBlankShouldReturnTrueIfFirstArgIsEmptyString()
		{
			var func = new IsBlank();
			var args = FunctionsHelper.CreateArgs(string.Empty);
			var result = func.Execute(args, _context);
			Assert.IsTrue((bool)result.Result);
		}

		[TestMethod]
		public void IsNumberShouldReturnTrueWhenArgIsNumeric()
		{
			var func = new IsNumber();
			var args = FunctionsHelper.CreateArgs(1d);
			var result = func.Execute(args, _context);
			Assert.IsTrue((bool)result.Result);
		}

		[TestMethod]
		public void IsNumberShouldReturnfalseWhenArgIsNonNumeric()
		{
			var func = new IsNumber();
			var args = FunctionsHelper.CreateArgs("1");
			var result = func.Execute(args, _context);
			Assert.IsFalse((bool)result.Result);
		}

		[TestMethod]
		public void IsErrorShouldReturnTrueIfArgIsAnErrorCode()
		{
			var args = FunctionsHelper.CreateArgs(ExcelErrorValue.Parse("#DIV/0!"));
			var func = new IsError();
			var result = func.Execute(args, _context);
			Assert.IsTrue((bool)result.Result);
		}

		[TestMethod]
		public void IsErrorShouldReturnFalseIfArgIsNotAnError()
		{
			var args = FunctionsHelper.CreateArgs("A", 1);
			var func = new IsError();
			var result = func.Execute(args, _context);
			Assert.IsFalse((bool)result.Result);
		}

		[TestMethod]
		public void IsTextShouldReturnTrueWhenFirstArgIsAString()
		{
			var args = FunctionsHelper.CreateArgs("1");
			var func = new IsText();
			var result = func.Execute(args, _context);
			Assert.IsTrue((bool)result.Result);
		}

		[TestMethod]
		public void IsTextShouldReturnFalseWhenFirstArgIsNotAString()
		{
			var args = FunctionsHelper.CreateArgs(1);
			var func = new IsText();
			var result = func.Execute(args, _context);
			Assert.IsFalse((bool)result.Result);
		}

		[TestMethod]
		public void IsNonTextShouldReturnFalseWhenFirstArgIsAString()
		{
			var args = FunctionsHelper.CreateArgs("1");
			var func = new IsNonText();
			var result = func.Execute(args, _context);
			Assert.IsFalse((bool)result.Result);
		}

		[TestMethod]
		public void IsNonTextShouldReturnTrueWhenFirstArgIsNotAString()
		{
			var args = FunctionsHelper.CreateArgs(1);
			var func = new IsNonText();
			var result = func.Execute(args, _context);
			Assert.IsTrue((bool)result.Result);
		}

		[TestMethod]
		public void IsOddShouldReturnCorrectResult()
		{
			var args = FunctionsHelper.CreateArgs(3.123);
			var func = new IsOdd();
			var result = func.Execute(args, _context);
			Assert.IsTrue((bool)result.Result);
		}


		[TestMethod]
		public void IsOddShouldPoundValueOnNonNumericInput()
		{
			var args = FunctionsHelper.CreateArgs("Not odd");
			var func = new IsOdd();
			var result = func.Execute(args, _context);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void IsEvenShouldReturnCorrectResult()
		{
			var args = FunctionsHelper.CreateArgs(4.123);
			var func = new IsEven();
			var result = func.Execute(args, _context);
			Assert.IsTrue((bool)result.Result);
		}

		[TestMethod]
		public void IsEvenShouldPoundValueOnNonNumericInput()
		{
			var args = FunctionsHelper.CreateArgs("Not odd");
			var func = new IsEven();
			var result = func.Execute(args, _context);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void IsLogicalShouldReturnCorrectResult()
		{
			var func = new IsLogical();

			var args = FunctionsHelper.CreateArgs(1);
			var result = func.Execute(args, _context);
			Assert.IsFalse((bool)result.Result);

			args = FunctionsHelper.CreateArgs("true");
			result = func.Execute(args, _context);
			Assert.IsFalse((bool)result.Result);

			args = FunctionsHelper.CreateArgs(false);
			result = func.Execute(args, _context);
			Assert.IsTrue((bool)result.Result);
		}

		[TestMethod]
		public void NshouldReturnCorrectResult()
		{
			var func = new N();

			var args = FunctionsHelper.CreateArgs(1.2);
			var result = func.Execute(args, _context);
			Assert.AreEqual(1.2, result.Result);

			args = FunctionsHelper.CreateArgs("abc");
			result = func.Execute(args, _context);
			Assert.AreEqual(0d, result.Result);

			args = FunctionsHelper.CreateArgs(true);
			result = func.Execute(args, _context);
			Assert.AreEqual(1d, result.Result);

			var errorCode = ExcelErrorValue.Create(eErrorType.Value);
			args = FunctionsHelper.CreateArgs(errorCode);
			result = func.Execute(args, _context);
			Assert.AreEqual(errorCode, result.Result);
		}

		[TestMethod]
		public void ErrorTypeWithInvalidArgumentReturnsPoundValue()
		{
			var func = new ErrorType();
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void IsEvenWithInvalidArgumentReturnsPoundValue()
		{
			var func = new IsEven();
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void IsLogicalWithInvalidArgumentReturnsPoundValue()
		{
			var func = new IsLogical();
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void IsNonTextWithInvalidArgumentReturnsPoundValue()
		{
			var func = new IsNonText();
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void IsNumberWithInvalidArgumentReturnsPoundValue()
		{
			var func = new IsNumber();
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void IsOddWithInvalidArgumentReturnsPoundValue()
		{
			var func = new IsOdd();
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void IsTextWithInvalidArgumentReturnsPoundValue()
		{
			var func = new IsText();
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void NWithInvalidArgumentReturnsPoundValue()
		{
			var func = new N();
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, parsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

	}
}
