using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace EPPlusTest.Excel.Functions.Text
{
	[TestClass]
	public class TextFunctionsTests
	{
		#region Properties
		private ParsingContext ParsingContext { get; } = ParsingContext.Create();
		#endregion

		#region Test Methods
		[TestMethod]
		public void CStrShouldConvertNumberToString()
		{
			var func = new CStr();
			var result = func.Execute(FunctionsHelper.CreateArgs(1), this.ParsingContext);
			Assert.AreEqual(DataType.String, result.DataType);
			Assert.AreEqual("1", result.Result);
		}

		[TestMethod]
		public void LenShouldReturnStringsLength()
		{
			var func = new Len();
			var result = func.Execute(FunctionsHelper.CreateArgs("abc"), this.ParsingContext);
			Assert.AreEqual(3d, result.Result);
		}

		[TestMethod]
		public void LowerShouldReturnLowerCaseString()
		{
			var func = new Lower();
			var result = func.Execute(FunctionsHelper.CreateArgs("ABC"), this.ParsingContext);
			Assert.AreEqual("abc", result.Result);
		}

		[TestMethod]
		public void UpperShouldReturnUpperCaseString()
		{
			var func = new Upper();
			var result = func.Execute(FunctionsHelper.CreateArgs("abc"), this.ParsingContext);
			Assert.AreEqual("ABC", result.Result);
		}

		[TestMethod]
		public void LeftShouldReturnSubstringFromLeft()
		{
			var func = new Left();
			var result = func.Execute(FunctionsHelper.CreateArgs("abcd", 2), this.ParsingContext);
			Assert.AreEqual("ab", result.Result);
		}

		[TestMethod]
		public void RightShouldReturnSubstringFromRight()
		{
			var func = new Right();
			var result = func.Execute(FunctionsHelper.CreateArgs("abcd", 2), this.ParsingContext);
			Assert.AreEqual("cd", result.Result);
		}

		[TestMethod]
		public void MidShouldReturnSubstringAccordingToParams()
		{
			var func = new Mid();
			var result = func.Execute(FunctionsHelper.CreateArgs("abcd", 1, 2), this.ParsingContext);
			Assert.AreEqual("ab", result.Result);
		}

		[TestMethod]
		public void ReplaceShouldReturnAReplacedStringAccordingToParamsWhenStartIxIs1()
		{
			var func = new Replace();
			var result = func.Execute(FunctionsHelper.CreateArgs("testar", 1, 2, "hej"), this.ParsingContext);
			Assert.AreEqual("hejstar", result.Result);
		}

		[TestMethod]
		public void ReplaceShouldReturnAReplacedStringAccordingToParamsWhenStartIxIs3()
		{
			var func = new Replace();
			var result = func.Execute(FunctionsHelper.CreateArgs("testar", 3, 3, "hej"), this.ParsingContext);
			Assert.AreEqual("tehejr", result.Result);
		}

		[TestMethod]
		public void SubstituteShouldReturnAReplacedStringAccordingToParamsWhen()
		{
			var func = new Substitute();
			var result = func.Execute(FunctionsHelper.CreateArgs("testar testar", "es", "xx"), this.ParsingContext);
			Assert.AreEqual("txxtar txxtar", result.Result);
		}

		[TestMethod]
		public void ConcatenateShouldConcatenateThreeStrings()
		{
			var func = new Concatenate();
			var result = func.Execute(FunctionsHelper.CreateArgs("One", "Two", "Three"), this.ParsingContext);
			Assert.AreEqual("OneTwoThree", result.Result);
		}

		[TestMethod]
		public void ConcatenateShouldConcatenateStringWithInt()
		{
			var func = new Concatenate();
			var result = func.Execute(FunctionsHelper.CreateArgs(1, "Two"), this.ParsingContext);
			Assert.AreEqual("1Two", result.Result);
		}

		[TestMethod]
		public void ConcatenatePropogatesErrors()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Test");
				ws.Cells["A1"].Formula = @"noname"; // #NAME?
				ws.Cells["A2"].Formula = @"""hey""+1"; // #VALUE!
				ws.Cells["A3"].Formula = "CONCATENATE(A1,A2)";
				ws.Calculate();
				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Name), ws.Cells["A1"].Value);
				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), ws.Cells["A2"].Value);
				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Name), ws.Cells["A3"].Value);
			}
		}

		[TestMethod]
		public void ConcatenateSingleString()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Test");
				ws.Cells["A3"].Formula = @"CONCATENATE(""hey"")";
				ws.Calculate();
				Assert.AreEqual("hey", ws.Cells["A3"].Value);
			}
		}

		[TestMethod]
		public void ConcatenateSingleReference()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Test");
				ws.Cells["A1"].Value = @"hey";
				ws.Cells["A3"].Formula = @"CONCATENATE(A1)";
				ws.Calculate();
				Assert.AreEqual("hey", ws.Cells["A3"].Value);
			}
		}

		[TestMethod]
		public void ExactShouldReturnTrueWhenTwoEqualStrings()
		{
			var func = new Exact();
			var result = func.Execute(FunctionsHelper.CreateArgs("abc", "abc"), this.ParsingContext);
			Assert.IsTrue((bool)result.Result);
		}

		[TestMethod]
		public void ExactShouldReturnTrueWhenEqualStringAndDouble()
		{
			var func = new Exact();
			var result = func.Execute(FunctionsHelper.CreateArgs("1", 1d), this.ParsingContext);
			Assert.IsTrue((bool)result.Result);
		}

		[TestMethod]
		public void ExactShouldReturnFalseWhenStringAndNull()
		{
			var func = new Exact();
			var result = func.Execute(FunctionsHelper.CreateArgs("1", null), this.ParsingContext);
			Assert.IsFalse((bool)result.Result);
		}

		[TestMethod]
		public void ExactShouldReturnFalseWhenTwoEqualStringsWithDifferentCase()
		{
			var func = new Exact();
			var result = func.Execute(FunctionsHelper.CreateArgs("abc", "Abc"), this.ParsingContext);
			Assert.IsFalse((bool)result.Result);
		}

		[TestMethod]
		public void FindShouldReturnIndexOfFoundPhrase()
		{
			var func = new Find();
			var result = func.Execute(FunctionsHelper.CreateArgs("hopp", "hej hopp"), this.ParsingContext);
			Assert.AreEqual(5, result.Result);
		}

		[TestMethod]
		public void FindShouldReturnIndexOfFoundPhraseBasedOnStartIndex()
		{
			var func = new Find();
			var result = func.Execute(FunctionsHelper.CreateArgs("hopp", "hopp hopp", 2), this.ParsingContext);
			Assert.AreEqual(6, result.Result);
		}
		
		[TestMethod]
		public void ProperShouldSetFirstLetterToUpperCase()
		{
			var func = new Proper();
			var result = func.Execute(FunctionsHelper.CreateArgs("this IS A tEst.wi3th SOME w0rds östEr"), this.ParsingContext);
			Assert.AreEqual("This Is A Test.Wi3Th Some W0Rds Öster", result.Result);
		}

		[TestMethod]
		public void HyperLinkShouldReturnArgIfOneArgIsSupplied()
		{
			var func = new Hyperlink();
			var result = func.Execute(FunctionsHelper.CreateArgs("http://epplus.codeplex.com"), this.ParsingContext);
			Assert.AreEqual("http://epplus.codeplex.com", result.Result);
		}

		[TestMethod]
		public void HyperLinkShouldReturnLastArgIfTwoArgsAreSupplied()
		{
			var func = new Hyperlink();
			var result = func.Execute(FunctionsHelper.CreateArgs("http://epplus.codeplex.com", "EPPlus"), this.ParsingContext);
			Assert.AreEqual("EPPlus", result.Result);
		}

		[TestMethod]
		public void CStrWithInvalidArgumentReturnsPoundValue()
		{
			var func = new CStr();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void ExactWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Exact();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void FindWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Find();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void FixedWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Fixed();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void HyperlinkWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Hyperlink();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void LeftWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Left();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void CharWithInvalidArgumentReturnsPoundValue()
		{
			var func = new CharFunction();
			var args = FunctionsHelper.CreateArgs(277);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void CharWithNoArgumentReturnsPoundValue()
		{
			var func = new CharFunction();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void LenWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Len();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void LowerWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Lower();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void MidWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Mid();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void ProperWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Proper();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void ReplaceWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Replace();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void ReptWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Rept();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void RightWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Right();
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SubstituteWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Substitute();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void TWithInvalidArgumentReturnsPoundValue()
		{
			var func = new T();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void TextWithInvalidArgumentReturnsPoundValue()
		{
			var func = new OfficeOpenXml.FormulaParsing.Excel.Functions.Text.Text();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void UpperWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Upper();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void ValueWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Value();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}
		#endregion
	}
}
