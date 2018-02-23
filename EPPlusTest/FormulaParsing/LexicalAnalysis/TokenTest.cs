using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace EPPlusTest.FormulaParsing.LexicalAnalysis
{
	[TestClass]
	public class TokenTest
	{
		#region Constructor Tests
		[TestMethod]
		public void ConstructorTest()
		{
			var token = new Token("value", TokenType.Boolean);
			Assert.AreEqual("value", token.Value);
			Assert.AreEqual(TokenType.Boolean, token.TokenType);
		}
		#endregion

		#region Negate Tests
		[TestMethod]
		public void NegateDecimal()
		{
			var token = new Token("5.3", TokenType.Decimal);
			Assert.IsFalse(token.IsNegated);
			token.Negate();
			Assert.IsTrue(token.IsNegated);
		}

		[TestMethod]
		public void NegateInteger()
		{
			var token = new Token("5", TokenType.Integer);
			Assert.IsFalse(token.IsNegated);
			token.Negate();
			Assert.IsTrue(token.IsNegated);
		}

		[TestMethod]
		public void NegateExcelAddress()
		{
			var token = new Token("E5", TokenType.Integer);
			Assert.IsFalse(token.IsNegated);
			token.Negate();
			Assert.IsTrue(token.IsNegated);
		}

		[TestMethod]
		public void NegateStructuredReference()
		{
			var token = new Token("TableName[#this row]", TokenType.StructuredReference);
			Assert.IsFalse(token.IsNegated);
			token.Negate();
			Assert.IsTrue(token.IsNegated);
		}

		[TestMethod]
		public void NegateInvalidType()
		{
			var token = new Token(";", TokenType.SemiColon);
			Assert.IsFalse(token.IsNegated);
			token.Negate();
			Assert.IsFalse(token.IsNegated);
		}
		#endregion

		#region ToString Tests
		[TestMethod]
		public void ToStringTest()
		{
			var token = new Token("value", TokenType.String);
			Assert.AreEqual("String, value", token.ToString());
		}
		#endregion
	}
}
