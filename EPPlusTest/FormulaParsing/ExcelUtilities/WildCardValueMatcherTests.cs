using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;

namespace EPPlusTest.ExcelUtilities
{
	[TestClass]
	public class WildCardValueMatcherTests
	{
		private WildCardValueMatcher _matcher;

		[TestInitialize]
		public void Setup()
		{
			_matcher = new WildCardValueMatcher();
		}

		[TestMethod]
		public void IsMatchTests()
		{
			Assert.AreEqual(0, _matcher.IsMatch("a", "a"));
			Assert.AreEqual(0, _matcher.IsMatch("abcd", "abcd"));
			Assert.AreEqual(0, _matcher.IsMatch("a?c?", "abcd"));
			Assert.AreEqual(0, _matcher.IsMatch("a???", "abcd"));
			Assert.AreEqual(0, _matcher.IsMatch("??c?", "abcd"));
			Assert.AreEqual(0, _matcher.IsMatch("????", "abcd"));
			Assert.AreEqual(0, _matcher.IsMatch("*", "abcd"));
			Assert.AreEqual(0, _matcher.IsMatch("*d", "abcd"));
			Assert.AreEqual(0, _matcher.IsMatch("*cd", "abcd"));
			Assert.AreEqual(0, _matcher.IsMatch("*bcd", "abcd"));
			Assert.AreEqual(0, _matcher.IsMatch("abc*", "abcd"));
			Assert.AreEqual(0, _matcher.IsMatch("ab*", "abcd"));
			Assert.AreEqual(0, _matcher.IsMatch("*b*", "abcd"));
			Assert.AreEqual(0, _matcher.IsMatch("a*", "abcc."));
			Assert.AreEqual(0, _matcher.IsMatch("a*c.", "abcc."));
			Assert.AreEqual(0, _matcher.IsMatch("a*?", "abcd"));
			Assert.AreEqual(0, _matcher.IsMatch("*?", "abcd"));
			Assert.AreEqual(0, _matcher.IsMatch("?*?", "abcd"));
			Assert.AreEqual(0, _matcher.IsMatch("~*", "*"));
			Assert.AreEqual(0, _matcher.IsMatch("~*~*", "**"));
			Assert.AreEqual(0, _matcher.IsMatch("~**", "*otherstuff"));
			Assert.AreEqual(0, _matcher.IsMatch("~?~?", "??"));
			Assert.AreEqual(0, _matcher.IsMatch("~?*", "?what"));
			Assert.AreEqual(0, _matcher.IsMatch("hey~?", "hey?"));
			Assert.AreEqual(0, _matcher.IsMatch("hey?", "hey?"));
			Assert.AreEqual(0, _matcher.IsMatch("hey~~*", "hey~*"));
			Assert.AreEqual(0, _matcher.IsMatch("~~", "~"));
			Assert.AreEqual(0, _matcher.IsMatch("~~~*ab~~~~~*a", "~*ab~~*a"));
			Assert.AreEqual(0, _matcher.IsMatch("~~~*ab~~~~~*~??", "~*ab~~*?a"));
			Assert.AreEqual(0, _matcher.IsMatch("'", "'"));
			Assert.AreEqual(0, _matcher.IsMatch(@"""", @""""));
			Assert.AreEqual(0, _matcher.IsMatch(@"""""", @""""""));
			Assert.AreEqual(0, _matcher.IsMatch(@".", @"."));
			Assert.AreEqual(0, _matcher.IsMatch(@"hey.", @"hey."));
			Assert.AreEqual(0, _matcher.IsMatch(@"hey!", @"hey!"));
		}

		[TestMethod]
		public void IsMatchNotAMatchTests()
		{
			Assert.AreNotEqual(0, _matcher.IsMatch("a", "abcd"));
			Assert.AreNotEqual(0, _matcher.IsMatch("a?", "abcd"));
			Assert.AreNotEqual(0, _matcher.IsMatch("a??", "abcd"));
			Assert.AreNotEqual(0, _matcher.IsMatch("a????", "abcd"));
			Assert.AreNotEqual(0, _matcher.IsMatch("???", "abcd"));
			Assert.AreNotEqual(0, _matcher.IsMatch("?????", "abcd"));
			Assert.AreNotEqual(0, _matcher.IsMatch("?", "abcd"));
			Assert.AreNotEqual(0, _matcher.IsMatch("~*", "a"));
			Assert.AreNotEqual(0, _matcher.IsMatch("~*", "?"));
			Assert.AreNotEqual(0, _matcher.IsMatch("~?", "*"));
			Assert.AreEqual(0, _matcher.IsMatch("~", "~"));
		}
	}
}
