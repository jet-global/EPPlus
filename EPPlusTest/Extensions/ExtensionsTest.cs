using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Extensions;

namespace EPPlusTest.Extensions
{
	[TestClass]
	public class ExtensionsTest
	{
		#region Test Methods
		[TestMethod]
		public void StringIsEquivalentToExactMatch()
		{
			Assert.IsTrue("Sheet1".IsEquivalentTo("Sheet1"));
		}

		[TestMethod]
		public void StringIsEquivalentToCaseInsensitive()
		{
			Assert.IsTrue("Sheet1".IsEquivalentTo("sheet1"));
			Assert.IsTrue("Sheet1".IsEquivalentTo("SHEET1"));
			Assert.IsTrue("sheet1".IsEquivalentTo("SHEET1"));
			Assert.IsTrue("SHEET1".IsEquivalentTo("sheet1"));
		}

		[TestMethod]
		public void StringIsEquivalentToNullAndEmpty()
		{
			var s1 = string.Empty;
			Assert.IsTrue(s1.IsEquivalentTo(string.Empty));
			Assert.IsTrue(s1.IsEquivalentTo(""));
			Assert.IsTrue(s1.IsEquivalentTo(null));
			s1 = null;
			Assert.IsTrue(s1.IsEquivalentTo(string.Empty));
		}

		[TestMethod]
		public void StringIsEquivalentToNotEquivalent()
		{
			Assert.IsFalse("stuff".IsEquivalentTo(string.Empty));
			Assert.IsFalse("stuff".IsEquivalentTo(null));
			Assert.IsFalse("stufff".IsEquivalentTo("stuff"));
			Assert.IsFalse("stuff".IsEquivalentTo("stuff1"));
		}
		#endregion
	}
}
