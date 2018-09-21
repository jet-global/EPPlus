using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Utils;

namespace EPPlusTest.Table.PivotTable
{
	[TestClass]
	public class UriHelperTests
	{
		#region GetUriEndTargetName Tests
		[TestMethod]
		public void GetUriEndTargetNameTest()
		{
			var uri = new Uri("../pivotCache/pivotCacheDefinition1.xml", UriKind.Relative);
			var uriName = UriHelper.GetUriEndTargetName(uri);
			Assert.AreEqual("pivotCacheDefinition1.xml", uriName);
		}

		[TestMethod]
		public void GetUriEndTargetNameTestIgnoreCase()
		{
			var uri = new Uri("../pivotCache/PIVOTCACHEDEFINITION1.xml", UriKind.Relative);
			var uriName = UriHelper.GetUriEndTargetName(uri);
			Assert.AreEqual("PIVOTCACHEDEFINITION1.xml", uriName);
		}

		[TestMethod]
		public void GetUriEndTargetNameWithExpectedName()
		{
			var uri = new Uri("/worksheets/sheet1.xml", UriKind.Relative);
			var actualName = UriHelper.GetUriEndTargetName(uri, "sheet");
			Assert.AreEqual("sheet1.xml", actualName);
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void GetUriEndTargetNameNullUri()
		{
			UriHelper.GetUriEndTargetName(null);
		}

		[TestMethod]
		[ExpectedException(typeof(InvalidOperationException))]
		public void GetUriEndTargetNameNotFound()
		{
			var uri = new Uri("/worksheets/sheet1.xml", UriKind.Relative);
			var uriName = UriHelper.GetUriEndTargetName(uri, "pivotCacheDefinition");
		}
		#endregion
	}
}