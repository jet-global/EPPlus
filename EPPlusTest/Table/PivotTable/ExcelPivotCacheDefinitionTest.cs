using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Table.PivotTable;

namespace EPPlusTest.Table.PivotTable
{
	[TestClass]
	public class ExcelPivotCacheDefinitionTest
	{
		#region GetCacheDefinitionUriName Tests
		[TestMethod]
		public void GetCacheDefinitionUriNameTest()
		{
			var uri = new Uri("../pivotCache/pivotCacheDefinition1.xml", UriKind.Relative);
			var uriName = ExcelPivotCacheDefinition.GetCacheDefinitionUriName(uri);
			Assert.AreEqual("pivotCacheDefinition1.xml", uriName);
		}

		[TestMethod]
		public void GetCacheDefinitionUriNameTestIgnoreCase()
		{
			var uri = new Uri("../pivotCache/PIVOTCACHEDEFINITION1.xml", UriKind.Relative);
			var uriName = ExcelPivotCacheDefinition.GetCacheDefinitionUriName(uri);
			Assert.AreEqual("PIVOTCACHEDEFINITION1.xml", uriName);
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void GetCacheDefinitionUriNameNullUri()
		{
			ExcelPivotCacheDefinition.GetCacheDefinitionUriName(null);
		}

		[TestMethod]
		[ExpectedException(typeof(InvalidOperationException))]
		public void GetCacheDefinitionUriNameNotFound()
		{
			var uri = new Uri("/worksheets/sheet1.xml", UriKind.Relative);
			ExcelPivotCacheDefinition.GetCacheDefinitionUriName(uri);
		}
		#endregion
	}
}