using System;
using System.Xml;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Table.PivotTable;

namespace EPPlusTest.Table.PivotTable
{
	[TestClass]
	public class CacheFieldItemTest
	{
		#region Constructor Tests
		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void CacheFieldItemNullXmlNodeThrowsException()
		{
			new CacheFieldItem(null);
		}
		#endregion

		#region Value Tests
		[TestMethod]
		public void SlicerRangeItemName()
		{
			var xmlDoc = new XmlDocument( );
			xmlDoc.LoadXml("<s v=\"01445544\"/>");
			var node = xmlDoc.FirstChild;
			var cacheFieldItem = new CacheFieldItem(node);
			Assert.AreEqual("01445544", cacheFieldItem.Value);
			cacheFieldItem.Value = "SomethingElse";
			Assert.AreEqual("SomethingElse", cacheFieldItem.Value);
			Assert.AreEqual("<s v=\"SomethingElse\" />", node.OuterXml);
		}
		#endregion
	}
}
