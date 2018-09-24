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
		public void CacheFieldItemConstructorTest()
		{
			XmlDocument document = new XmlDocument();
			document.LoadXml(@"<cacheField name=""Item"" numFmtId=""0""><sharedItems count=""2""><s v=""Bike""/><s v=""Car""/></sharedItems></cacheField>");
			var node = document.SelectSingleNode("//cacheField");
			var item = new CacheFieldItem(node, "jet");
			Assert.AreEqual("jet", item.Value);
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void CacheFieldItemNullParentNodeThrowsException()
		{
			new CacheFieldItem(null, "jet");
		}

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
