using System;
using System.Collections.Generic;
using System.Xml;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Table.PivotTable;

namespace EPPlusTest.Table.PivotTable
{
	[TestClass]
	public class CacheItemTest : PivotTableTestBase
	{
		#region Constructor Tests
		[TestMethod]
		public void CacheItemParsesTypeCorrectly()
		{
			XmlDocument document = new XmlDocument();
			document.LoadXml(@"<pivotCacheRecords count=""1""><r><n v=""20100076""/><x v=""0""/> <b v=""0""/> <m v=""0""/> <e v=""415.75""/><d v=""1""/></r></pivotCacheRecords>");
			var node = document.SelectSingleNode("//n");
			var namespaceManager = TestUtility.CreateDefaultNSM();
			var record = new CacheItem(namespaceManager, node);
			Assert.AreEqual(PivotCacheRecordType.n, record.Type);
			Assert.AreEqual("20100076", record.Value);
			node = document.SelectSingleNode("//x");
			record = new CacheItem(namespaceManager, node);
			Assert.AreEqual(PivotCacheRecordType.x, record.Type);
			node = document.SelectSingleNode("//m");
			record = new CacheItem(namespaceManager, node);
			Assert.AreEqual(PivotCacheRecordType.m, record.Type);
			node = document.SelectSingleNode("//e");
			record = new CacheItem(namespaceManager, node);
			Assert.AreEqual(PivotCacheRecordType.e, record.Type);
		}

		[TestMethod]
		public void CacheItemCacheRecordsTest()
		{
			XmlDocument document = new XmlDocument();
			document.LoadXml(@"<pivotCacheRecords count=""1""><r><n v=""20100076""/> <x v=""0""/> <b v=""0""/> <e v=""415.75""/><d v=""1""/></r></pivotCacheRecords>");
			var namespaceManager = TestUtility.CreateDefaultNSM();
			var node = document.SelectSingleNode("//r");
			var record = new CacheItem(namespaceManager, node, PivotCacheRecordType.n, "382");
			Assert.AreEqual(PivotCacheRecordType.n, record.Type);
			Assert.AreEqual("382", record.Value);
			Assert.IsNotNull(node.SelectSingleNode("//n[@v=382]"));
			record = new CacheItem(namespaceManager, node, PivotCacheRecordType.b, "1");
			Assert.AreEqual(PivotCacheRecordType.b, record.Type);
			Assert.AreEqual("1", record.Value);
			Assert.IsNotNull(node.SelectSingleNode("//b[@v=1]"));
			record = new CacheItem(namespaceManager, node, PivotCacheRecordType.m, "");
			Assert.AreEqual(PivotCacheRecordType.m, record.Type);
			Assert.IsNull(record.Value);
			Assert.IsNotNull(node.SelectSingleNode("//m"));
		}

		[TestMethod]
		public void CacheItemCacheFieldTest()
		{
			XmlDocument document = new XmlDocument();
			document.LoadXml(@"<sharedItems><n v=""20100076""/> <x v=""0""/> <b v=""0""/> <e v=""415.75""/><d v=""1""/></sharedItems>");
			var namespaceManager = TestUtility.CreateDefaultNSM();
			var node = document.SelectSingleNode("//sharedItems");
			var record = new CacheItem(namespaceManager, node, PivotCacheRecordType.n, "382");
			Assert.AreEqual(PivotCacheRecordType.n, record.Type);
			Assert.AreEqual("382", record.Value);
			Assert.IsNotNull(node.SelectSingleNode("//n[@v=382]"));
			record = new CacheItem(namespaceManager, node, PivotCacheRecordType.b, "1");
			Assert.AreEqual(PivotCacheRecordType.b, record.Type);
			Assert.AreEqual("1", record.Value);
			Assert.IsNotNull(node.SelectSingleNode("//b[@v=1]"));
			record = new CacheItem(namespaceManager, node, PivotCacheRecordType.m, "");
			Assert.AreEqual(PivotCacheRecordType.m, record.Type);
			Assert.IsNull(record.Value);
			Assert.IsNotNull(node.SelectSingleNode("//m"));
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentException))]
		public void CacheItemCreateRecordItemWithIncorrectParentNode()
		{
			XmlDocument document = new XmlDocument();
			document.LoadXml(@"<cacheField name=""Item"" numFmtId=""0""><sharedItems count=""2""><s v=""Bike""/><s v=""Car""/></sharedItems></cacheField>");
			var node = document.SelectSingleNode("//cacheField", TestUtility.CreateDefaultNSM());
			var item = new CacheItem(TestUtility.CreateDefaultNSM(), node, PivotCacheRecordType.n, "382");
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void CacheItemNullNodeTest()
		{
			new CacheItem(TestUtility.CreateDefaultNSM(), null);
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void CacheItemNullParentNodeThrowsException()
		{
			new CacheItem(TestUtility.CreateDefaultNSM(), null, PivotCacheRecordType.s, "jet");
		}
		#endregion

		#region ReplaceNode Tests
		[TestMethod]
		public void ReplaceNode()
		{
			XmlDocument document = new XmlDocument();
			document.LoadXml(@"<cacheField name=""Item"" numFmtId=""0""><sharedItems count=""1""><s v=""Bike""/></sharedItems></cacheField>");
			var namespaceManager = TestUtility.CreateDefaultNSM();
			var parentNode = document.SelectSingleNode("//sharedItems", namespaceManager);
			var itemNode = document.SelectSingleNode("//s", namespaceManager);
			var cacheItem = new CacheItem(namespaceManager, itemNode);
			cacheItem.ReplaceNode(PivotCacheRecordType.n, "930", parentNode);
			Assert.AreEqual("930", cacheItem.Value);
			Assert.AreEqual(PivotCacheRecordType.n, cacheItem.Type);
			Assert.AreEqual(1, parentNode.ChildNodes.Count);
			Assert.AreEqual("n", parentNode.FirstChild.Name);
			Assert.AreEqual("930", parentNode.FirstChild.Attributes["v"].Value);
		}

		[TestMethod]
		public void ReplaceNodeWithNoValueType()
		{
			XmlDocument document = new XmlDocument();
			document.LoadXml(@"<cacheField name=""Item"" numFmtId=""0""><sharedItems count=""1""><s v=""Bike""/></sharedItems></cacheField>");
			var namespaceManager = TestUtility.CreateDefaultNSM();
			var parentNode = document.SelectSingleNode("//sharedItems", namespaceManager);
			var itemNode = document.SelectSingleNode("//s", namespaceManager);
			var cacheItem = new CacheItem(namespaceManager, itemNode);
			cacheItem.ReplaceNode(PivotCacheRecordType.m, null, parentNode);
			Assert.IsNull(cacheItem.Value);
			Assert.AreEqual(PivotCacheRecordType.m, cacheItem.Type);
			Assert.AreEqual(1, parentNode.ChildNodes.Count);
			Assert.AreEqual("m", parentNode.FirstChild.Name);
			Assert.IsNull(parentNode.FirstChild.Attributes["v"]);
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void ReplaceNodeNullParentNode()
		{
			XmlDocument document = new XmlDocument();
			document.LoadXml(@"<cacheField name=""Item"" numFmtId=""0""><sharedItems count=""1""><s v=""Bike""/></sharedItems></cacheField>");
			var namespaceManager = TestUtility.CreateDefaultNSM();
			var itemNode = document.SelectSingleNode("//s", namespaceManager);
			var cacheItem = new CacheItem(namespaceManager, itemNode);
			cacheItem.ReplaceNode(PivotCacheRecordType.m, null, null);
		}
		#endregion

		#region GetObjectType Tests
		[TestMethod]
		public void GetObectType()
		{
			var type = CacheItem.GetObjectType(true);
			Assert.AreEqual(PivotCacheRecordType.b, type);
			type = CacheItem.GetObjectType(null);
			Assert.AreEqual(PivotCacheRecordType.m, type);
			type = CacheItem.GetObjectType(string.Empty);
			Assert.AreEqual(PivotCacheRecordType.m, type);
			type = CacheItem.GetObjectType("string");
			Assert.AreEqual(PivotCacheRecordType.s, type);
			type = CacheItem.GetObjectType(83);
			Assert.AreEqual(PivotCacheRecordType.n, type);
			type = CacheItem.GetObjectType(3920.124);
			Assert.AreEqual(PivotCacheRecordType.n, type);
			type = CacheItem.GetObjectType(DateTime.Now);
			Assert.AreEqual(PivotCacheRecordType.d, type);
			type = CacheItem.GetObjectType(ExcelErrorValue.Create(eErrorType.NA));
			Assert.AreEqual(PivotCacheRecordType.e, type);
		}

		[TestMethod]
		[ExpectedException(typeof(InvalidOperationException))]
		public void GetObectTypeInvalidType()
		{
			CacheItem.GetObjectType(new List<int>());
		}
		#endregion

		#region UpdateValue Tests
		//[TestMethod]
		//public void UpdateValueDifferentType()
		//{
		//	XmlDocument document = new XmlDocument();
		//	document.LoadXml(@"<pivotCacheRecords count=""1""><r><n v=""20100076""/></r></pivotCacheRecords>");
		//	var node = document.SelectSingleNode("//n");
		//	var record = new CacheRecordItem(node, TestUtility.CreateDefaultNSM());
		//	var parentNode = document.SelectSingleNode("//r");
		//	var cacheFieldNode = base.GetTestCacheFieldNode();
		//	Assert.IsFalse(record.UpdateValue(true, parentNode, cacheFieldNode));
		//	Assert.AreEqual("True", record.Value);
		//	Assert.AreEqual(PivotCacheRecordType.b, record.Type);
		//}

		//[TestMethod]
		//public void UpdateValueSameType()
		//{
		//	XmlDocument document = new XmlDocument();
		//	document.LoadXml(@"<pivotCacheRecords count=""1""><r><n v=""20100076""/></r></pivotCacheRecords>");
		//	var node = document.SelectSingleNode("//n");
		//	var record = new CacheRecordItem(node, TestUtility.CreateDefaultNSM());
		//	var parentNode = document.SelectSingleNode("//r");
		//	var cacheFieldNode = base.GetTestCacheFieldNode();
		//	Assert.IsFalse(record.UpdateValue(389471230, parentNode, cacheFieldNode));
		//	Assert.AreEqual("389471230", record.Value);
		//	Assert.AreEqual(PivotCacheRecordType.n, record.Type);
		//}

		//[TestMethod]
		//public void UpdateValueWithNewSharedString()
		//{
		//	XmlDocument document = new XmlDocument();
		//	document.LoadXml(@"<pivotCacheRecords count=""1""><r><n v=""20100076""/></r></pivotCacheRecords>");
		//	var node = document.SelectSingleNode("//n");
		//	var record = new CacheRecordItem(node, TestUtility.CreateDefaultNSM());
		//	var parentNode = document.SelectSingleNode("//r");
		//	var cacheFieldNode = base.GetTestCacheFieldNode();
		//	Assert.IsTrue(record.UpdateValue("red", parentNode, cacheFieldNode));
		//	Assert.AreEqual("2", record.Value);
		//	Assert.AreEqual(PivotCacheRecordType.x, record.Type);
		//}

		//[TestMethod]
		//public void UpdateValueWithNewSharedStringSameType()
		//{
		//	XmlDocument document = new XmlDocument();
		//	document.LoadXml(@"<pivotCacheRecords count=""1""><r><x v=""1""/></r></pivotCacheRecords>");
		//	var node = document.SelectSingleNode("//x");
		//	var record = new CacheRecordItem(node, TestUtility.CreateDefaultNSM());
		//	var parentNode = document.SelectSingleNode("//r");
		//	var cacheFieldNode = base.GetTestCacheFieldNode();
		//	Assert.IsTrue(record.UpdateValue("red", parentNode, cacheFieldNode));
		//	Assert.AreEqual("2", record.Value);
		//	Assert.AreEqual(PivotCacheRecordType.x, record.Type);
		//}

		//[TestMethod]
		//public void UpdateValueWithExistingSharedStringDifferentTypes()
		//{
		//	XmlDocument document = new XmlDocument();
		//	document.LoadXml(@"<pivotCacheRecords count=""1""><r><n v=""20100076""/></r></pivotCacheRecords>");
		//	var node = document.SelectSingleNode("//n");
		//	var record = new CacheRecordItem(node, TestUtility.CreateDefaultNSM());
		//	var parentNode = document.SelectSingleNode("//r");
		//	var cacheFieldNode = base.GetTestCacheFieldNode();
		//	Assert.IsFalse(record.UpdateValue("Car", parentNode, cacheFieldNode));
		//	Assert.AreEqual("1", record.Value);
		//	Assert.AreEqual(PivotCacheRecordType.x, record.Type);
		//}

		//[TestMethod]
		//public void UpdateValueWithExistingSharedStringSameTypes()
		//{
		//	XmlDocument document = new XmlDocument();
		//	document.LoadXml(@"<pivotCacheRecords count=""1""><r><n x=""4""/></r></pivotCacheRecords>");
		//	var node = document.SelectSingleNode("//n");
		//	var record = new CacheRecordItem(node, TestUtility.CreateDefaultNSM());
		//	var parentNode = document.SelectSingleNode("//r");
		//	var cacheFieldNode = base.GetTestCacheFieldNode();
		//	Assert.IsFalse(record.UpdateValue("Car", parentNode, cacheFieldNode));
		//	Assert.AreEqual("1", record.Value);
		//	Assert.AreEqual(PivotCacheRecordType.x, record.Type);
		//}

		//[TestMethod]
		//public void UpdateValueNullValue()
		//{
		//	XmlDocument document = new XmlDocument();
		//	document.LoadXml(@"<pivotCacheRecords count=""1""><r><n v=""20100076""/></r></pivotCacheRecords>");
		//	var node = document.SelectSingleNode("//n");
		//	var record = new CacheRecordItem(node, TestUtility.CreateDefaultNSM());
		//	var parentNode = document.SelectSingleNode("//r");
		//	var cacheFieldNode = base.GetTestCacheFieldNode();
		//	Assert.IsFalse(record.UpdateValue(null, parentNode, cacheFieldNode));
		//	Assert.IsNull(record.Value);
		//	Assert.AreEqual(PivotCacheRecordType.m, record.Type);
		//}

		//[TestMethod]
		//public void UpdateValueEmptyStringValue()
		//{
		//	XmlDocument document = new XmlDocument();
		//	document.LoadXml(@"<pivotCacheRecords count=""1""><r><n v=""20100076""/></r></pivotCacheRecords>");
		//	var node = document.SelectSingleNode("//n");
		//	var record = new CacheRecordItem(node, TestUtility.CreateDefaultNSM());
		//	var parentNode = document.SelectSingleNode("//r");
		//	var cacheFieldNode = base.GetTestCacheFieldNode();
		//	Assert.IsFalse(record.UpdateValue(string.Empty, parentNode, cacheFieldNode));
		//	Assert.IsNull(record.Value);
		//	Assert.AreEqual(PivotCacheRecordType.m, record.Type);
		//}

		//[TestMethod]
		//[ExpectedException(typeof(InvalidOperationException))]
		//public void UpdateValueNotSupportedType()
		//{
		//	XmlDocument document = new XmlDocument();
		//	document.LoadXml(@"<pivotCacheRecords count=""1""><r><n v=""20100076""/></r></pivotCacheRecords>");
		//	var node = document.SelectSingleNode("//n");
		//	var record = new CacheRecordItem(node, TestUtility.CreateDefaultNSM());
		//	var parentNode = document.SelectSingleNode("//r");
		//	var cacheFieldNode = base.GetTestCacheFieldNode();
		//	record.UpdateValue(record, parentNode, cacheFieldNode);
		//}

		//[TestMethod]
		//[ExpectedException(typeof(ArgumentNullException))]
		//public void UpdateValueNullParentNode()
		//{
		//	XmlDocument document = new XmlDocument();
		//	document.LoadXml(@"<pivotCacheRecords count=""1""><r><n v=""20100076""/></r></pivotCacheRecords>");
		//	var node = document.SelectSingleNode("//n");
		//	var record = new CacheRecordItem(node, TestUtility.CreateDefaultNSM());
		//	var parentNode = document.SelectSingleNode("//r");
		//	var cacheFieldNode = base.GetTestCacheFieldNode();
		//	record.UpdateValue(record, null, cacheFieldNode);
		//}
		#endregion
	}
}