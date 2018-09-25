using System;
using System.Xml;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Table.PivotTable;

namespace EPPlusTest.Table.PivotTable
{
	[TestClass]
	public class CacheRecordItemTest : PivotTableTestBase
	{
		#region Constructor Tests
		[TestMethod]
		public void CacheRecordItemParsesTypeCorrectly()
		{
			XmlDocument document = new XmlDocument();
			document.LoadXml(@"<pivotCacheRecords count=""1""><r><n v=""20100076""/><x v=""0""/> <b v=""0""/> <m v=""0""/> <e v=""415.75""/><d v=""1""/></r></pivotCacheRecords>");
			var node = document.SelectSingleNode("//n");
			var record = new CacheRecordItem(node);
			Assert.AreEqual(PivotCacheRecordType.n, record.Type);
			Assert.AreEqual("20100076", record.Value);
			node = document.SelectSingleNode("//x");
			record = new CacheRecordItem(node);
			Assert.AreEqual(PivotCacheRecordType.x, record.Type);
			node = document.SelectSingleNode("//m");
			record = new CacheRecordItem(node);
			Assert.AreEqual(PivotCacheRecordType.m, record.Type);
			node = document.SelectSingleNode("//e");
			record = new CacheRecordItem(node);
			Assert.AreEqual(PivotCacheRecordType.e, record.Type);
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void CacheRecordItemNullNodeTest()
		{
			new CacheRecordItem(null);
		}
		#endregion

		#region UpdateValue Tests
		[TestMethod]
		public void UpdateValueDifferentType()
		{
			XmlDocument document = new XmlDocument();
			document.LoadXml(@"<pivotCacheRecords count=""1""><r><n v=""20100076""/></r></pivotCacheRecords>");
			var node = document.SelectSingleNode("//n");
			var record = new CacheRecordItem(node);
			var parentNode = document.SelectSingleNode("//r");
			var cacheFieldNode = base.GetTestCacheFieldNode();
			record.UpdateValue(true, parentNode, cacheFieldNode);
			Assert.AreEqual("True", record.Value);
			Assert.AreEqual(PivotCacheRecordType.b, record.Type);
		}

		[TestMethod]
		public void UpdateValueSameType()
		{
			XmlDocument document = new XmlDocument();
			document.LoadXml(@"<pivotCacheRecords count=""1""><r><n v=""20100076""/></r></pivotCacheRecords>");
			var node = document.SelectSingleNode("//n");
			var record = new CacheRecordItem(node);
			var parentNode = document.SelectSingleNode("//r");
			var cacheFieldNode = base.GetTestCacheFieldNode();
			record.UpdateValue(389471230, parentNode, cacheFieldNode);
			Assert.AreEqual("389471230", record.Value);
			Assert.AreEqual(PivotCacheRecordType.n, record.Type);
		}

		[TestMethod]
		public void UpdateValueWithNewSharedString()
		{
			XmlDocument document = new XmlDocument();
			document.LoadXml(@"<pivotCacheRecords count=""1""><r><n v=""20100076""/></r></pivotCacheRecords>");
			var node = document.SelectSingleNode("//n");
			var record = new CacheRecordItem(node);
			var parentNode = document.SelectSingleNode("//r");
			var cacheFieldNode = base.GetTestCacheFieldNode();
			record.UpdateValue("red", parentNode, cacheFieldNode);
			Assert.AreEqual("2", record.Value);
			Assert.AreEqual(PivotCacheRecordType.x, record.Type);
		}

		[TestMethod]
		public void UpdateValueWithNewSharedStringSameType()
		{
			XmlDocument document = new XmlDocument();
			document.LoadXml(@"<pivotCacheRecords count=""1""><r><x v=""1""/></r></pivotCacheRecords>");
			var node = document.SelectSingleNode("//x");
			var record = new CacheRecordItem(node);
			var parentNode = document.SelectSingleNode("//r");
			var cacheFieldNode = base.GetTestCacheFieldNode();
			record.UpdateValue("red", parentNode, cacheFieldNode);
			Assert.AreEqual("2", record.Value);
			Assert.AreEqual(PivotCacheRecordType.x, record.Type);
		}

		[TestMethod]
		public void UpdateValueWithExistingSharedStringDifferentTypes()
		{
			XmlDocument document = new XmlDocument();
			document.LoadXml(@"<pivotCacheRecords count=""1""><r><n v=""20100076""/></r></pivotCacheRecords>");
			var node = document.SelectSingleNode("//n");
			var record = new CacheRecordItem(node);
			var parentNode = document.SelectSingleNode("//r");
			var cacheFieldNode = base.GetTestCacheFieldNode();
			record.UpdateValue("Car", parentNode, cacheFieldNode);
			Assert.AreEqual("1", record.Value);
			Assert.AreEqual(PivotCacheRecordType.x, record.Type);
		}

		[TestMethod]
		public void UpdateValueWithExistingSharedStringSameTypes()
		{
			XmlDocument document = new XmlDocument();
			document.LoadXml(@"<pivotCacheRecords count=""1""><r><n x=""4""/></r></pivotCacheRecords>");
			var node = document.SelectSingleNode("//n");
			var record = new CacheRecordItem(node);
			var parentNode = document.SelectSingleNode("//r");
			var cacheFieldNode = base.GetTestCacheFieldNode();
			record.UpdateValue("Car", parentNode, cacheFieldNode);
			Assert.AreEqual("1", record.Value);
			Assert.AreEqual(PivotCacheRecordType.x, record.Type);
		}

		[TestMethod]
		public void UpdateValueNullValue()
		{
			XmlDocument document = new XmlDocument();
			document.LoadXml(@"<pivotCacheRecords count=""1""><r><n v=""20100076""/></r></pivotCacheRecords>");
			var node = document.SelectSingleNode("//n");
			var record = new CacheRecordItem(node);
			var parentNode = document.SelectSingleNode("//r");
			var cacheFieldNode = base.GetTestCacheFieldNode();
			record.UpdateValue(null, parentNode, cacheFieldNode);
			Assert.IsNull(record.Value);
			Assert.AreEqual(PivotCacheRecordType.m, record.Type);
		}

		[TestMethod]
		public void UpdateValueEmptyStringValue()
		{
			XmlDocument document = new XmlDocument();
			document.LoadXml(@"<pivotCacheRecords count=""1""><r><n v=""20100076""/></r></pivotCacheRecords>");
			var node = document.SelectSingleNode("//n");
			var record = new CacheRecordItem(node);
			var parentNode = document.SelectSingleNode("//r");
			var cacheFieldNode = base.GetTestCacheFieldNode();
			record.UpdateValue(string.Empty, parentNode, cacheFieldNode);
			Assert.IsNull(record.Value);
			Assert.AreEqual(PivotCacheRecordType.m, record.Type);
		}

		[TestMethod]
		[ExpectedException(typeof(InvalidOperationException))]
		public void UpdateValueNotSupportedType()
		{
			XmlDocument document = new XmlDocument();
			document.LoadXml(@"<pivotCacheRecords count=""1""><r><n v=""20100076""/></r></pivotCacheRecords>");
			var node = document.SelectSingleNode("//n");
			var record = new CacheRecordItem(node);
			var parentNode = document.SelectSingleNode("//r");
			var cacheFieldNode = base.GetTestCacheFieldNode();
			record.UpdateValue(record, parentNode, cacheFieldNode);
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void UpdateValueNullParentNode()
		{
			XmlDocument document = new XmlDocument();
			document.LoadXml(@"<pivotCacheRecords count=""1""><r><n v=""20100076""/></r></pivotCacheRecords>");
			var node = document.SelectSingleNode("//n");
			var record = new CacheRecordItem(node);
			var parentNode = document.SelectSingleNode("//r");
			var cacheFieldNode = base.GetTestCacheFieldNode();
			record.UpdateValue(record, null, cacheFieldNode);
		}
		#endregion
	}
}