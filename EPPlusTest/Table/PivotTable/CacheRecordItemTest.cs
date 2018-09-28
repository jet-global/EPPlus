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
			var record = new CacheRecordItem(node, TestUtility.CreateDefaultNSM());
			Assert.AreEqual(PivotCacheRecordType.n, record.Type);
			Assert.AreEqual("20100076", record.Value);
			node = document.SelectSingleNode("//x");
			record = new CacheRecordItem(node, TestUtility.CreateDefaultNSM());
			Assert.AreEqual(PivotCacheRecordType.x, record.Type);
			node = document.SelectSingleNode("//m");
			record = new CacheRecordItem(node, TestUtility.CreateDefaultNSM());
			Assert.AreEqual(PivotCacheRecordType.m, record.Type);
			node = document.SelectSingleNode("//e");
			record = new CacheRecordItem(node, TestUtility.CreateDefaultNSM());
			Assert.AreEqual(PivotCacheRecordType.e, record.Type);
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void CacheRecordItemNullNodeTest()
		{
			new CacheRecordItem(null, TestUtility.CreateDefaultNSM());
		}
		#endregion
	}
}