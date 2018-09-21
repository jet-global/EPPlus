using System;
using System.Linq;
using System.Xml;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Table.PivotTable;

namespace EPPlusTest.Table.PivotTable
{
	[TestClass]
	public class CacheRecordNodeTest
	{
		#region Constructor Tests
		[TestMethod]
		public void CacheRecordNodeConstructorTest()
		{
			XmlDocument document = new XmlDocument();
			document.LoadXml(@"<pivotCacheRecords xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" count=""1""><r><n v=""20100076""/><x v=""0""/> <b v=""0""/> <m v=""0""/> <e v=""415.75""/><d v=""1""/></r></pivotCacheRecords>");
			var ns = TestUtility.CreateDefaultNSM();
			var node = new CacheRecordNode(document.SelectSingleNode("//d:r", ns), ns);
			Assert.AreEqual(6, node.Items.Count);
			Assert.AreEqual(1, node.Items.Count(i => i.Type == PivotCacheRecordType.b));
			Assert.AreEqual(1, node.Items.Count(i => i.Type == PivotCacheRecordType.x));
			Assert.AreEqual(1, node.Items.Count(i => i.Type == PivotCacheRecordType.d));
			Assert.AreEqual(1, node.Items.Count(i => i.Type == PivotCacheRecordType.e));
			Assert.AreEqual(1, node.Items.Count(i => i.Type == PivotCacheRecordType.m));
			Assert.AreEqual(1, node.Items.Count(i => i.Type == PivotCacheRecordType.n));
			Assert.AreEqual(0, node.Items.Count(i => i.Type == PivotCacheRecordType.s));
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void CacheFieldNodeTestNullNode()
		{
			new CacheFieldNode(null, TestUtility.CreateDefaultNSM());
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void CacheFieldNodeTestNullNamespaceManager()
		{
			var xml = new XmlDocument();
			xml.LoadXml(@"<pivotCacheRecords count=""1""><r><n v=""20100076""/></r></pivotCacheRecords>");
			new CacheFieldNode(xml.FirstChild, null);
		}
		#endregion
	}
}