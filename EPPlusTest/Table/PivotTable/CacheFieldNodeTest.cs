using System;
using System.Xml;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Table.PivotTable;

namespace EPPlusTest.Table.PivotTable
{
	[TestClass]
	public class CacheFieldNodeTest
	{
		#region Constructor Tests
		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void SlicerRangeItemNullXmlNodeThrowsException()
		{
			new CacheFieldNode(null, TestUtility.CreateDefaultNSM());
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void SlicerRangeItemNullNamespaceManagerThrowsException()
		{
			new CacheFieldNode(this.CreateCacheFieldNode(), null);
		}
		#endregion

		#region Name Tests
		[TestMethod]
		public void SlicerRangeNodeName()
		{
			var node = this.CreateCacheFieldNode();
			var cacheFieldNode = new CacheFieldNode(node, TestUtility.CreateDefaultNSM());
			Assert.AreEqual("Customer No.", cacheFieldNode.Name);
			cacheFieldNode.Name = "Other Name";
			Assert.AreEqual("Other Name", cacheFieldNode.Name);
			Assert.AreEqual($@"<cacheField name=""Other Name"" numFmtId=""49"" xmlns=""{ExcelPackage.schemaMain}""><sharedItems count=""3""><s v=""10000"" /><s v=""20000"" /><s v=""30000"" /></sharedItems></cacheField>", node.OuterXml);
		}
		#endregion

		#region NumFormatId Tests
		[TestMethod]
		public void SlicerRangeNodeNumFormatId()
		{
			var node = this.CreateCacheFieldNode();
			var cacheFieldNode = new CacheFieldNode(node, TestUtility.CreateDefaultNSM());
			Assert.AreEqual("49", cacheFieldNode.NumFormatId);
			cacheFieldNode.NumFormatId = "30";
			Assert.AreEqual("30", cacheFieldNode.NumFormatId);
			Assert.AreEqual($@"<cacheField name=""Customer No."" numFmtId=""30"" xmlns=""{ExcelPackage.schemaMain}""><sharedItems count=""3""><s v=""10000"" /><s v=""20000"" /><s v=""30000"" /></sharedItems></cacheField>", node.OuterXml);
		}
		#endregion

		#region Items Tests
		[TestMethod]
		public void SlicerRangeNodeItems()
		{
			var node = this.CreateCacheFieldNode();
			var cacheFieldNode = new CacheFieldNode(node, TestUtility.CreateDefaultNSM());
			Assert.AreEqual(3, cacheFieldNode.Items.Count);
			Assert.AreEqual("10000", cacheFieldNode.Items[0].Value);
			Assert.AreEqual("20000", cacheFieldNode.Items[1].Value);
			Assert.AreEqual("30000", cacheFieldNode.Items[2].Value);
		}
		#endregion

		#region Helper Methods
		private XmlNode CreateCacheFieldNode()
		{
			var xmlDoc = new XmlDocument(TestUtility.CreateDefaultNSM().NameTable);
			xmlDoc.LoadXml(
			$@"<cacheField name=""Customer No."" numFmtId=""49"" xmlns=""{ExcelPackage.schemaMain}"">
				<sharedItems count=""3"">
					<s v=""10000""/>
					<s v=""20000""/>
					<s v=""30000""/>
				</sharedItems>
			</cacheField>");
			return xmlDoc.FirstChild;
		}
		#endregion
	}
}