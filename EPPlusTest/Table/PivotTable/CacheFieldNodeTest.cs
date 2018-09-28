using System;
using System.Xml;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Table.PivotTable;

namespace EPPlusTest.Table.PivotTable
{
	[TestClass]
	public class CacheFieldNodeTest : PivotTableTestBase
	{
		#region Constructor Tests
		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void SlicerRangeItemNullXmlNodeThrowsException()
		{
			new CacheFieldNode(TestUtility.CreateDefaultNSM(), null);
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void SlicerRangeItemNullNamespaceManagerThrowsException()
		{
			new CacheFieldNode(null, this.CreateCacheFieldNode());
		}
		#endregion

		#region Name Tests
		[TestMethod]
		public void SlicerRangeNodeName()
		{
			var node = this.CreateCacheFieldNode();
			var cacheFieldNode = new CacheFieldNode(TestUtility.CreateDefaultNSM(), node);
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
			var cacheFieldNode = new CacheFieldNode(TestUtility.CreateDefaultNSM(), node);
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
			var cacheFieldNode = new CacheFieldNode(TestUtility.CreateDefaultNSM(), node);
			Assert.AreEqual(3, cacheFieldNode.SharedItems.Count);
			Assert.AreEqual("10000", cacheFieldNode.SharedItems.Items[0].Value);
			Assert.AreEqual("20000", cacheFieldNode.SharedItems.Items[1].Value);
			Assert.AreEqual("30000", cacheFieldNode.SharedItems.Items[2].Value);
		}
		#endregion

		#region GetSharedItemIndex Tests
		[TestMethod]
		public void GetSharedItemIndexTest()
		{
			var node = base.GetTestCacheFieldNode();
			var index = node.GetSharedItemIndex(PivotCacheRecordType.s, "Car");
			Assert.AreEqual(1, index);
		}

		[TestMethod]
		public void GetSharedItemIndexSameValueDifferentType()
		{
			var node = base.GetTestCacheFieldNode();
			var index = node.GetSharedItemIndex(PivotCacheRecordType.x, "Car");
			Assert.AreEqual(-1, index);
		}

		[TestMethod]
		public void GetSharedItemIndexValueNotFound()
		{
			var node = base.GetTestCacheFieldNode();
			var index = node.GetSharedItemIndex(PivotCacheRecordType.s, "Mountain");
			Assert.AreEqual(-1, index);
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