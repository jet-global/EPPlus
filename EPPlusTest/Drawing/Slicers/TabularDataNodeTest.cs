using System;
using System.Xml;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Slicers;

namespace EPPlusTest.Drawing.Slicers
{
	[TestClass]
	public class TabularDataNodeTest
	{
		#region Constructor Tests
		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void TabularDataNodeNullXmlNodeThrowsException()
		{
			new TabularDataNode(null, ExcelSlicer.SlicerDocumentNamespaceManager);
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void TabularDataNodeNullNamespaceManagerThrowsException()
		{
			var node = this.CreateTabularDataNode();
			new TabularDataNode(node, null);
		}
		#endregion

		#region PivotCacheId Tests
		[TestMethod]
		public void TabularDataNodePivotCacheId()
		{
			var node = this.CreateTabularDataNode();
			var tabularDataNode = new TabularDataNode(node, ExcelSlicer.SlicerDocumentNamespaceManager);
			Assert.AreEqual("1", tabularDataNode.PivotCacheId);
			tabularDataNode.PivotCacheId = "2";
			Assert.AreEqual("2", tabularDataNode.PivotCacheId);
			Assert.AreEqual($@"<tabular pivotCacheId=""2"" xmlns=""{ExcelPackage.schemaMain2009}""><items count=""4""><i x=""0"" /><i x=""1"" /><i x=""2"" s=""1"" /><i x=""3"" s=""1"" /></items></tabular>", node.OuterXml);
		}
		#endregion

		#region Items Tests
		[TestMethod]
		public void TabularDataNodeItems()
		{
			var node = this.CreateTabularDataNode();
			var tabularDataNode = new TabularDataNode(node, ExcelSlicer.SlicerDocumentNamespaceManager);
			Assert.AreEqual(4, tabularDataNode.Items.Count);
			Assert.AreEqual(0, tabularDataNode.Items[0].AtomIndex);
			Assert.AreEqual(false, tabularDataNode.Items[0].IsSelected);
			Assert.AreEqual(1, tabularDataNode.Items[1].AtomIndex);
			Assert.AreEqual(false, tabularDataNode.Items[1].IsSelected);
			Assert.AreEqual(2, tabularDataNode.Items[2].AtomIndex);
			Assert.AreEqual(true, tabularDataNode.Items[2].IsSelected);
			Assert.AreEqual(3, tabularDataNode.Items[3].AtomIndex);
			Assert.AreEqual(true, tabularDataNode.Items[3].IsSelected);
		}
		#endregion

		#region Helper Methods
		private XmlNode CreateTabularDataNode()
		{
			var xmlDoc = new XmlDocument(ExcelSlicer.SlicerDocumentNamespaceManager.NameTable);
			xmlDoc.LoadXml(
			$@"<tabular pivotCacheId=""1"" xmlns=""{ExcelPackage.schemaMain2009}"">
				<items count=""4"">
					<i x=""0""/>
					<i x=""1""/>
					<i x=""2"" s=""1""/>
					<i x=""3"" s=""1""/>
				</items>
			</tabular>");
			return xmlDoc.FirstChild;
		}
		#endregion
	}
}
