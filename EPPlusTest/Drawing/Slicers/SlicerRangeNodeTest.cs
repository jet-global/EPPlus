using System;
using System.Xml;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Slicers;

namespace EPPlusTest.Drawing.Slicers
{
	[TestClass]
	public class SlicerRangeNodeTest
	{
		#region Constructor Tests
		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void SlicerRangeNodeNullXmlNodeThrowsException()
		{
			new SlicerRangeNode(null, ExcelSlicer.SlicerDocumentNamespaceManager);
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void SlicerRangeNodeNullNamespaceManagerThrowsException()
		{
			var node = this.CreateSlicerRangeNode();
			new SlicerRangeNode(node, null);
		}
		#endregion

		#region StartItem Tests
		[TestMethod]
		public void SlicerRangeNodeStartItem()
		{
			var node = this.CreateSlicerRangeNode();
			var slicerRangeNode = new SlicerRangeNode(node, ExcelSlicer.SlicerDocumentNamespaceManager);
			Assert.AreEqual("0", slicerRangeNode.StartItem);
			slicerRangeNode.StartItem = "1";
			Assert.AreEqual("1", slicerRangeNode.StartItem);
			Assert.AreEqual($@"<range startItem=""1"" xmlns=""{ExcelPackage.schemaMain2009}""><i n=""[Bill-to Customer].[by Country by City].[Country].&amp;[]"" c="""" /><i n=""[Bill-to Customer].[by Country by City].[Country].&amp;[AT]"" c=""Austria"" /><i n=""[Bill-to Customer].[by Country by City].[Country].&amp;[BE]"" c=""Belgium"" /></range>", node.OuterXml);
		}
		#endregion

		#region Items Tests
		[TestMethod]
		public void SlicerRangeNodeItems()
		{
			var node = this.CreateSlicerRangeNode();
			var slicerRangeNode = new SlicerRangeNode(node, ExcelSlicer.SlicerDocumentNamespaceManager);
			Assert.AreEqual(3, slicerRangeNode.Items.Count);
			Assert.AreEqual(string.Empty, slicerRangeNode.Items[0].DisplayName);
			Assert.AreEqual("Austria", slicerRangeNode.Items[1].DisplayName);
			Assert.AreEqual("Belgium", slicerRangeNode.Items[2].DisplayName);
		}
		#endregion

		#region Helper Methods
		private XmlNode CreateSlicerRangeNode()
		{
			var xmlDoc = new XmlDocument(ExcelSlicer.SlicerDocumentNamespaceManager.NameTable);
			xmlDoc.LoadXml(
			$@"<range startItem=""0"" xmlns=""{ExcelPackage.schemaMain2009}"">
				<i n=""[Bill-to Customer].[by Country by City].[Country].&amp;[]"" c=""""/>
				<i n=""[Bill-to Customer].[by Country by City].[Country].&amp;[AT]"" c=""Austria""/>
				<i n=""[Bill-to Customer].[by Country by City].[Country].&amp;[BE]"" c=""Belgium""/>
			</range>");
			return xmlDoc.FirstChild;
		}
		#endregion
	}
}
