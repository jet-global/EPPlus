using System;
using System.Xml;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Slicers;

namespace EPPlusTest.Drawing.Slicers
{
	[TestClass]
	public class SlicerLevelNodeTest
	{
		#region Constructor Tests
		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void SlicerLevelNodeNullXmlNodeThrowsException()
		{
			new SlicerLevelNode(null, ExcelSlicer.SlicerDocumentNamespaceManager);
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void SlicerLevelNodeNullNamespaceManagerThrowsException()
		{
			var node = this.CreateSlicerLevelNode();
			new SlicerLevelNode(node, null);
		}
		#endregion

		#region UniqueName Tests
		[TestMethod]
		public void SlicerLevelNodeUniqueName()
		{
			var node = this.CreateSlicerLevelNode();
			var slicerLevelNode = new SlicerLevelNode(node, ExcelSlicer.SlicerDocumentNamespaceManager);
			Assert.AreEqual("[Bill-to Customer].[by Country by City].[Country]", slicerLevelNode.UniqueName);
			slicerLevelNode.UniqueName = "UniqueName";
			Assert.AreEqual("UniqueName", slicerLevelNode.UniqueName);
			Assert.AreEqual($@"<level uniqueName=""UniqueName"" sourceCaption=""Country"" count=""21"" xmlns=""{ExcelPackage.schemaMain2009}""><ranges><range startItem=""0""><i n=""[Bill-to Customer].[by Country by City].[Country].&amp;[]"" c="""" /><i n=""[Bill-to Customer].[by Country by City].[Country].&amp;[AT]"" c=""Austria"" /><i n=""[Bill-to Customer].[by Country by City].[Country].&amp;[BE]"" c=""Belgium"" /></range><range startItem=""1""><i n=""[Bill-to Customer].[by Country by City].[Country].&amp;[CA]"" c=""Canada"" /><i n=""[Bill-to Customer].[by Country by City].[Country].&amp;[DK]"" c=""Denmark"" /><i n=""[Bill-to Customer].[by Country by City].[Country].&amp;[FR]"" c=""France"" /></range></ranges></level>", node.OuterXml);
		}
		#endregion

		#region SourceCaption Tests
		[TestMethod]
		public void SlicerLevelNodeSourceCaption()
		{
			var node = this.CreateSlicerLevelNode();
			var slicerLevelNode = new SlicerLevelNode(node, ExcelSlicer.SlicerDocumentNamespaceManager);
			Assert.AreEqual("Country", slicerLevelNode.SourceCaption);
			slicerLevelNode.SourceCaption = "SourceCaption";
			Assert.AreEqual("SourceCaption", slicerLevelNode.SourceCaption);
			Assert.AreEqual($@"<level uniqueName=""[Bill-to Customer].[by Country by City].[Country]"" sourceCaption=""SourceCaption"" count=""21"" xmlns=""{ExcelPackage.schemaMain2009}""><ranges><range startItem=""0""><i n=""[Bill-to Customer].[by Country by City].[Country].&amp;[]"" c="""" /><i n=""[Bill-to Customer].[by Country by City].[Country].&amp;[AT]"" c=""Austria"" /><i n=""[Bill-to Customer].[by Country by City].[Country].&amp;[BE]"" c=""Belgium"" /></range><range startItem=""1""><i n=""[Bill-to Customer].[by Country by City].[Country].&amp;[CA]"" c=""Canada"" /><i n=""[Bill-to Customer].[by Country by City].[Country].&amp;[DK]"" c=""Denmark"" /><i n=""[Bill-to Customer].[by Country by City].[Country].&amp;[FR]"" c=""France"" /></range></ranges></level>", node.OuterXml);
		}
		#endregion

		#region Count Tests
		[TestMethod]
		public void SlicerLevelNodeCount()
		{
			var node = this.CreateSlicerLevelNode();
			var slicerLevelNode = new SlicerLevelNode(node, ExcelSlicer.SlicerDocumentNamespaceManager);
			Assert.AreEqual("21", slicerLevelNode.Count);
			slicerLevelNode.Count = "20";
			Assert.AreEqual("20", slicerLevelNode.Count);
			Assert.AreEqual($@"<level uniqueName=""[Bill-to Customer].[by Country by City].[Country]"" sourceCaption=""Country"" count=""20"" xmlns=""{ExcelPackage.schemaMain2009}""><ranges><range startItem=""0""><i n=""[Bill-to Customer].[by Country by City].[Country].&amp;[]"" c="""" /><i n=""[Bill-to Customer].[by Country by City].[Country].&amp;[AT]"" c=""Austria"" /><i n=""[Bill-to Customer].[by Country by City].[Country].&amp;[BE]"" c=""Belgium"" /></range><range startItem=""1""><i n=""[Bill-to Customer].[by Country by City].[Country].&amp;[CA]"" c=""Canada"" /><i n=""[Bill-to Customer].[by Country by City].[Country].&amp;[DK]"" c=""Denmark"" /><i n=""[Bill-to Customer].[by Country by City].[Country].&amp;[FR]"" c=""France"" /></range></ranges></level>", node.OuterXml);
		}
		#endregion

		#region SlicerRanges Tests
		[TestMethod]
		public void SlicerLevelNodeSlicerRanges()
		{
			var node = this.CreateSlicerLevelNode();
			var slicerLevelNode = new SlicerLevelNode(node, ExcelSlicer.SlicerDocumentNamespaceManager);
			Assert.AreEqual(2, slicerLevelNode.SlicerRanges.Count);
			Assert.AreEqual("0", slicerLevelNode.SlicerRanges[0].StartItem);
			Assert.AreEqual(3, slicerLevelNode.SlicerRanges[0].Items.Count);
			Assert.AreEqual(string.Empty, slicerLevelNode.SlicerRanges[0].Items[0].DisplayName);
			Assert.AreEqual("Austria", slicerLevelNode.SlicerRanges[0].Items[1].DisplayName);
			Assert.AreEqual("Belgium", slicerLevelNode.SlicerRanges[0].Items[2].DisplayName);
			Assert.AreEqual("1", slicerLevelNode.SlicerRanges[1].StartItem);
			Assert.AreEqual(3, slicerLevelNode.SlicerRanges[1].Items.Count);
			Assert.AreEqual("Canada", slicerLevelNode.SlicerRanges[1].Items[0].DisplayName);
			Assert.AreEqual("Denmark", slicerLevelNode.SlicerRanges[1].Items[1].DisplayName);
			Assert.AreEqual("France", slicerLevelNode.SlicerRanges[1].Items[2].DisplayName);
		}
		#endregion

		#region Helper Methods
		private XmlNode CreateSlicerLevelNode()
		{
			var xmlDoc = new XmlDocument(ExcelSlicer.SlicerDocumentNamespaceManager.NameTable);
			xmlDoc.LoadXml(
			$@"<level uniqueName=""[Bill-to Customer].[by Country by City].[Country]"" sourceCaption=""Country"" count=""21"" xmlns=""{ExcelPackage.schemaMain2009}"">
				<ranges>
					<range startItem=""0"">
						<i n=""[Bill-to Customer].[by Country by City].[Country].&amp;[]"" c=""""/>
						<i n=""[Bill-to Customer].[by Country by City].[Country].&amp;[AT]"" c=""Austria""/>
						<i n=""[Bill-to Customer].[by Country by City].[Country].&amp;[BE]"" c=""Belgium""/>
					</range>
					<range startItem=""1"">
						<i n=""[Bill-to Customer].[by Country by City].[Country].&amp;[CA]"" c=""Canada""/>
						<i n=""[Bill-to Customer].[by Country by City].[Country].&amp;[DK]"" c=""Denmark""/>
						<i n=""[Bill-to Customer].[by Country by City].[Country].&amp;[FR]"" c=""France""/>
					</range>
				</ranges>
			</level>");
			return xmlDoc.FirstChild;
		}
		#endregion
	}
}
