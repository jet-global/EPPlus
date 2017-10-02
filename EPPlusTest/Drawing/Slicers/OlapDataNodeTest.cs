using System;
using System.Xml;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Slicers;

namespace EPPlusTest.Drawing.Slicers
{
	[TestClass]
	public class OlapDataNodeTest
	{
		#region Constructor Tests
		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void OlapDataNodeNullXmlNodeThrowsException()
		{
			new OlapDataNode(null, ExcelSlicer.SlicerDocumentNamespaceManager);
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void OlapDataNodeNullNamespaceManagerThrowsException()
		{
			var node = this.CreateOlapDataNode();
			new OlapDataNode(node, null);
		}
		#endregion

		#region PivotCacheId Tests
		[TestMethod]
		public void OlapDataNodePivotCacheId()
		{
			var node = this.CreateOlapDataNode();
			var olapDataNode = new OlapDataNode(node, ExcelSlicer.SlicerDocumentNamespaceManager);
			Assert.AreEqual("86", olapDataNode.PivotCacheId);
			olapDataNode.PivotCacheId = "68";
			Assert.AreEqual("68", olapDataNode.PivotCacheId);
			Assert.AreEqual($@"<olap pivotCacheId=""68"" xmlns=""{ExcelPackage.schemaMain2009}""><levels count=""4""><level uniqueName=""[Bill-to Customer].[by Country by City].[(All)]"" sourceCaption=""(All)"" count=""0"" /><level uniqueName=""[Bill-to Customer].[by Country by City].[Country]"" sourceCaption=""Country"" count=""21""><ranges><range startItem=""0""><i n=""[Bill-to Customer].[by Country by City].[Country].&amp;[]"" c="""" /><i n=""[Bill-to Customer].[by Country by City].[Country].&amp;[AT]"" c=""Austria"" /><i n=""[Bill-to Customer].[by Country by City].[Country].&amp;[BE]"" c=""Belgium"" /></range></ranges></level><level uniqueName=""[Bill-to Customer].[by Country by City].[City]"" sourceCaption=""City"" count=""0"" /><level uniqueName=""[Bill-to Customer].[by Country by City].[Customer]"" sourceCaption=""Customer"" count=""0"" /></levels><selections count=""1""><selection n=""[Bill-to Customer].[by Country by City].[All Customer]"" /></selections></olap>", node.OuterXml);
		}
		#endregion

		#region SlicerLevels Tests
		[TestMethod]
		public void OlapDataNodeSlicerLevels()
		{
			var node = this.CreateOlapDataNode();
			var olapDataNode = new OlapDataNode(node, ExcelSlicer.SlicerDocumentNamespaceManager);
			Assert.AreEqual(4, olapDataNode.SlicerLevels.Count);
			Assert.AreEqual("(All)", olapDataNode.SlicerLevels[0].SourceCaption);
			Assert.AreEqual("Country", olapDataNode.SlicerLevels[1].SourceCaption);
			Assert.AreEqual(1, olapDataNode.SlicerLevels[1].SlicerRanges.Count);
			Assert.AreEqual(string.Empty, olapDataNode.SlicerLevels[1].SlicerRanges[0].Items[0].DisplayName);
			Assert.AreEqual("Austria", olapDataNode.SlicerLevels[1].SlicerRanges[0].Items[1].DisplayName);
			Assert.AreEqual("Belgium", olapDataNode.SlicerLevels[1].SlicerRanges[0].Items[2].DisplayName);
			Assert.AreEqual("City", olapDataNode.SlicerLevels[2].SourceCaption);
			Assert.AreEqual("Customer", olapDataNode.SlicerLevels[3].SourceCaption);
		}
		#endregion

		#region Selections Tests
		[TestMethod]
		public void OlapDataNodeSelections()
		{
			var node = this.CreateOlapDataNode();
			var olapDataNode = new OlapDataNode(node, ExcelSlicer.SlicerDocumentNamespaceManager);
			Assert.AreEqual(1, olapDataNode.Selections.Count);
			Assert.AreEqual("[Bill-to Customer].[by Country by City].[All Customer]", olapDataNode.Selections[0].Name);
		}
		#endregion

		#region Helper Methods
		private XmlNode CreateOlapDataNode()
		{
			var xmlDoc = new XmlDocument(ExcelSlicer.SlicerDocumentNamespaceManager.NameTable);
			xmlDoc.LoadXml(
			$@"<olap pivotCacheId=""86"" xmlns=""{ExcelPackage.schemaMain2009}"">
				<levels count=""4"">
					<level uniqueName=""[Bill-to Customer].[by Country by City].[(All)]"" sourceCaption=""(All)"" count=""0""/>
					<level uniqueName=""[Bill-to Customer].[by Country by City].[Country]"" sourceCaption=""Country"" count=""21"">
						<ranges>
							<range startItem=""0"">
								<i n=""[Bill-to Customer].[by Country by City].[Country].&amp;[]"" c=""""/>
								<i n=""[Bill-to Customer].[by Country by City].[Country].&amp;[AT]"" c=""Austria""/>
								<i n=""[Bill-to Customer].[by Country by City].[Country].&amp;[BE]"" c=""Belgium""/>
							</range>
						</ranges>
					</level>
					<level uniqueName=""[Bill-to Customer].[by Country by City].[City]"" sourceCaption=""City"" count=""0""/>
					<level uniqueName=""[Bill-to Customer].[by Country by City].[Customer]"" sourceCaption=""Customer"" count=""0""/>
				</levels>
				<selections count=""1"">
					<selection n=""[Bill-to Customer].[by Country by City].[All Customer]""/>
				</selections>
			</olap>");
			return xmlDoc.FirstChild;
		}
		#endregion
	}
}
