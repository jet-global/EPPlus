using System;
using System.Xml;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Slicers;

namespace EPPlusTest.Drawing.Slicers
{
	[TestClass]
	public class SlicerSelectionNodeTest
	{
		#region Constructor Tests
		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void SlicerSelectionNodeNullXmlNodeThrowsException()
		{
			new SlicerSelectionNode(null, ExcelSlicer.SlicerDocumentNamespaceManager);
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void SlicerSelectionNodeNullXmlNamespaceManagerThrowsException()
		{
			var xmlDoc = new XmlDocument(ExcelSlicer.SlicerDocumentNamespaceManager.NameTable);
			xmlDoc.LoadXml("<selection n=\"SlicerSelectionValue\" />");
			var node = xmlDoc.FirstChild;
			new SlicerSelectionNode(node, null);
		}
		#endregion

		#region Name Tests
		[TestMethod]
		public void SlicerSelectionName()
		{
			var xmlDoc = new XmlDocument(ExcelSlicer.SlicerDocumentNamespaceManager.NameTable);
			xmlDoc.LoadXml("<selection n=\"SlicerSelectionValue\" />");
			var node = xmlDoc.FirstChild;
			var slicerSelectionNode = new SlicerSelectionNode(node, ExcelSlicer.SlicerDocumentNamespaceManager);
			Assert.AreEqual("SlicerSelectionValue", slicerSelectionNode.Name);
			slicerSelectionNode.Name = "SlicerSelectionValueUpdated";
			Assert.AreEqual("SlicerSelectionValueUpdated", slicerSelectionNode.Name);
			Assert.AreEqual("<selection n=\"SlicerSelectionValueUpdated\" />", node.OuterXml);
		}
		#endregion

		#region Parents Tests
		[TestMethod]
		public void SlicerSelectionLoadsParentsIfTheyExist()
		{
			var xmlDoc = new XmlDocument(ExcelSlicer.SlicerDocumentNamespaceManager.NameTable);
			xmlDoc.LoadXml($@"
				<selection n=""SlicerSelectionValue"" xmlns=""{ExcelPackage.schemaMain2009}"">
					<p n=""parent1"" />
					<p n=""parent2"" />
				</selection>");
			var node = xmlDoc.FirstChild;
			var slicerSelectionNode = new SlicerSelectionNode(node, ExcelSlicer.SlicerDocumentNamespaceManager);
			Assert.AreEqual(2, slicerSelectionNode.Parents.Count);
			Assert.AreEqual("parent1", slicerSelectionNode.Parents[0].Name);
			Assert.AreEqual("parent2", slicerSelectionNode.Parents[1].Name);
		}
		#endregion
	}
}
