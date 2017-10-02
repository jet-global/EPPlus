using System;
using System.Xml;
using Microsoft.VisualStudio.TestTools.UnitTesting;
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
			new SlicerSelectionNode(null);
		}
		#endregion

		#region Name Tests
		[TestMethod]
		public void SlicerSelectionName()
		{
			var xmlDoc = new XmlDocument(ExcelSlicer.SlicerDocumentNamespaceManager.NameTable);
			xmlDoc.LoadXml("<selection n=\"SlicerSelectionValue\" />");
			var node = xmlDoc.FirstChild;
			var slicerSelectionNode = new SlicerSelectionNode(node);
			Assert.AreEqual("SlicerSelectionValue", slicerSelectionNode.Name);
			slicerSelectionNode.Name = "SlicerSelectionValueUpdated";
			Assert.AreEqual("SlicerSelectionValueUpdated", slicerSelectionNode.Name);
			Assert.AreEqual("<selection n=\"SlicerSelectionValueUpdated\" />", node.OuterXml);
		}
		#endregion
	}
}
