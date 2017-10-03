using System;
using System.Xml;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Drawing.Slicers;

namespace EPPlusTest.Drawing.Slicers
{
	[TestClass]
	public class SlicerCacheItemParentTest
	{
		#region Constructor Tests
		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void SlicerSelectionNodeNullXmlNodeThrowsException()
		{
			new SlicerCacheItemParent(null);
		}
		#endregion

		#region Name Tests
		[TestMethod]
		public void SlicerSelectionName()
		{
			var xmlDoc = new XmlDocument(ExcelSlicer.SlicerDocumentNamespaceManager.NameTable);
			xmlDoc.LoadXml("<p n=\"SlicerParent\" />");
			var node = xmlDoc.FirstChild;
			var slicerCacheItemParent = new SlicerCacheItemParent(node);
			Assert.AreEqual("SlicerParent", slicerCacheItemParent.Name);
			slicerCacheItemParent.Name = "SlicerParentUpdated";
			Assert.AreEqual("SlicerParentUpdated", slicerCacheItemParent.Name);
			Assert.AreEqual("<p n=\"SlicerParentUpdated\" />", node.OuterXml);
		}
		#endregion
	}
}
