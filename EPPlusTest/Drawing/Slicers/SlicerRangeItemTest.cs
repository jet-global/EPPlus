using System;
using System.Xml;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Drawing.Slicers;

namespace EPPlusTest.Drawing.Slicers
{
	[TestClass]
	public class SlicerRangeItemTest
	{
		#region Constructor Tests
		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void SlicerRangeItemNullXmlNodeThrowsException()
		{
			new SlicerRangeItem(null);
		}
		#endregion

		#region Name Tests
		[TestMethod]
		public void SlicerRangeItemName()
		{
			var xmlDoc = new XmlDocument(ExcelSlicer.SlicerDocumentNamespaceManager.NameTable);
			xmlDoc.LoadXml("<i n=\"[Bill-to Customer].[by Country by City].[Country].&amp;[AT]\" c=\"Austria\"/>");
			var node = xmlDoc.FirstChild;
			var slicerRangeItem = new SlicerRangeItem(node);
			Assert.AreEqual("[Bill-to Customer].[by Country by City].[Country].&[AT]", slicerRangeItem.Name);
			slicerRangeItem.Name = "[Bill-to Customer].[by Country by City].[Country].&[BE]";
			Assert.AreEqual("[Bill-to Customer].[by Country by City].[Country].&[BE]", slicerRangeItem.Name);
			Assert.AreEqual("<i n=\"[Bill-to Customer].[by Country by City].[Country].&amp;[BE]\" c=\"Austria\" />", node.OuterXml);
		}
		#endregion

		#region DisplayName Tests
		[TestMethod]
		public void SlicerRangeItemDisplayName()
		{
			var xmlDoc = new XmlDocument(ExcelSlicer.SlicerDocumentNamespaceManager.NameTable);
			xmlDoc.LoadXml("<i n=\"[Bill-to Customer].[by Country by City].[Country].&amp;[AT]\" c=\"Austria\"/>");
			var node = xmlDoc.FirstChild;
			var slicerRangeItem = new SlicerRangeItem(node);
			Assert.AreEqual("Austria", slicerRangeItem.DisplayName);
			slicerRangeItem.DisplayName = "Belgium";
			Assert.AreEqual("Belgium", slicerRangeItem.DisplayName);
			Assert.AreEqual("<i n=\"[Bill-to Customer].[by Country by City].[Country].&amp;[AT]\" c=\"Belgium\" />", node.OuterXml);
		}
		#endregion

		#region NonDisplay Tests
		[TestMethod]
		public void SlicerRangeItemNonDisplay()
		{
			var xmlDoc = new XmlDocument(ExcelSlicer.SlicerDocumentNamespaceManager.NameTable);
			xmlDoc.LoadXml("<i n=\"[Bill-to Customer].[by Country by City].[Country].&amp;[AT]\" c=\"Austria\" nd=\"1\"/>");
			var node = xmlDoc.FirstChild;
			var slicerRangeItem = new SlicerRangeItem(node);
			Assert.IsTrue(slicerRangeItem.NonDisplay);
			slicerRangeItem.NonDisplay = false;
			Assert.IsFalse(slicerRangeItem.NonDisplay);
			Assert.AreEqual("<i n=\"[Bill-to Customer].[by Country by City].[Country].&amp;[AT]\" c=\"Austria\" nd=\"0\" />", node.OuterXml);
		}

		[TestMethod]
		public void SlicerRangeItemNonDisplayDoesNotExist()
		{
			var xmlDoc = new XmlDocument(ExcelSlicer.SlicerDocumentNamespaceManager.NameTable);
			xmlDoc.LoadXml("<i n=\"[Bill-to Customer].[by Country by City].[Country].&amp;[AT]\" c=\"Austria\"/>");
			var node = xmlDoc.FirstChild;
			var slicerRangeItem = new SlicerRangeItem(node);
			Assert.IsFalse(slicerRangeItem.NonDisplay);
			slicerRangeItem.NonDisplay = true;
			Assert.IsTrue(slicerRangeItem.NonDisplay);
			Assert.AreEqual("<i n=\"[Bill-to Customer].[by Country by City].[Country].&amp;[AT]\" c=\"Austria\" nd=\"1\" />", node.OuterXml);
		}
		#endregion
	}
}
