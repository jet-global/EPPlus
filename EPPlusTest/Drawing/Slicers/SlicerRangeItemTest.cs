using System;
using System.Xml;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
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
			new SlicerRangeItem(null, ExcelSlicer.SlicerDocumentNamespaceManager);
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void SlicerRangeItemNullXmlNamespaceManagerThrowsException()
		{
			var xmlDoc = new XmlDocument(ExcelSlicer.SlicerDocumentNamespaceManager.NameTable);
			xmlDoc.LoadXml("<i n=\"[Bill-to Customer].[by Country by City].[Country].&amp;[AT]\" c=\"Austria\"/>");
			var node = xmlDoc.FirstChild;
			new SlicerRangeItem(node, null);
		}
		#endregion

		#region Name Tests
		[TestMethod]
		public void SlicerRangeItemName()
		{
			var xmlDoc = new XmlDocument(ExcelSlicer.SlicerDocumentNamespaceManager.NameTable);
			xmlDoc.LoadXml("<i n=\"[Bill-to Customer].[by Country by City].[Country].&amp;[AT]\" c=\"Austria\"/>");
			var node = xmlDoc.FirstChild;
			var slicerRangeItem = new SlicerRangeItem(node, ExcelSlicer.SlicerDocumentNamespaceManager);
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
			var slicerRangeItem = new SlicerRangeItem(node, ExcelSlicer.SlicerDocumentNamespaceManager);
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
			var slicerRangeItem = new SlicerRangeItem(node, ExcelSlicer.SlicerDocumentNamespaceManager);
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
			var slicerRangeItem = new SlicerRangeItem(node, ExcelSlicer.SlicerDocumentNamespaceManager);
			Assert.IsFalse(slicerRangeItem.NonDisplay);
			slicerRangeItem.NonDisplay = true;
			Assert.IsTrue(slicerRangeItem.NonDisplay);
			Assert.AreEqual("<i n=\"[Bill-to Customer].[by Country by City].[Country].&amp;[AT]\" c=\"Austria\" nd=\"1\" />", node.OuterXml);
		}
		#endregion

		#region Parents Tests
		[TestMethod]
		public void SlicerRangeItemLoadsParentsIfTheyExist()
		{
			var xmlDoc = new XmlDocument(ExcelSlicer.SlicerDocumentNamespaceManager.NameTable);
			xmlDoc.LoadXml($@"
				<i n=""[Bill-to Customer].[by Country by City].[Country].&amp;[AT]"" c=""Austria"" xmlns=""{ExcelPackage.schemaMain2009}"">
					<p n=""parent1"" />
					<p n=""parent2"" />
				</i>");
			var node = xmlDoc.FirstChild;
			var slicerRangeItem = new SlicerRangeItem(node, ExcelSlicer.SlicerDocumentNamespaceManager);
			Assert.AreEqual(2, slicerRangeItem.Parents.Count);
			Assert.AreEqual("parent1", slicerRangeItem.Parents[0].Name);
			Assert.AreEqual("parent2", slicerRangeItem.Parents[1].Name);
		}
		#endregion
	}
}
