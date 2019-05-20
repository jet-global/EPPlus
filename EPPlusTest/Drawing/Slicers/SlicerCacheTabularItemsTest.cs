using System;
using System.Xml;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Drawing.Slicers;

namespace EPPlusTest.Drawing.Slicers
{
	[TestClass]
	public class SlicerCacheTabularItemsTest
	{
		#region Constructor Tests
		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void ConstructSlicerCacheTabularItemNullNodeThrowsException()
		{
			var slicerCacheNamespaceManager = ExcelSlicer.SlicerDocumentNamespaceManager;
			new SlicerCacheTabularItems(null, slicerCacheNamespaceManager);
		}
		#endregion

		#region Add Tests
		[TestMethod]
		public void AddItemTest()
		{
			var xmlDoc = new XmlDocument(ExcelSlicer.SlicerDocumentNamespaceManager.NameTable);
			xmlDoc.LoadXml("<items count=\"1\" xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"><i x=\"2\" s=\"1\"/></items>");
			var node = xmlDoc.FirstChild;
			var slicerCacheNamespaceManager = ExcelSlicer.SlicerDocumentNamespaceManager;
			var tabularItems = new SlicerCacheTabularItems(node, slicerCacheNamespaceManager);
			Assert.AreEqual(1, tabularItems.Count);
			tabularItems.Add(1, true);
			Assert.AreEqual(2, tabularItems.Count);
			Assert.AreEqual(2, tabularItems[0].AtomIndex);
			Assert.IsTrue(tabularItems[0].IsSelected);
			Assert.AreEqual(1, tabularItems[1].AtomIndex);
			Assert.IsTrue(tabularItems[1].IsSelected);
		}

		[TestMethod]
		public void AddItemIsSelectedFalse()
		{
			var xmlDoc = new XmlDocument(ExcelSlicer.SlicerDocumentNamespaceManager.NameTable);
			xmlDoc.LoadXml("<items count=\"1\" xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"><i x=\"2\" s=\"1\"/></items>");
			var node = xmlDoc.FirstChild;
			var slicerCacheNamespaceManager = ExcelSlicer.SlicerDocumentNamespaceManager;
			var tabularItems = new SlicerCacheTabularItems(node, slicerCacheNamespaceManager);
			tabularItems.Add(1, false);
			Assert.AreEqual(2, tabularItems.Count);
			Assert.AreEqual(2, tabularItems[0].AtomIndex);
			Assert.IsTrue(tabularItems[0].IsSelected);
			Assert.AreEqual(1, tabularItems[1].AtomIndex);
			Assert.IsFalse(tabularItems[1].IsSelected);
		}
		#endregion

		#region Clear Tests
		[TestMethod]
		public void ClearItemsTest()
		{
			var xmlDoc = new XmlDocument(ExcelSlicer.SlicerDocumentNamespaceManager.NameTable);
			xmlDoc.LoadXml("<items count=\"1\" xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"><i x=\"2\" s=\"1\"/></items>");
			var node = xmlDoc.FirstChild;
			var slicerCacheNamespaceManager = ExcelSlicer.SlicerDocumentNamespaceManager;
			var tabularItems = new SlicerCacheTabularItems(node, slicerCacheNamespaceManager);
			tabularItems.Add(1, false);
			tabularItems.Clear();
		}
		#endregion
	}
}
