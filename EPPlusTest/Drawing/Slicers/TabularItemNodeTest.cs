﻿using System;
using System.Xml;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Drawing.Slicers;

namespace EPPlusTest.Drawing.Slicers
{
	[TestClass]
	public class TabularItemNodeTest
	{
		#region Constructor Tests
		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void TabularItemNodeNullXmlNodeThrowsException()
		{
			var namespaceManager = TestUtility.CreateDefaultNSM();
			new TabularItemNode(null, namespaceManager);
		}
		#endregion

		#region AtomIndex Tests
		[TestMethod]
		public void TabularItemNodeAtomIndex()
		{
			var xmlDoc = new XmlDocument(ExcelSlicer.SlicerDocumentNamespaceManager.NameTable);
			xmlDoc.LoadXml("<i x=\"2\" s=\"1\"/>");
			var node = xmlDoc.FirstChild;
			var namespaceManager = TestUtility.CreateDefaultNSM();
			var tabularItemNode = new TabularItemNode(node, namespaceManager);
			Assert.AreEqual(2, tabularItemNode.AtomIndex);
			tabularItemNode.AtomIndex = 3;
			Assert.AreEqual(3, tabularItemNode.AtomIndex);
			Assert.AreEqual("<i x=\"3\" s=\"1\" />", node.OuterXml);
		}

		[TestMethod]
		public void TabularItemNodeAtomIndexDoesNotExist()
		{
			var xmlDoc = new XmlDocument(ExcelSlicer.SlicerDocumentNamespaceManager.NameTable);
			xmlDoc.LoadXml("<i s=\"1\"/>");
			var node = xmlDoc.FirstChild;
			var namespaceManager = TestUtility.CreateDefaultNSM();
			var tabularItemNode = new TabularItemNode(node, namespaceManager);
			Assert.AreEqual(-1, tabularItemNode.AtomIndex);
			tabularItemNode.AtomIndex = 3;
			Assert.AreEqual(3, tabularItemNode.AtomIndex);
			Assert.AreEqual("<i s=\"1\" x=\"3\" />", node.OuterXml);
		}
		#endregion

		#region IsSelected Tests
		[TestMethod]
		public void TabularItemNodeIsSelected()
		{
			var xmlDoc = new XmlDocument(ExcelSlicer.SlicerDocumentNamespaceManager.NameTable);
			xmlDoc.LoadXml("<i x=\"2\" s=\"1\"/>");
			var node = xmlDoc.FirstChild;
			var namespaceManager = TestUtility.CreateDefaultNSM();
			var tabularItemNode = new TabularItemNode(node, namespaceManager);
			Assert.AreEqual(true, tabularItemNode.IsSelected);
			tabularItemNode.IsSelected = false;
			Assert.AreEqual(false, tabularItemNode.IsSelected);
			Assert.AreEqual("<i x=\"2\" s=\"0\" />", node.OuterXml);
		}

		[TestMethod]
		public void TabularItemNodeIsSelectedDoesNotExist()
		{
			var xmlDoc = new XmlDocument(ExcelSlicer.SlicerDocumentNamespaceManager.NameTable);
			xmlDoc.LoadXml("<i x=\"1\"/>");
			var node = xmlDoc.FirstChild;
			var namespaceManager = TestUtility.CreateDefaultNSM();
			var tabularItemNode = new TabularItemNode(node, namespaceManager);
			Assert.AreEqual(false, tabularItemNode.IsSelected);
			tabularItemNode.IsSelected = true;
			Assert.AreEqual(true, tabularItemNode.IsSelected);
			Assert.AreEqual("<i x=\"1\" s=\"1\" />", node.OuterXml);
		}
		#endregion

		#region NonDisplay Tests
		[TestMethod]
		public void TabularItemNodeNonDisplay()
		{
			var xmlDoc = new XmlDocument(ExcelSlicer.SlicerDocumentNamespaceManager.NameTable);
			xmlDoc.LoadXml("<i x=\"2\" s=\"1\" nd=\"1\"/>");
			var node = xmlDoc.FirstChild;
			var namespaceManager = TestUtility.CreateDefaultNSM();
			var tabularItemNode = new TabularItemNode(node, namespaceManager);
			Assert.AreEqual(true, tabularItemNode.NoData);
			tabularItemNode.NoData = false;
			Assert.AreEqual(false, tabularItemNode.NoData);
			Assert.AreEqual("<i x=\"2\" s=\"1\" nd=\"0\" />", node.OuterXml);
		}

		[TestMethod]
		public void TabularItemNodeNonDisplayDoesNotExist()
		{
			var xmlDoc = new XmlDocument(ExcelSlicer.SlicerDocumentNamespaceManager.NameTable);
			xmlDoc.LoadXml("<i x=\"1\" s=\"1\"/>");
			var node = xmlDoc.FirstChild;
			var namespaceManager = TestUtility.CreateDefaultNSM();
			var tabularItemNode = new TabularItemNode(node, namespaceManager);
			Assert.AreEqual(false, tabularItemNode.NoData);
			tabularItemNode.NoData = true;
			Assert.AreEqual(true, tabularItemNode.NoData);
			Assert.AreEqual("<i x=\"1\" s=\"1\" nd=\"1\" />", node.OuterXml);
		}
		#endregion
	}
}
