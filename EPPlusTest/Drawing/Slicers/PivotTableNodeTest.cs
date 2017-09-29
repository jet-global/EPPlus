using System;
using System.Xml;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Drawing.Slicers;

namespace EPPlusTest.Drawing.Slicers
{
	[TestClass]
	public class PivotTableNodeTest
	{
		#region Constructor Tests
		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void PivotTableNodeNullXmlNodeThrowsException()
		{
			new PivotTableNode(null);
		}
		#endregion

		#region TabId Tests
		[TestMethod]
		public void TabId()
		{
			var xmlDoc = new XmlDocument(ExcelSlicer.SlicerDocumentNamespaceManager.NameTable);
			xmlDoc.LoadXml("<pivotTable tabId=\"2\" name=\"PivotTable1\" />");
			var node = xmlDoc.FirstChild;
			var pivotTable = new PivotTableNode(node);
			Assert.AreEqual("2", pivotTable.TabId);
			pivotTable.TabId = "4";
			Assert.AreEqual("4", pivotTable.TabId);
			Assert.AreEqual("<pivotTable tabId=\"4\" name=\"PivotTable1\" />", node.OuterXml);
		}
		#endregion

		#region PivotTableName Tests
		[TestMethod]
		public void PivotTableName()
		{
			var xmlDoc = new XmlDocument(ExcelSlicer.SlicerDocumentNamespaceManager.NameTable);
			xmlDoc.LoadXml("<pivotTable tabId=\"2\" name=\"PivotTable1\" />");
			var node = xmlDoc.FirstChild;
			var pivotTable = new PivotTableNode(node);
			Assert.AreEqual("PivotTable1", pivotTable.PivotTableName);
			pivotTable.PivotTableName = "PivotTable2";
			Assert.AreEqual("PivotTable2", pivotTable.PivotTableName);
			Assert.AreEqual("<pivotTable tabId=\"2\" name=\"PivotTable2\" />", node.OuterXml);
		}
		#endregion
	}
}
