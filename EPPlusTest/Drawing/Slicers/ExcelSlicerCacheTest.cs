﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Drawing.Slicers;

namespace EPPlusTest.Drawing.Slicers
{
	[TestClass]
	public class ExcelSlicerCacheTest
	{
		#region Nested Class Tests

		#region PivotTableNode Tests

		#region Constructor Tests
		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void PivotTableNodeNullXmlNodeThrowsException()
		{
			new ExcelSlicerCache.PivotTableNode(null);
		}
		#endregion

		#region TabId Tests
		[TestMethod]
		public void TabId()
		{
			var xmlDoc = new XmlDocument();
			xmlDoc.LoadXml("<pivotTable tabId=\"2\" name=\"PivotTable1\" />");
			var node = xmlDoc.FirstChild;
			var pivotTable = new ExcelSlicerCache.PivotTableNode(node);
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
			var xmlDoc = new XmlDocument();
			xmlDoc.LoadXml("<pivotTable tabId=\"2\" name=\"PivotTable1\" />");
			var node = xmlDoc.FirstChild;
			var pivotTable = new ExcelSlicerCache.PivotTableNode(node);
			Assert.AreEqual("PivotTable1", pivotTable.PivotTableName);
			pivotTable.PivotTableName = "PivotTable2";
			Assert.AreEqual("PivotTable2", pivotTable.PivotTableName);
			Assert.AreEqual("<pivotTable tabId=\"2\" name=\"PivotTable2\" />", node.OuterXml);
		}
		#endregion

		#endregion

		#endregion
	}
}
