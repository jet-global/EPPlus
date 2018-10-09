/*******************************************************************************
* You may amend and distribute as you like, but don't remove this header!
*
* EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
* See http://www.codeplex.com/EPPlus for details.
*
* Copyright (C) 2011-2018 Michelle Lau, Evan Schallerer, and others as noted in the source history.
*
* This library is free software; you can redistribute it and/or
* modify it under the terms of the GNU Lesser General Public
* License as published by the Free Software Foundation; either
* version 2.1 of the License, or (at your option) any later version.
* This library is distributed in the hope that it will be useful,
* but WITHOUT ANY WARRANTY; without even the implied warranty of
* MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
* See the GNU Lesser General Public License for more details.
*
* The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
* If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
*
* All code and executables are provided "as is" with no warranty either express or implied. 
* The author accepts no liability for any damage or loss of business that this product may cause.
*
* For code change notes, see the source control history.
*******************************************************************************/
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Table.PivotTable;

namespace EPPlusTest.Table.PivotTable
{
	[TestClass]
	public class CacheRecordNodeTest : PivotTableTestBase
	{
		#region Constructor Tests
		[TestMethod]
		public void CacheRecordNodeConstructorTest()
		{
			XmlDocument document = new XmlDocument();
			document.LoadXml(@"<pivotCacheRecords xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" count=""1""><r><n v=""20100076""/><x v=""0""/> <b v=""0""/> <m v=""0""/> <e v=""415.75""/><d v=""1""/></r></pivotCacheRecords>");
			var ns = TestUtility.CreateDefaultNSM();
			var node = new CacheRecordNode(ns, document.SelectSingleNode("//d:r", ns));
			Assert.AreEqual(6, node.Items.Count);
			Assert.AreEqual(1, node.Items.Count(i => i.Type == PivotCacheRecordType.b));
			Assert.AreEqual(1, node.Items.Count(i => i.Type == PivotCacheRecordType.x));
			Assert.AreEqual(1, node.Items.Count(i => i.Type == PivotCacheRecordType.d));
			Assert.AreEqual(1, node.Items.Count(i => i.Type == PivotCacheRecordType.e));
			Assert.AreEqual(1, node.Items.Count(i => i.Type == PivotCacheRecordType.m));
			Assert.AreEqual(1, node.Items.Count(i => i.Type == PivotCacheRecordType.n));
			Assert.AreEqual(0, node.Items.Count(i => i.Type == PivotCacheRecordType.s));
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableDataSourceTypeWorksheet.xlsx")]
		public void CacheRecordNodeConstructFromRowData()
		{
			var file = new FileInfo("PivotTableDataSourceTypeWorksheet.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var cacheDefinition = package.Workbook.PivotCacheDefinitions.First();
				XmlDocument document = new XmlDocument();
				document.LoadXml(@"<pivotCacheRecords xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" count=""0""><r></r></pivotCacheRecords>");
				var ns = TestUtility.CreateDefaultNSM();
				var node = document.SelectSingleNode("//d:r", ns);
				var row = new List<object> { 5, "e", false, 100 };
				var cacheRecordNode = new CacheRecordNode(ns, node, row, cacheDefinition);
				Assert.AreEqual(4, cacheRecordNode.Items.Count);
				Assert.AreEqual("5", cacheRecordNode.Items[0].Value);
				Assert.AreEqual("3", cacheRecordNode.Items[1].Value);
				Assert.AreEqual("2", cacheRecordNode.Items[2].Value);
				Assert.AreEqual("100", cacheRecordNode.Items[3].Value);
				var cacheField1 = cacheDefinition.CacheFields[0];
				var cacheField2 = cacheDefinition.CacheFields[1];
				var cacheField3 = cacheDefinition.CacheFields[2];
				var cacheField4 = cacheDefinition.CacheFields[3];
				Assert.AreEqual(0, cacheField1.SharedItems.Count);
				Assert.IsNotNull(cacheField2.SharedItems.SingleOrDefault(i => i.Value == "e" && i.Type == PivotCacheRecordType.s));
				Assert.IsNotNull(cacheField3.SharedItems.SingleOrDefault(i => i.Value == "0" && i.Type == PivotCacheRecordType.b));
				Assert.AreEqual(0, cacheField4.SharedItems.Count);
			}
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void CacheFieldNodeTestNullNode()
		{
			new CacheFieldNode(TestUtility.CreateDefaultNSM(), null);
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void CacheFieldNodeTestNullNamespaceManager()
		{
			var xml = new XmlDocument();
			xml.LoadXml(@"<pivotCacheRecords count=""1""><r><n v=""20100076""/></r></pivotCacheRecords>");
			new CacheFieldNode(null, xml.FirstChild);
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableDataSourceTypeWorksheet.xlsx")]
		[ExpectedException(typeof(ArgumentNullException))]
		public void ConstructCacheRecordNodeWithNullRow()
		{
			var file = new FileInfo("PivotTableDataSourceTypeWorksheet.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				XmlDocument document = new XmlDocument();
				document.LoadXml(@"<pivotCacheRecords xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" count=""0""><r></r></pivotCacheRecords>");
				var ns = TestUtility.CreateDefaultNSM();
				var node = document.SelectSingleNode("//d:r", ns);
				var cacheDefinition = package.Workbook.PivotCacheDefinitions.First();
				var record = cacheDefinition.CacheRecords[0];
				new CacheRecordNode(ns, node, null, cacheDefinition);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableDataSourceTypeWorksheet.xlsx")]
		[ExpectedException(typeof(ArgumentNullException))]
		public void ConstructCacheRecordNodeWithNullCacheDefinition()
		{
			var file = new FileInfo("PivotTableDataSourceTypeWorksheet.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				XmlDocument document = new XmlDocument();
				document.LoadXml(@"<pivotCacheRecords xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" count=""0""><r></r></pivotCacheRecords>");
				var ns = TestUtility.CreateDefaultNSM();
				var node = document.SelectSingleNode("//d:r", ns);
				var cacheDefinition = package.Workbook.PivotCacheDefinitions.First();
				var record = cacheDefinition.CacheRecords[0];
				var row = new List<object> { 5, "e", false, 100 };
				new CacheRecordNode(ns, node, row, null);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableDataSourceTypeWorksheet.xlsx")]
		[ExpectedException(typeof(InvalidOperationException))]
		public void ConstructCacheRecordNodeWithIncorrectNumberOfFields()
		{
			var file = new FileInfo("PivotTableDataSourceTypeWorksheet.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				XmlDocument document = new XmlDocument();
				document.LoadXml(@"<pivotCacheRecords xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" count=""0""><r></r></pivotCacheRecords>");
				var ns = TestUtility.CreateDefaultNSM();
				var node = document.SelectSingleNode("//d:r", ns);
				var cacheDefinition = package.Workbook.PivotCacheDefinitions.First();
				var row = new List<object> { 5, "e", false };
				var record = cacheDefinition.CacheRecords[0];
				new CacheRecordNode(ns, node, row, cacheDefinition);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableDataSourceTypeWorksheet.xlsx")]
		[ExpectedException(typeof(ArgumentNullException))]
		public void ConstructCacheRecordNodeWithNullParentNode()
		{
			var file = new FileInfo("PivotTableDataSourceTypeWorksheet.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var ns = TestUtility.CreateDefaultNSM();
				var cacheDefinition = package.Workbook.PivotCacheDefinitions.First();
				var row = new List<object> { 5, "e", false, 100 };
				var record = cacheDefinition.CacheRecords[0];
				new CacheRecordNode(ns, null, row, cacheDefinition);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableDataSourceTypeWorksheet.xlsx")]
		[ExpectedException(typeof(ArgumentNullException))]
		public void ConstructCacheRecordNodeWithNullNamespaceManager()
		{
			var file = new FileInfo("PivotTableDataSourceTypeWorksheet.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				XmlDocument document = new XmlDocument();
				document.LoadXml(@"<pivotCacheRecords xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" count=""0""><r></r></pivotCacheRecords>");
				var node = document.SelectSingleNode("//d:r", TestUtility.CreateDefaultNSM());
				var cacheDefinition = package.Workbook.PivotCacheDefinitions.First();
				var row = new List<object> { 5, "e", false, 100 };
				var record = cacheDefinition.CacheRecords[0];
				new CacheRecordNode(null, node, row, cacheDefinition);
			}
		}
		#endregion

		#region Update Tests
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableDataSourceTypeWorksheet.xlsx")]
		public void UpdateRecords()
		{
			var file = new FileInfo("PivotTableDataSourceTypeWorksheet.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var cacheDefinition = package.Workbook.PivotCacheDefinitions.First();
				var row = new List<object> { 5, "e", false, 100 };
				var record = cacheDefinition.CacheRecords[0];
				record.Update(row, cacheDefinition);
				Assert.AreEqual(4, record.Items.Count);
				Assert.AreEqual("5", record.Items[0].Value);
				Assert.AreEqual("3", record.Items[1].Value);
				Assert.AreEqual("2", record.Items[2].Value);
				Assert.AreEqual("100", record.Items[3].Value);
				var cacheField1 = cacheDefinition.CacheFields[0];
				var cacheField2 = cacheDefinition.CacheFields[1];
				var cacheField3 = cacheDefinition.CacheFields[2];
				var cacheField4 = cacheDefinition.CacheFields[3];
				Assert.AreEqual(0, cacheField1.SharedItems.Count);
				Assert.IsNotNull(cacheField2.SharedItems.SingleOrDefault(i => i.Value == "e" && i.Type == PivotCacheRecordType.s));
				Assert.IsNotNull(cacheField3.SharedItems.SingleOrDefault(i => i.Value == "0" && i.Type == PivotCacheRecordType.b));
				Assert.AreEqual(0, cacheField4.SharedItems.Count);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableDataSourceTypeWorksheet.xlsx")]
		[ExpectedException(typeof(ArgumentNullException))]
		public void UpdateWithNullRow()
		{
			var file = new FileInfo("PivotTableDataSourceTypeWorksheet.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var cacheDefinition = package.Workbook.PivotCacheDefinitions.First();
				var record = cacheDefinition.CacheRecords[0];
				record.Update(null, cacheDefinition);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableDataSourceTypeWorksheet.xlsx")]
		[ExpectedException(typeof(ArgumentNullException))]
		public void UpdateWithNullCacheDefinition()
		{
			var file = new FileInfo("PivotTableDataSourceTypeWorksheet.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var cacheDefinition = package.Workbook.PivotCacheDefinitions.First();
				var record = cacheDefinition.CacheRecords[0];
				var row = new List<object> { 5, "e", false, 100 };
				record.Update(row, null);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PivotTableDataSourceTypeWorksheet.xlsx")]
		[ExpectedException(typeof(InvalidOperationException))]
		public void UpdateWithIncorrectNumberOfFields()
		{
			var file = new FileInfo("PivotTableDataSourceTypeWorksheet.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var cacheDefinition = package.Workbook.PivotCacheDefinitions.First();
				var row = new List<object> { 5, "e", false };
				var record = cacheDefinition.CacheRecords[0];
				record.Update(row, cacheDefinition);
			}
		}
		#endregion
	}
}
