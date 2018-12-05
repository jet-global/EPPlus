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
using System.Xml;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Table.PivotTable;

namespace EPPlusTest.Table.PivotTable
{
	[TestClass]
	public class SharedItemsCollectionTest : PivotTableTestBase
	{
		#region Constructor Tests
		[TestMethod]
		public void SharedItemsCollectionConstructorTest()
		{
			XmlDocument document = new XmlDocument();
			document.LoadXml(@"<sharedItems xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" count=""6""><n v=""20100076""/><x v=""0""/> <b v=""0""/> <m v=""0""/> <e v=""415.75""/><d v=""1""/></sharedItems>");
			var namespaceManager = TestUtility.CreateDefaultNSM();
			var sharedItems = new SharedItemsCollection(namespaceManager, document.SelectSingleNode("//d:sharedItems", namespaceManager));
			Assert.AreEqual(6, sharedItems.Count);
			Assert.AreEqual("20100076", sharedItems[0].Value);
			Assert.AreEqual(PivotCacheRecordType.n, sharedItems[0].Type);
			Assert.AreEqual("0", sharedItems[1].Value);
			Assert.AreEqual(PivotCacheRecordType.x, sharedItems[1].Type);
			Assert.AreEqual("0", sharedItems[2].Value);
			Assert.AreEqual(PivotCacheRecordType.b, sharedItems[2].Type);
			Assert.IsNull(sharedItems[3].Value);
			Assert.AreEqual(PivotCacheRecordType.m, sharedItems[3].Type);
			Assert.AreEqual("415.75", sharedItems[4].Value);
			Assert.AreEqual(PivotCacheRecordType.e, sharedItems[4].Type);
			Assert.AreEqual("1", sharedItems[5].Value);
			Assert.AreEqual(PivotCacheRecordType.d, sharedItems[5].Type);
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void SharedItemsCollectionNullNamespaceManager()
		{
			XmlDocument document = new XmlDocument();
			document.LoadXml(@"<sharedItems xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" count=""1""><n v=""20100076""/><x v=""0""/> <b v=""0""/> <m v=""0""/> <e v=""415.75""/><d v=""1""/></sharedItems>");
			new SharedItemsCollection(null, document.SelectSingleNode("//d:sharedItems", TestUtility.CreateDefaultNSM()));
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void SharedItemsCollectionNullNode()
		{
			new SharedItemsCollection(TestUtility.CreateDefaultNSM(), null);
		}
		#endregion

		#region Add Tests
		[TestMethod]
		public void AddItemTest()
		{
			var node = base.GetTestCacheFieldNode();
			node.SharedItems.Add("jet");
			Assert.AreEqual(4, node.SharedItems.Count);
			Assert.AreEqual("jet", node.SharedItems[2].Value);
			Assert.AreEqual(PivotCacheRecordType.s, node.SharedItems[2].Type);
		}

		[TestMethod]
		public void AddItemNullValueTest()
		{
			var node = base.GetTestCacheFieldNode();
			node.SharedItems.Add(null);
			Assert.AreEqual(4, node.SharedItems.Count);
			Assert.IsNull(node.SharedItems[2].Value);
			Assert.AreEqual(PivotCacheRecordType.m, node.SharedItems[2].Type);
		}
		#endregion

		#region LoadItems Tests
		[TestMethod]
		public void LoadItems()
		{
			var namespaceManager = TestUtility.CreateDefaultNSM();
			var xmlDoc = new XmlDocument(namespaceManager.NameTable);
			xmlDoc.LoadXml(
			@"<sharedItems xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" count=""4"">
				<n v=""20100076""/>
				<x v=""0""/>
				<b v=""0""/>
				<m/>
			</sharedItems>");
			var node = xmlDoc.FirstChild;
			var itemsCollection = new SharedItemsCollection(namespaceManager, node);
			Assert.AreEqual(4, itemsCollection.Count);
			var sharedItems = new List<CacheItem>
			{
				new CacheItem(namespaceManager, node, PivotCacheRecordType.n, "20100076"),
				new CacheItem(namespaceManager, node, PivotCacheRecordType.x, "0"),
				new CacheItem(namespaceManager, node, PivotCacheRecordType.b, "0"),
				new CacheItem(namespaceManager, node, PivotCacheRecordType.m, ""),
			};
			Assert.AreEqual(sharedItems.Count, itemsCollection.Count);
			for (int i = 0; i < itemsCollection.Count; i++)
			{
				var actual = itemsCollection[i];
				var expected = sharedItems[i];
				Assert.AreEqual(expected.Type, actual.Type);
				Assert.AreEqual(expected.Value, actual.Value);
			}
		}
		#endregion
	}
}