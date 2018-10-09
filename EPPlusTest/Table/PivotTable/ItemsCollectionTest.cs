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
	public class ItemsCollectionTest
	{
		#region Constructor Tests
		[TestMethod]
		public void ItemsCollection()
		{
			var xmlDoc = new XmlDocument(TestUtility.CreateDefaultNSM().NameTable);
			xmlDoc.LoadXml(
			$@"<rowItems xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" count=""3"">
				<i>
					<x v=""1""/>
				</i>
				<i r=""1"">
					<x v=""2""/>
				</i>
				<i r=""1"">
					<x v=""3""/>
				</i>
			</rowItems>");
			var node = xmlDoc.FirstChild;
			var itemsCollection = new ItemsCollection(TestUtility.CreateDefaultNSM(), node);
			Assert.AreEqual(3, itemsCollection.Count);
			Assert.AreEqual(1, itemsCollection[0].Count);
			Assert.AreEqual(1, itemsCollection[1].Count);
			Assert.AreEqual(1, itemsCollection[2].Count);
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void ItemsCollectionNullNamespaceManager()
		{
			var xmlDoc = new XmlDocument(TestUtility.CreateDefaultNSM().NameTable);
			xmlDoc.LoadXml(
			$@"<rowItems xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" count=""3"">
				<i>
					<x v=""1""/>
				</i>
				<i r=""1"">
					<x v=""2""/>
				</i>
				<i r=""1"">
					<x v=""3""/>
				</i>
			</rowItems>");
			var node = xmlDoc.FirstChild;
			new ItemsCollection(null, node);
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void ItemsCollectionNullNode()
		{
			new ItemsCollection(TestUtility.CreateDefaultNSM(), null);
		}
		#endregion

		#region Add Tests
		[TestMethod]
		public void Add()
		{
			var xmlDoc = new XmlDocument(TestUtility.CreateDefaultNSM().NameTable);
			xmlDoc.LoadXml(
			$@"<rowItems xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" count=""3"">
				<i>
					<x v=""1""/>
				</i>
				<i r=""1"">
					<x v=""2""/>
				</i>
				<i r=""1"">
					<x v=""3""/>
				</i>
			</rowItems>");
			var node = xmlDoc.FirstChild;
			var itemsCollection = new ItemsCollection(TestUtility.CreateDefaultNSM(), node);
			itemsCollection.Add(5, 3);
			Assert.AreEqual(4, itemsCollection.Count);
			Assert.AreEqual(4, node.ChildNodes.Count);
			Assert.AreEqual(5, itemsCollection[3].RepeatedItemsCount);
			Assert.AreEqual(1, itemsCollection[3].Count);
		}
		#endregion

		#region AddSumNode Tests
		[TestMethod]
		public void AddSumNode()
		{
			var xmlDoc = new XmlDocument(TestUtility.CreateDefaultNSM().NameTable);
			xmlDoc.LoadXml(
			$@"<rowItems xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" count=""3"">
				<i>
					<x v=""1""/>
				</i>
				<i r=""1"">
					<x v=""2""/>
				</i>
				<i r=""1"">
					<x v=""3""/>
				</i>
			</rowItems>");
			var node = xmlDoc.FirstChild;
			var itemsCollection = new ItemsCollection(TestUtility.CreateDefaultNSM(), node);
			itemsCollection.AddSumNode("grand");
			Assert.AreEqual(4, itemsCollection.Count);
			Assert.AreEqual(4, node.ChildNodes.Count);
			Assert.AreEqual("grand", itemsCollection[3].ItemType);
		}
		#endregion

		#region Clear Tests
		[TestMethod]
		public void ClearCollection()
		{
			var xmlDoc = new XmlDocument(TestUtility.CreateDefaultNSM().NameTable);
			xmlDoc.LoadXml(
			$@"<rowItems xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" count=""3"">
				<i>
					<x v=""1""/>
				</i>
				<i r=""1"">
					<x v=""2""/>
				</i>
				<i r=""1"">
					<x v=""3""/>
				</i>
			</rowItems>");
			var node = xmlDoc.FirstChild;
			var itemsCollection = new ItemsCollection(TestUtility.CreateDefaultNSM(), node);
			itemsCollection.Clear();
			Assert.AreEqual(0, itemsCollection.Count);
			Assert.AreEqual(0, node.ChildNodes.Count);
		}
		#endregion

		#region LoadItems Tests
		[TestMethod]
		public void LoadItems()
		{
			var namespaceManager = TestUtility.CreateDefaultNSM();
			var xmlDoc = new XmlDocument(namespaceManager.NameTable);
			xmlDoc.LoadXml(
			$@"<rowItems xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" count=""3"">
				<i>
					<x v=""1""/>
				</i>
				<i r=""1"">
					<x v=""2""/>
				</i>
				<i r=""1"">
					<x v=""3""/>
				</i>
			</rowItems>");
			var node = xmlDoc.FirstChild;
			var itemsCollection = new ItemsCollection(namespaceManager, node);
			var rowItems = new List<RowColumnItem>
			{
				new RowColumnItem(namespaceManager, node, 0, 1),
				new RowColumnItem(namespaceManager, node, 1, 2),
				new RowColumnItem(namespaceManager, node, 1, 3)
			};
			Assert.AreEqual(itemsCollection.Count, rowItems.Count);
			for (int i = 0; i < itemsCollection.Count; i++)
			{
				var actual = itemsCollection[i];
				var expected = rowItems[i];
				Assert.AreEqual(expected.RepeatedItemsCount, actual.RepeatedItemsCount);
				Assert.AreEqual(expected.Count, actual.Count);
				Assert.AreEqual(expected[0], actual[0]);
				Assert.AreEqual(expected.ItemType, actual.ItemType);
			}
		}
		#endregion
	}
}