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
	public class DiscreteGroupingPropertiesCollectionTest
	{
		#region Constructor Tests
		[TestMethod]
		public void DiscreteGroupingPropertiesCollectionConstructorTest()
		{
			XmlDocument document = new XmlDocument();
			document.LoadXml(@"<discretePr xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" count=""4"">
					<x v=""0""/>
					<x v=""1""/>
					<x v=""0""/>
					<x v=""1""/>
			</discretePr>");
			var namespaceManager = TestUtility.CreateDefaultNSM();
			var discreteGroupingProperties = new DiscreteGroupingPropertiesCollection(namespaceManager, document.SelectSingleNode("//d:discretePr", namespaceManager));
			Assert.AreEqual(4, discreteGroupingProperties.Count);
			Assert.AreEqual("0", discreteGroupingProperties[0].Value);
			Assert.AreEqual("1", discreteGroupingProperties[1].Value);
			Assert.AreEqual("0", discreteGroupingProperties[2].Value);
			Assert.AreEqual("1", discreteGroupingProperties[3].Value);
			foreach (var item in discreteGroupingProperties)
			{
				Assert.AreEqual(PivotCacheRecordType.x, item.Type);
			}
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void DiscreteGroupingPropertiesCollectionConstructorNullNamespaceManagerTest()
		{
			XmlDocument document = new XmlDocument();
			document.LoadXml(@"<discretePr xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" count=""4"">
					<x v=""0""/>
					<x v=""1""/>
					<x v=""0""/>
					<x v=""1""/>
			</discretePr>");
			var discreteGroupingProperties = new DiscreteGroupingPropertiesCollection(null, document.SelectSingleNode("//d:discretePr", TestUtility.CreateDefaultNSM()));
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void DiscreteGroupingPropertiesCollectionConstructorNullNodeTest()
		{
			new DiscreteGroupingPropertiesCollection(TestUtility.CreateDefaultNSM(), null);
		}
		#endregion

		#region LoadItems Test
		[TestMethod]
		public void LoadItems()
		{
			var namespaceManager = TestUtility.CreateDefaultNSM();
			var xmlDoc = new XmlDocument(namespaceManager.NameTable);
			xmlDoc.LoadXml(@"<discretePr xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" count=""4"">
					<x v=""0""/>
					<x v=""1""/>
					<x v=""0""/>
					<x v=""1""/>
			</discretePr>");
			var node = xmlDoc.FirstChild;
			var itemsCollection = new DiscreteGroupingPropertiesCollection(namespaceManager, node);
			Assert.AreEqual(4, itemsCollection.Count);
			var discreteGroupingProperties = new List<CacheItem>
			{
				new CacheItem(namespaceManager, node, PivotCacheRecordType.x, "0"),
				new CacheItem(namespaceManager, node, PivotCacheRecordType.x, "1"),
				new CacheItem(namespaceManager, node, PivotCacheRecordType.x, "0"),
				new CacheItem(namespaceManager, node, PivotCacheRecordType.x, "1"),
			};
			Assert.AreEqual(discreteGroupingProperties.Count, itemsCollection.Count);
			for(int i = 0; i < itemsCollection.Count; i++)
			{
				var actual = itemsCollection[i];
				var expected = discreteGroupingProperties[i];
				Assert.AreEqual(expected.Type, actual.Type);
				Assert.AreEqual(expected.Value, actual.Value);
			}
		}
		#endregion
	}
}
