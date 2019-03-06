/*******************************************************************************
* You may amend and distribute as you like, but don't remove this header!
*
* EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
* See http://www.codeplex.com/EPPlus for details.
*
* Copyright (C) 2011-2018 Michelle Lau and others as noted in the source history.
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
using System.Xml;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Table.PivotTable.Filters;

namespace EPPlusTest.Table.PivotTable.Filters
{
	[TestClass]
	public class ExcelCustomFiltersCollectionTest
	{
		#region Constructor Tests
		[TestMethod]
		public void ExcelCustomFiltersCollectionConstructorTest()
		{
			XmlDocument document = new XmlDocument();
			document.LoadXml(@"<customFilters xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
								<customFilter operator=""greaterThan"" val=""512*""/>
							</customFilters>");
			var namespaceManager = TestUtility.CreateDefaultNSM();
			var customFilterCollection = new ExcelCustomFiltersCollection(namespaceManager, document.SelectSingleNode("//d:customFilters", namespaceManager));
			Assert.IsNotNull(customFilterCollection);
			Assert.AreEqual("greaterThan", customFilterCollection[0].FilterComparisonOperator);
			Assert.AreEqual("512*", customFilterCollection[0].TopOrBottomValue);
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void ExcelCustomFiltersCollectionConstructorNullNamespaceManagerTest()
		{
			XmlDocument document = new XmlDocument();
			document.LoadXml(@"<customFilters xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
								<customFilter operator=""greaterThan"" val=""512*""/>
							</customFilters>");
			new ExcelCustomFiltersCollection(null, document.SelectSingleNode("//d:customFilters", TestUtility.CreateDefaultNSM()));
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void ExcelCustomFiltersCollectionConstructorNullNodeTest()
		{
			new ExcelCustomFiltersCollection(TestUtility.CreateDefaultNSM(), null);
		}
		#endregion

		#region LoadItems Tests
		[TestMethod]
		public void LoadItems()
		{
			var namespaceManager = TestUtility.CreateDefaultNSM();
			var xmlDoc = new XmlDocument(namespaceManager.NameTable);
			xmlDoc.LoadXml(@"<customFilters xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" and=""1"">
					<customFilter operator=""greaterThanOrEqual"" val=""c""/>
					<customFilter operator=""lessThanOrEqual"" val=""p""/>
			</customFilters>");
			var node = xmlDoc.FirstChild;
			var itemsCollection = new ExcelCustomFiltersCollection(namespaceManager, node);
			Assert.AreEqual("greaterThanOrEqual", itemsCollection[0].FilterComparisonOperator);
			Assert.AreEqual("c", itemsCollection[0].TopOrBottomValue);
			Assert.AreEqual("lessThanOrEqual", itemsCollection[1].FilterComparisonOperator);
			Assert.AreEqual("p", itemsCollection[1].TopOrBottomValue);
		}
		#endregion
	}
}
