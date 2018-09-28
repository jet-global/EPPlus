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
			Assert.AreEqual(3, itemsCollection.Items.Count);
			Assert.AreEqual(itemsCollection.Count, itemsCollection.Items.Count);
		}
		#endregion
	}
}