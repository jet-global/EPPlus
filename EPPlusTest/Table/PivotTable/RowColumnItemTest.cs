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
	public class RowColumnItemTest
	{
		#region Constructor Tests
		[TestMethod]
		public void RowColumnItem()
		{
			var xmlDoc = new XmlDocument(TestUtility.CreateDefaultNSM().NameTable);
			xmlDoc.LoadXml(
			$@"<i xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" r=""1"" i=""1"" t=""grand"">
					<x v=""1""/>
					<x v=""1048832""/>
					<x/>
				</i>");
			var node = xmlDoc.FirstChild;
			var itemsCollection = new RowColumnItem(TestUtility.CreateDefaultNSM(), node);
			Assert.AreEqual(3, itemsCollection.MemberPropertyIndex.Count);
			Assert.AreEqual(1, itemsCollection.RepeatedItemsCount);
			Assert.AreEqual(1, itemsCollection.DataFieldIndex);
			Assert.AreEqual("grand", itemsCollection.ItemType);
		}
		#endregion
	}
}
