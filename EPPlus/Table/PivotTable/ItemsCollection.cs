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
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Table.PivotTable
{
	/// <summary>
	/// Collection of row or column items.
	/// </summary>
	public class ItemsCollection : XmlCollectionBase<RowColumnItem>
	{
		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="ItemsCollection"/>.
		/// </summary>
		/// <param name="namespaceManager">The namespace manager.</param>
		/// <param name="node">The top node.</param>
		public ItemsCollection(XmlNamespaceManager namespaceManager, XmlNode node) : base(namespaceManager, node)
		{
			if (node == null)
				throw new ArgumentNullException(nameof(node));
		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Adds a new <see cref="RowColumnItem"/> to this collection.
		/// </summary>
		/// <param name="rowDepth">The row depth of this item in the pivot table.</param>
		/// <param name="value">The value of it's 'x' attribute.</param>
		/// <param name="itemType">The value of the 't' attribute.</param>
		/// <param name="dataFieldIndex">The value of the 'i' attribute.</param>
		public void Add(int rowDepth, int value, string itemType = null, int dataFieldIndex = 0)
		{
			base.AddItem(new RowColumnItem(this.NameSpaceManager, base.TopNode, rowDepth, value, itemType, dataFieldIndex));
		}

		/// <summary>
		/// Adds a new <see cref="RowColumnItem"/> to this collection.
		/// </summary>
		/// <param name="xMembers">The list of member property indices.</param>
		/// <param name="repeatedItemsCount">The repeated items count value.</param>
		/// <param name="dataFieldIndex">The 'i' attribute value which points to a data field.</param>
		public void AddColumnItem(List<Tuple<int, int>> xMembers, int repeatedItemsCount, int dataFieldIndex)
		{
			// Shared member property indexes are implicit to the parent.
			xMembers.RemoveRange(0, repeatedItemsCount);
			base.AddItem(new RowColumnItem(this.NameSpaceManager, base.TopNode, xMembers, repeatedItemsCount, dataFieldIndex));
		}

		/// <summary>
		/// Adds a new <see cref="RowColumnItem"/> to this collection.
		/// </summary>
		/// <param name="itemType">The string value of the 't' attribute.</param>
		/// <param name="repeatedItemsCount">The 'r' attribute value.</param>
		/// <param name="xAttribute">The 'x' attribute value.</param>
		/// <param name="dataFieldIndex">The 'i' attribute value which points to a data field.</param>
		public void AddSumNode(string itemType, int repeatedItemsCount = 0, int xAttribute = 0, int dataFieldIndex = 0)
		{
			base.AddItem(new RowColumnItem(this.NameSpaceManager, base.TopNode, repeatedItemsCount, xAttribute, itemType, dataFieldIndex));
		}

		/// <summary>
		/// Clears all existing rowItems.
		/// </summary>
		public void Clear()
		{
			base.ClearItems();
		}
		#endregion

		#region XmlCollectionBase Overrides
		/// <summary>
		/// Loads the row/column items from the xml document.
		/// </summary>
		/// <returns>The collection of items.</returns>
		protected override List<RowColumnItem> LoadItems()
		{
			var items = new List<RowColumnItem>();
			var xNodes = base.TopNode.SelectNodes("d:i", base.NameSpaceManager);
			foreach (XmlNode xmlNode in xNodes)
			{
				items.Add(new RowColumnItem(base.NameSpaceManager, xmlNode));
			}
			return items;
		}
		#endregion
	}
}