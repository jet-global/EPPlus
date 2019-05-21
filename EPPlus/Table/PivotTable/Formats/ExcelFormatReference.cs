﻿/*******************************************************************************
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
using System.Linq;
using System.Xml;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Table.PivotTable.Formats
{
	/// <summary>
	/// A format reference that indicates the field and row/column item the formatting applies to.
	/// </summary>
	public class ExcelFormatReference : XmlCollectionItemBase
	{
		#region Properties
		/// <summary>
		/// Gets the index of the field the format filter applies to.
		/// </summary>
		public int FieldIndex
		{
			get { return base.GetXmlNodeInt("@field"); }
		}

		/// <summary>
		/// Gets the number of item indexes in the collection of indexes (x tags).
		/// </summary>
		public int ItemIndexCount
		{
			get { return base.GetXmlNodeInt("@count"); }
		}

		/// <summary>
		/// Gets a flag indicating if this reference is selected.
		/// </summary>
		public bool Selected
		{
			get { return base.GetXmlNodeBool("@selected", true); }
		}

		/// <summary>
		/// Gets the sharedItems for this node.
		/// </summary>
		public SharedItemsCollection SharedItems { get; }
		#endregion

		#region Constructors
		/// <summary>
		/// Creates a new instance of <see cref="ExcelFormatReference"/>.
		/// </summary>
		/// <param name="namespaceManager">The namespace manager.</param>
		/// <param name="node">The xml top node.</param>
		public ExcelFormatReference(XmlNamespaceManager namespaceManager, XmlNode node) : base(namespaceManager, node)
		{
			if (node == null)
				throw new ArgumentNullException(nameof(node));
			this.SharedItems = new SharedItemsCollection(this.NameSpaceManager, node);
		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Updates the values of each shared item in the reference's shared items collection.
		/// </summary>
		/// <param name="lookupNode">Function to get the node with the shared string value from the tree.</param>
		/// <param name="currentIndex">The index passed into getting the tree node.</param>
		public void UpdateSharedItemValues(Func<int, PivotItemTreeNode> lookupNode, ref int currentIndex)
		{
			if (this.FieldIndex < 0)
				return;
			foreach (var item in this.SharedItems)
			{
				var node = lookupNode(currentIndex);
				if (node != null)
					item.Value = node.PivotFieldItemIndex.ToString();
				if (this.SharedItems.Count > 1 && item != this.SharedItems.Last())
					currentIndex++;
			}
		}
		#endregion
	}
}
