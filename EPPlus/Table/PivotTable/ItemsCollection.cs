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

namespace OfficeOpenXml.Table.PivotTable
{
	/// <summary>
	/// Collection of row or column items.
	/// </summary>
	public class ItemsCollection : XmlHelper
	{
		#region Class Variables
		private List<RowColumnItem> myItems;
		#endregion

		#region Properties
		/// <summary>
		/// Gets the row/column items.
		/// </summary>
		public IReadOnlyList<RowColumnItem> Items
		{
			get
			{
				if (myItems == null)
				{
					myItems = new List<RowColumnItem>();
					var xNodes = base.TopNode.SelectNodes("d:i", base.NameSpaceManager);
					foreach (XmlNode xmlNode in xNodes)
					{
						myItems.Add(new RowColumnItem(base.NameSpaceManager, base.TopNode));
					}
				}
				return myItems;
			}
		}

		/// <summary>
		/// Gets or sets the count.
		/// </summary>
		public int Count
		{
			get { return base.GetXmlNodeIntNull("@count") ?? 0; }
			private set { base.SetXmlNodeString("@count", value.ToString()); }
		}
		#endregion

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
			if (namespaceManager == null)
				throw new ArgumentNullException(nameof(namespaceManager));
		}
		#endregion
	}
}