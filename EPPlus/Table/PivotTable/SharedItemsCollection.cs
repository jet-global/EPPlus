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
	/// Collection class for <see cref="CacheItem"/>.
	/// </summary>
	public class SharedItemsCollection : XmlCollectionBase<CacheItem>
	{
		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="SharedItemsCollection"/>.
		/// </summary>
		/// <param name="namespaceManager">The namespace manager.</param>
		/// <param name="node">The xml top node.</param>
		public SharedItemsCollection(XmlNamespaceManager namespaceManager, XmlNode node) : base(namespaceManager, node)
		{
			if (node == null)
				throw new ArgumentNullException(nameof(node));
		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Adds a new field item to the list.
		/// </summary>
		/// <param name="value">The value.</param>
		/// <returns>The index of the new item.</returns>
		public void Add(object value)
		{
			string stringValue = ConvertUtil.ConvertObjectToXmlAttributeString(value);
			var item = new CacheItem(this.NameSpaceManager, base.TopNode, CacheItem.GetObjectType(value), stringValue);
			item.AddSelf(base.TopNode);
			base.AddItem(item);
		}
		#endregion

		#region XmlCollectionBase Overrides
		/// <summary>
		/// Loads the sharedItems from the xml document.
		/// </summary>
		/// <returns>The collection of sharedItems.</returns>
		protected override List<CacheItem> LoadItems()
		{
			var items = new List<CacheItem>();
			// Selects all possible child node types.
			foreach (XmlNode sharedItem in base.TopNode.SelectNodes("d:b | d:d | d:e | d:m | d:n | d:s | d:x", this.NameSpaceManager))
			{
				items.Add(new CacheItem(this.NameSpaceManager, sharedItem));
			}
			return items;
		}
		#endregion
	}
}