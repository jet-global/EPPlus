/*******************************************************************************
* You may amend and distribute as you like, but don't remove this header!
*
* EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
* See http://www.codeplex.com/EPPlus for details.
*
* Copyright (C) 2011-2018 Jan Källman, Evan Schallerer, and others as noted in the source history.
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
using OfficeOpenXml.Extensions;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Table.PivotTable
{
	/// <summary>
	/// Wraps a <cacheField/> node in <pivotcachedefinition-cacheFields/>.
	/// </summary>
	public class CacheFieldNode
	{
		#region Properties
		/// <summary>
		/// Gets or sets the name for this <see cref="CacheFieldNode"/>.
		/// </summary>
		public string Name
		{
			get { return this.Node.Attributes["name"].Value; }
			set { this.Node.Attributes["name"].Value = value; }
		}

		/// <summary>
		/// Gets or sets the number format ID for this <see cref="CacheFieldNode"/>.
		/// </summary>
		public string NumFormatId
		{
			get { return this.Node.Attributes["numFmtId"].Value; }
			set { this.Node.Attributes["numFmtId"].Value = value; }
		}

		/// <summary>
		/// Gets the sharedItems for this node.
		/// </summary>
		public SharedItemsCollection SharedItems { get; }

		/// <summary>
		/// Gets a value indicating whether or not this node has shared items.
		/// </summary>
		public bool HasSharedItems
		{
			get { return this.SharedItems.Count > 0; }
		}

		private XmlNode Node { get; set; }
		private XmlNamespaceManager NameSpaceManager { get; set; }
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="CacheFieldNode"/>.
		/// </summary>
		/// <param name="node">The <see cref="XmlNode"/> for this <see cref="CacheFieldNode"/>.</param>
		/// <param name="namespaceManager">The namespace manager to use for searching child nodes.</param>
		public CacheFieldNode(XmlNamespaceManager namespaceManager, XmlNode node)
		{
			if (node == null)
				throw new ArgumentNullException(nameof(node));
			if (namespaceManager == null)
				throw new ArgumentNullException(nameof(namespaceManager));
			this.Node = node;
			this.NameSpaceManager = namespaceManager;
			var sharedItemsNode = node.SelectSingleNode("d:sharedItems", this.NameSpaceManager);
			if (sharedItemsNode != null)
				this.SharedItems = new SharedItemsCollection(this.NameSpaceManager, sharedItemsNode);
		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Gets the index of the target value.
		/// </summary>
		/// <param name="type">The type of the target value.</param>
		/// <param name="value">The target value in the list.</param>
		/// <returns>The index of the value in the list.</returns>
		public int GetSharedItemIndex(PivotCacheRecordType type, object value)
		{
			string stringValue = ConvertUtil.ConvertObjectToXmlAttributeString(value);
			for (int i = 0; i < this.SharedItems.Count; i++)
			{
				var item = this.SharedItems[i];
				if (type == item.Type && stringValue.IsEquivalentTo(item.Value))
					return i;
			}
			return -1;
		}
		#endregion
	}
}