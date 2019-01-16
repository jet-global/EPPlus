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
using System.Collections.Generic;
using System.Xml;
using OfficeOpenXml.Extensions;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Table.PivotTable
{
	/// <summary>
	/// Wraps a <cacheField/> node in <pivotcachedefinition-cacheFields/>.
	/// </summary>
	public class CacheFieldNode : XmlHelper
	{
		#region Properties
		/// <summary>
		/// Gets or sets the name for this <see cref="CacheFieldNode"/>.
		/// </summary>
		public string Name
		{
			get { return base.GetXmlNodeString("@name"); }
			set { base.SetXmlNodeString("@name", value); }
		}

		/// <summary>
		/// Gets or sets the number format ID for this <see cref="CacheFieldNode"/>.
		/// </summary>
		public int NumFormatId
		{
			get { return base.GetXmlNodeInt("@numFmtId"); }
			set { base.SetXmlNodeString("@numFmtId", value.ToString()); }
		}

		/// <summary>
		/// Gets the formula that defines the <see cref="CacheFieldNode"/>.
		/// </summary>
		public string Formula
		{
			get { return base.GetXmlNodeString("@formula"); }
			set { base.SetXmlNodeString("@formula", value); }
		}

		/// <summary>
		/// Gets or sets the formula for this cache field that has had any 
		/// references to other calculated cache fields resolved.
		/// </summary>
		public string ResolvedFormula { get; set; }

		/// <summary>
		/// Gets or sets a dictionary of a cache field name that was referenced in the cacheField formula 
		/// to the referenced cache field's index in the cache definition.
		/// </summary>
		public Dictionary<string, int> ReferencedCacheFieldsToIndex { get; set; } = new Dictionary<string, int>();

		/// <summary>
		/// Gets the sharedItems for this node.
		/// </summary>
		public SharedItemsCollection SharedItems { get; }

		/// <summary>
		/// Gets a value indicating whether or not this node has shared items.
		/// </summary>
		public bool HasSharedItems
		{
			get { return this.SharedItems?.Count > 0; }
		}
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="CacheFieldNode"/>.
		/// </summary>
		/// <param name="node">The <see cref="XmlNode"/> for this <see cref="CacheFieldNode"/>.</param>
		/// <param name="namespaceManager">The namespace manager to use for searching child nodes.</param>
		public CacheFieldNode(XmlNamespaceManager namespaceManager, XmlNode node) : base(namespaceManager, node)
		{
			if (node == null)
				throw new ArgumentNullException(nameof(node));
			if (namespaceManager == null)
				throw new ArgumentNullException(nameof(namespaceManager));
			var sharedItemsNode = base.TopNode.SelectSingleNode("d:sharedItems", this.NameSpaceManager);
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
				// Empty strings are sometimes put in as string values by Excel
				// so we will let empty types match with empty string shared items.
				if ((type == PivotCacheRecordType.m || type == item.Type) && stringValue.IsEquivalentTo(item.Value))
					return i;
			}
			return -1;
		}

		/// <summary>
		/// Remove the 'u' (unused) xml attribute from each <see cref="CacheItem"/> in the <see cref="SharedItemsCollection"/>.
		/// </summary>
		public void RemoveXmlUAttribute()
		{
			foreach (var item in this.SharedItems)
			{
				var unusedAttribute = item.TopNode.Attributes["u"];
				if (unusedAttribute != null && int.Parse(unusedAttribute.Value) == 1)
					item.TopNode.Attributes.Remove(unusedAttribute);
			}
		}
		#endregion
	}
}