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
using System.Xml;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Table.PivotTable
{
	/// <summary>
	/// Wraps a node in <pivotCacheRecords-x/>.
	/// </summary>
	public class CacheItem : XmlCollectionItemBase
	{
		#region Properties
		/// <summary>
		/// Gets or sets the type of this item.
		/// </summary>
		public PivotCacheRecordType Type { get; private set; }

		/// <summary>
		/// Gets or sets the value of this item.
		/// </summary>
		public string Value
		{
			get
			{
				if (this.Type == PivotCacheRecordType.m)
					return null;
				return base.GetXmlNodeString("@v");
			}
			set
			{
				if (this.Type == PivotCacheRecordType.m)
					value = null;
				base.SetXmlNodeString("@v", value, true);
			}
		}
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="CacheItem"/> object.
		/// </summary>
		/// <param name="node">The <see cref="XmlNode"/> for this <see cref="CacheItem"/>.</param>
		/// <param name="namespaceManager">The namespace manger.</param>
		public CacheItem(XmlNamespaceManager namespaceManager, XmlNode node) : base(namespaceManager, node)
		{
			if (node == null)
				throw new ArgumentNullException(nameof(node));
			this.Type = (PivotCacheRecordType)Enum.Parse(typeof(PivotCacheRecordType), node.Name);
		}

		/// <summary>
		/// Creates an instance of a <see cref="CacheItem"/> object given a type and value.
		/// </summary>
		/// <param name="namespaceManager">The namespace manager.</param>
		/// <param name="parentNode">The parent xml node. It must be a sharedItems <see cref="XmlNode"/> or a cacheRecord <see cref="XmlNode"/>.</param>
		/// <param name="type">The type of this item.</param>
		/// <param name="value">The value of this item.</param>
		public CacheItem(XmlNamespaceManager namespaceManager, XmlNode parentNode, PivotCacheRecordType type, string value) : base(namespaceManager, null)
		{
			if (parentNode == null)
				throw new ArgumentNullException(nameof(parentNode));
			if (parentNode.LocalName != "sharedItems" && parentNode.LocalName != "r")
				throw new ArgumentException($"{nameof(parentNode)} type: '{parentNode.Name}' was not the expected type.");
			base.TopNode = parentNode.OwnerDocument.CreateElement(type.ToString(), parentNode.NamespaceURI);
			this.Type = type;
			if (!string.IsNullOrEmpty(value))
			{
				var attr = parentNode.OwnerDocument.CreateAttribute("v");
				base.TopNode.Attributes.Append(attr);
				this.Value = value;
			}
		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Replace <see cref="XmlNode"/> with new node type and value.
		/// </summary>
		/// <param name="type">The new <see cref="PivotCacheRecordType"/>.</param>
		/// <param name="value">The value.</param>
		/// <param name="parentNode">The parent <see cref="XmlNode"/>.</param>
		public void ReplaceNode(PivotCacheRecordType type, string value, XmlNode parentNode)
		{
			if (parentNode == null)
				throw new ArgumentNullException(nameof(parentNode));
			var newNode = parentNode.OwnerDocument.CreateElement(type.ToString(), parentNode.NamespaceURI);
			if (type != PivotCacheRecordType.m)
			{
				var attr = parentNode.OwnerDocument.CreateAttribute("v");
				newNode.Attributes.Append(attr);
			}
			parentNode.ReplaceChild(newNode, base.TopNode);
			base.TopNode = newNode;
			this.Type = type;
			this.Value = value;
		}
		#endregion

		#region Public Statics Methods
		/// <summary>
		/// Gets the type of the given value.
		/// </summary>
		/// <param name="value">The object to get the type of.</param>
		/// <returns>The item's type.</returns>
		public static PivotCacheRecordType GetObjectType(object value)
		{
			if (value is bool)
				return PivotCacheRecordType.b;
			else if (value is DateTime)
				return PivotCacheRecordType.d;
			else if (value is ExcelErrorValue)
				return PivotCacheRecordType.e;
			else if (value == null || (value is string stringValue && string.IsNullOrEmpty(stringValue)))
				return PivotCacheRecordType.m;
			else if (ConvertUtil.IsNumeric(value, true))
				return PivotCacheRecordType.n;
			else if (value is string)
				return PivotCacheRecordType.s;
			else
				throw new InvalidOperationException($"Unknown type of {value.GetType()}.");
		}
		#endregion
	}
}