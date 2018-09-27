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
using OfficeOpenXml.Extensions;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Table.PivotTable
{
	/// <summary>
	/// Wraps a node in <pivotCacheRecords-x/>.
	/// </summary>
	public class CacheRecordItem : XmlHelper
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
			private set
			{
				if (this.Type == PivotCacheRecordType.m)
					value = null;
				base.SetXmlNodeString("@v", value, true);
			}
		}
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="CacheRecordItem"/>.
		/// </summary>
		/// <param name="node">The <see cref="XmlNode"/> for this <see cref="CacheRecordItem"/>.</param>
		/// <param name="namespaceManager">The namespace manger.</param>
		public CacheRecordItem(XmlNode node, XmlNamespaceManager namespaceManager) : base(namespaceManager, node)
		{
			if (node == null)
				throw new ArgumentNullException(nameof(node));
			if (namespaceManager == null)
				throw new ArgumentNullException(nameof(namespaceManager));
			this.Type = (PivotCacheRecordType)Enum.Parse(typeof(PivotCacheRecordType), node.Name);
		}
		#endregion

		#region Private Methods
		private PivotCacheRecordType GetObjectType(object value)
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

		private void ReplaceNode(PivotCacheRecordType type, XmlNode parentNode)
		{
			var newNode = parentNode.OwnerDocument.CreateElement(type.ToString(), parentNode.NamespaceURI);
			if (type != PivotCacheRecordType.m)
			{
				var attr = parentNode.OwnerDocument.CreateAttribute("v");
				newNode.Attributes.Append(attr);
			}
			parentNode.ReplaceChild(newNode, base.TopNode);
			base.TopNode = newNode;
			this.Type = type;
		}
		#endregion
	}
}