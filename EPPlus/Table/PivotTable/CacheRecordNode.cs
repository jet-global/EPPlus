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
using System.Linq;
using System.Xml;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Table.PivotTable
{
	#region Enums
	/// <summary>
	/// The possible types of values in a record.
	/// </summary>
	public enum PivotCacheRecordType
	{
		/// <summary>
		/// A boolean type.
		/// </summary>
		b,
		/// <summary>
		/// A date time type.
		/// </summary>
		d,
		/// <summary>
		/// An error value type.
		/// </summary>
		e,
		/// <summary>
		/// A no value type.
		/// </summary>
		m,
		/// <summary>
		/// A numeric type.
		/// </summary>
		n,
		/// <summary>
		/// A character value type.
		/// </summary>
		s,
		/// <summary>
		/// A shared items index type.
		/// </summary>
		x
	}
	#endregion

	/// <summary>
	/// Wraps a <r/> node in <pivotCacheRecords/>.
	/// </summary>
	public class CacheRecordNode
	{
		#region Class Variables
		private List<CacheItem> myItems = new List<CacheItem>();
		#endregion

		#region Properties
		/// <summary>
		/// Gets a readonly list of the items in this <see cref="CacheRecordNode"/>.
		/// </summary>
		public IReadOnlyList<CacheItem> Items
		{
			get { return myItems; }
		}

		private XmlNode Node { get; set; }
		private XmlNamespaceManager NameSpaceManager { get; set; }
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="CacheRecordNode"/>.
		/// </summary>
		/// <param name="node">The <see cref="XmlNode"/> for this <see cref="CacheRecordNode"/>.</param>
		/// <param name="namespaceManager">The namespace manager to use for searching child nodes.</param>
		public CacheRecordNode(XmlNamespaceManager namespaceManager, XmlNode node)
		{
			if (node == null)
				throw new ArgumentNullException(nameof(node));
			if (namespaceManager == null)
				throw new ArgumentNullException(nameof(namespaceManager));
			this.Node = node;
			this.NameSpaceManager = namespaceManager;
			// Selects all possible child node types.
			foreach (XmlNode cacheRecordItem in this.Node.SelectNodes("d:b | d:d | d:e | d:m | d:n | d:s | d:x", this.NameSpaceManager))
			{
				myItems.Add(new CacheItem(this.NameSpaceManager, cacheRecordItem));
			}
		}

		/// <summary>
		/// Creates a new <see cref="CacheRecordNode"/> and items as specified by the <paramref name="row"/> values.
		/// Adds the resulting <see cref="CacheRecordNode"/> to the specified <paramref name="parentNode"/>.
		/// </summary>
		/// <param name="namespaceManager">The namespace manager.</param>
		/// <param name="parentNode">The parent <see cref="ExcelPivotCacheRecords"/> <see cref="XmlNode"/>.</param>
		/// <param name="row">A list of object values that this node represents.</param>
		/// <param name="cacheDefinition">The parent <see cref="ExcelPivotCacheDefinition"/>.</param>
		public CacheRecordNode(XmlNamespaceManager namespaceManager, XmlNode parentNode, IEnumerable<object> row, ExcelPivotCacheDefinition cacheDefinition)
		{
			if (parentNode == null)
				throw new ArgumentNullException(nameof(parentNode));
			if (namespaceManager == null)
				throw new ArgumentNullException(nameof(namespaceManager));
			if (row == null)
				throw new ArgumentNullException(nameof(row));
			if (cacheDefinition == null)
				throw new ArgumentNullException(nameof(cacheDefinition));
			if (row.Count() != cacheDefinition.CacheFields.Count)
				throw new InvalidOperationException("An attempt was made to create a CacheRecord node with an invalid number of fields.");
			this.NameSpaceManager = namespaceManager;
			var recordNode = parentNode.OwnerDocument.CreateElement("d:r");
			int col = 0;
			foreach (var value in row)
			{
				var type = CacheItem.GetObjectType(value);
				var cacheField = cacheDefinition.CacheFields[col];
				if (cacheField.HasSharedItems)
				{
					// The corresponding cacheField has shared items; map the new cacheRecord entry 
					// into shared items if a matching entry exists, otherwise create a new sharedItem entry and map accordingly.
					var indexStringValue = this.GetCacheFieldSharedItemIndexString(cacheField, type, value);
					var item = new CacheItem(namespaceManager, recordNode, PivotCacheRecordType.x, indexStringValue);
					item.AddSelf(recordNode);
					myItems.Add(item);
				}
				else
				{
					// If no SharedItems exist, simply create a record item entry.
					var stringValue = ConvertUtil.ConvertObjectToXmlAttributeString(value);
					var item = new CacheItem(namespaceManager, recordNode, type, stringValue);
					item.AddSelf(recordNode);
					myItems.Add(item);
				}
				col++;
			}
			parentNode.AppendChild(recordNode);
		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Update the existing <see cref="CacheRecordNode"/>.
		/// </summary>
		/// <param name="row">The row of data from the source table.</param>
		/// <param name="cacheDefinition">The cacheDefinition.</param>
		public void Update(IEnumerable<object> row, ExcelPivotCacheDefinition cacheDefinition)
		{
			if (row == null)
				throw new ArgumentNullException(nameof(row));
			if (cacheDefinition == null)
				throw new ArgumentNullException(nameof(cacheDefinition));
			if (row.Count() != this.Items.Count)
				throw new InvalidOperationException("An attempt was made to update a CacheRecordNode with a different number of fields.");
			int col = 0;
			foreach (var value in row)
			{
				var type = CacheItem.GetObjectType(value);
				var currentItem = myItems[col];
				var cacheField = cacheDefinition.CacheFields[col];
				if (cacheField.HasSharedItems)
				{
					// If shared items contains value, update this.Value to index
					// otherwise, create and add new sharedItem, update this.Value to new index
					currentItem.Value = this.GetCacheFieldSharedItemIndexString(cacheField, type, value);
				}
				else
				{
					// If only the value changed, update it. If the type changed,
					// replace the node with one of the correct type and value.
					string stringValue = ConvertUtil.ConvertObjectToXmlAttributeString(value);
					if (currentItem.Type == type && currentItem.Value != stringValue)
						currentItem.Value = stringValue;
					else if (currentItem.Type != type)
						currentItem.ReplaceNode(type, stringValue, this.Node);
				}
				col++;
			}
		}

		/// <summary>
		/// Removes this child node from the specified <paramref name="parentNode"/>.
		/// </summary>
		/// <param name="parentNode">The parent xml node.</param>
		public void Remove(XmlNode parentNode)
		{
			parentNode.RemoveChild(this.Node);
		}
		#endregion

		#region Private Methods
		private string GetCacheFieldSharedItemIndexString(CacheFieldNode cacheField, PivotCacheRecordType type, object value)
		{
			int cacheFieldItemIndex = cacheField.GetSharedItemIndex(type, value);
			// Adds a new sharedItem if the item does not exist.
			if (cacheFieldItemIndex < 0)
			{
				cacheField.SharedItems.Add(value);
				cacheFieldItemIndex = cacheField.SharedItems.Count - 1;
			}
			return ConvertUtil.ConvertObjectToXmlAttributeString(cacheFieldItemIndex);
		}
		#endregion
	}
}