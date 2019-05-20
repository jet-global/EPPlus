﻿/*******************************************************************************
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
		#region Properties
		/// <summary>
		/// Gets or sets the minimum date value in the shared items collection.
		/// Note: This property is mainly used for date groupings and date fields.
		/// </summary>
		public DateTime? MinDate
		{
			get
			{
				string value = base.GetXmlNodeString("@minDate");
				if (string.IsNullOrEmpty(value))
					return null;
				return DateTime.Parse(value);
			}
			set { base.SetXmlNodeString("@minDate", ConvertUtil.ConvertObjectToXmlAttributeString(value), true); }
		}

		/// <summary>
		/// Gets or sets the maximum date value in the shared items collection.
		/// Note: This property is mainly used for date groupings and date fields.
		/// </summary>
		public DateTime? MaxDate
		{
			get
			{
				string value = base.GetXmlNodeString("@maxDate");
				if (string.IsNullOrEmpty(value))
					return null;
				return DateTime.Parse(value);
			}
			set { base.SetXmlNodeString("@maxDate", ConvertUtil.ConvertObjectToXmlAttributeString(value), true); }
		}

		/// <summary>
		/// Gets or sets a value inducating that this field contains a mix of text values and other types.
		/// </summary>
		public bool ContainsSemiMixedTypes
		{
			get { return base.GetXmlNodeBool("@containsSemiMixedTypes", true); }
			set { base.SetXmlNodeBool("@containsSemiMixedTypes", value, true); }
		}

		/// <summary>
		/// Gets or sets a value indicating that this field contains values of more than one data type.
		/// </summary>
		public bool ContainsMixedTypes
		{
			get { return base.GetXmlNodeBool("@containsMixedTypes", false); }
			set { base.SetXmlNodeBool("@containsMixedTypes", value, false); }
		}

		/// <summary>
		/// Gets or sets a value indicating that this field contains text values longer than 255 characters.
		/// </summary>
		public bool LongText
		{
			get { return base.GetXmlNodeBool("@longText", false); }
			set { base.SetXmlNodeBool("@longText", value, false); }
		}

		/// <summary>
		///  Gets or sets a value indicating that this field contains string values.
		/// </summary>
		public bool ContainsString
		{
			get { return base.GetXmlNodeBool("@containsString", true); }
			set { base.SetXmlNodeBool("@containsString", value, true); }
		}

		/// <summary>
		///  Gets or sets a value indicating that this field contains numeric values.
		/// </summary>
		public bool ContainsNumbers
		{
			get { return base.GetXmlNodeBool("@containsNumber", false); }
			set { base.SetXmlNodeBool("@containsNumber", value, false); }
		}

		/// <summary>
		///  Gets or sets a value indicating that this field contains integer values.
		/// </summary>
		public bool ContainsInteger
		{
			get { return base.GetXmlNodeBool("@containsInteger", false); }
			set { base.SetXmlNodeBool("@containsInteger", value, false); }
		}

		/// <summary>
		/// Gets or sets a value indicating whether any of these shared items are blank.
		/// </summary>
		public bool ContainsBlank
		{
			get { return base.GetXmlNodeBool("@containsBlank", false); }
			set { base.SetXmlNodeBool("@containsBlank", value, false); }
		}

		/// <summary>
		///  Gets or sets a value indicating that this field contains non-date values.
		/// </summary>
		public bool ContainsNonDate
		{
			get { return base.GetXmlNodeBool("@containsNonDate", true); }
			set { base.SetXmlNodeBool("@containsNonDate", value, true); }
		}

		/// <summary>
		///  Gets or sets a value indicating that this field contains date values.
		/// </summary>
		public bool ContainsDate
		{
			get { return base.GetXmlNodeBool("@containsDate", false); }
			set { base.SetXmlNodeBool("@containsDate", value, false); }
		}
		#endregion

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
			if (item.Type == PivotCacheRecordType.m)
				this.ContainsBlank = true;
			base.AddItem(item);
		}

		/// <summary>
		/// Clears all items out of the collection.
		/// </summary>
		public void Clear() => base.ClearItems();
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