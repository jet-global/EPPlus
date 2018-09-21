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
		private List<CacheRecordItem> myItems = new List<CacheRecordItem>();
		#endregion

		#region Properties
		/// <summary>
		/// Gets a readonly list of the items in this <see cref="CacheRecordNode"/>.
		/// </summary>
		public IReadOnlyList<CacheRecordItem> Items
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
		public CacheRecordNode(XmlNode node, XmlNamespaceManager namespaceManager)
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
				myItems.Add(new CacheRecordItem(cacheRecordItem));
			}
		}
		#endregion
	}
}