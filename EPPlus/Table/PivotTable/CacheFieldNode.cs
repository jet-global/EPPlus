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
	/// <summary>
	/// Wraps a <cacheField/> node in <pivotcachedefinition-cacheFields/>.
	/// </summary>
	public class CacheFieldNode
	{
		#region Class Variables
		private List<CacheFieldItem> myItems = new List<CacheFieldItem>();
		#endregion

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
		/// Gets a readonly list of the items in this <see cref="CacheFieldNode"/>.
		/// </summary>
		public IReadOnlyList<CacheFieldItem> Items
		{
			get { return myItems; }
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
		public CacheFieldNode(XmlNode node, XmlNamespaceManager namespaceManager)
		{
			if (node == null)
				throw new ArgumentNullException(nameof(node));
			if (namespaceManager == null)
				throw new ArgumentNullException(nameof(namespaceManager));
			this.Node = node;
			this.NameSpaceManager = namespaceManager;
			foreach (XmlNode cacheFieldItem in this.Node.SelectNodes("d:sharedItems/d:s", this.NameSpaceManager))
			{
				myItems.Add(new CacheFieldItem(cacheFieldItem));
			}
		}
		#endregion
	}
}