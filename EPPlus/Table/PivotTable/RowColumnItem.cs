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
using System.Collections;
using System.Collections.Generic;
using System.Xml;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Table.PivotTable
{
	/// <summary>
	/// A row or column item object. An <i/> node.
	/// </summary>
	public class RowColumnItem : XmlCollectionItemBase, IEnumerable<int>
	{
		#region Class Variables
		private List<int> myMemberPropertyIndexes;
		#endregion

		#region Properties
		/// <summary>
		/// Gets or sets the data field index.
		/// </summary>
		public int DataFieldIndex
		{
			get { return base.GetXmlNodeIntNull("@i") ?? 0; }
			set
			{
				string val = value == 0 ? string.Empty : value.ToString();
				base.SetXmlNodeString("@i", val, true);
			}
		}

		/// <summary>
		/// Gets or sets the repeated items count (@r).
		/// </summary>
		public int RepeatedItemsCount
		{
			get { return base.GetXmlNodeIntNull("@r") ?? 0; }
			set { base.SetXmlNodeString("@r", value.ToString()); }
		}

		/// <summary>
		/// Gets or sets the item type (@t).
		/// </summary>
		public string ItemType
		{
			get { return base.GetXmlNodeString("@t"); }
			set { base.SetXmlNodeString("@t", value, true); }
		}

		/// <summary>
		/// Gets the 'x' node's v atribute at the given index.
		/// </summary>
		/// <param name="Index">The position in the list.</param>
		/// <returns>An index into the pivotField items.</returns>
		public int this[int Index]
		{
			get
			{
				if (Index < 0 || Index >= this.MemberPropertyIndexes.Count)
					throw (new ArgumentOutOfRangeException("Index out of range"));
				return this.MemberPropertyIndexes[Index];
			}
		}

		/// <summary>
		/// Gets the count of the member property indexes.
		/// </summary>
		public int Count
		{
			get { return this.MemberPropertyIndexes.Count; }
		}

		private List<int> MemberPropertyIndexes
		{
			get
			{
				if (myMemberPropertyIndexes == null)
				{
					myMemberPropertyIndexes = new List<int>();
					foreach (XmlNode xmlNode in base.TopNode.ChildNodes)
					{
						var value = xmlNode.Attributes["v"]?.Value;
						int index = value == null ? 0 : int.Parse(value);
						myMemberPropertyIndexes.Add(index);
					}
				}
				return myMemberPropertyIndexes;
			}
		}
		#endregion

		#region Constructors
		/// <summary>
		/// Creates a new <see cref="RowColumnItem"/> object.
		/// </summary>
		/// <param name="namespaceManager">The namespace manager.</param>
		/// <param name="node">The item <see cref="XmlNode"/>.</param>
		public RowColumnItem(XmlNamespaceManager namespaceManager, XmlNode node) : base(namespaceManager, node)
		{
			if (node == null)
				throw new ArgumentNullException(nameof(node));
			if (node.LocalName != "i")
				throw new ArgumentException($"Invalid node type {node.LocalName}.");
		}

		/// <summary>
		/// Creates a new <see cref="RowColumnItem"/> object.
		/// </summary>
		/// <param name="namespaceManager">The namespace manager.</param>
		/// <param name="parentNode">The row/colItems xml node.</param>
		/// <param name="repeatedItemsCount">The value of the 'r' attribute.</param>
		/// <param name="memberIndex">The value of the 'x' child node.</param>
		/// <param name="itemType">The value of the 't' attribute.</param>
		/// <param name="dataFieldIndex">The 'i' attribute value which points to a data field.</param>
		public RowColumnItem(XmlNamespaceManager namespaceManager, XmlNode parentNode, int repeatedItemsCount, int memberIndex, string itemType = null, int dataFieldIndex = 0) : base(namespaceManager, null)
		{
			if (parentNode == null)
				throw new ArgumentNullException(nameof(parentNode));
			base.TopNode = parentNode.OwnerDocument.CreateElement("i", parentNode.NamespaceURI);
			if (repeatedItemsCount > 0)
				this.RepeatedItemsCount = repeatedItemsCount;
			this.DataFieldIndex = dataFieldIndex;
			var xNode = parentNode.OwnerDocument.CreateElement("x", base.TopNode.NamespaceURI);
			if (memberIndex > 0)
			{
				var attr = parentNode.OwnerDocument.CreateAttribute("v");
				xNode.Attributes.Append(attr);
				xNode.Attributes["v"].Value = memberIndex.ToString();
			}
			base.TopNode.AppendChild(xNode);
			if (itemType != null)
				this.ItemType = itemType;
		}

		/// <summary>
		/// Creates a new <see cref="RowColumnItem"/> object.
		/// </summary>
		/// <param name="namespaceManager">The namespace manager.</param>
		/// <param name="parentNode">The row/colItems xml node.</param>
		/// <param name="memberIndices">The list of member property indices ('x' attributes).</param>
		/// <param name="repeatedItemsCount">The 'x' attribute value.</param>
		/// <param name="dataFieldIndex">The 'i' attribute value which points to a data field.</param>
		public RowColumnItem(XmlNamespaceManager namespaceManager, XmlNode parentNode, List<Tuple<int, int>> memberIndices, int repeatedItemsCount, int dataFieldIndex) : base(namespaceManager, null)
		{
			if (parentNode == null)
				throw new ArgumentNullException(nameof(parentNode));
			if (memberIndices == null || memberIndices.Count == 0)
				throw new ArgumentNullException(nameof(memberIndices));
			base.TopNode = parentNode.OwnerDocument.CreateElement("i", parentNode.NamespaceURI);
			if (repeatedItemsCount > 0)
				this.RepeatedItemsCount = repeatedItemsCount;
			this.DataFieldIndex = dataFieldIndex;
			for (int i = 0; i < memberIndices.Count; i++)
			{
				var xNode = parentNode.OwnerDocument.CreateElement("x", base.TopNode.NamespaceURI);
				var attr = parentNode.OwnerDocument.CreateAttribute("v");
				xNode.Attributes.Append(attr);
				xNode.Attributes["v"].Value = memberIndices[i].Item2.ToString();
				base.TopNode.AppendChild(xNode);
			}
		}
		#endregion

		#region IEnumerable Methods
		/// <summary>
		/// Gets the int enumerator of the list.
		/// </summary>
		/// <returns>The enumerator.</returns>
		public IEnumerator<int> GetEnumerator()
		{
			return myMemberPropertyIndexes.GetEnumerator();
		}

		/// <summary>
		/// Gets the specified type enumerator of the list.
		/// </summary>
		/// <returns>The enumerator.</returns>
		IEnumerator IEnumerable.GetEnumerator()
		{
			return myMemberPropertyIndexes.GetEnumerator();
		}
		#endregion
	}
}