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

namespace OfficeOpenXml.Table.PivotTable
{
	/// <summary>
	/// A row or column item object.
	/// </summary>
	public class RowColumnItem : XmlHelper
	{
		#region Class Variables
		private List<int> myMemberPropertyIndexes;
		#endregion

		#region Properties
		/// <summary>
		/// Gets the list of member property indexes.
		/// </summary>
		public IReadOnlyList<int> MemberPropertyIndex
		{
			get
			{
				if (myMemberPropertyIndexes == null)
				{
					myMemberPropertyIndexes = new List<int>();
					var xNodes = base.TopNode.SelectNodes("d:x", base.NameSpaceManager);
					foreach (XmlNode xmlNode in xNodes)
					{
						var value = xmlNode.Attributes["v"]?.Value;
						int index = value == null ? 0 : int.Parse(value);
						myMemberPropertyIndexes.Add(index);
					}
				}
				return myMemberPropertyIndexes;
			}
		}

		/// <summary>
		/// Gets or sets the data field index.
		/// </summary>
		public int DataFieldIndex
		{
			get { return base.GetXmlNodeIntNull("@i") ?? 0; }
			set { base.SetXmlNodeString("@i", value.ToString()); }
		}

		/// <summary>
		/// Gets or sets the repeated items count.
		/// </summary>
		public int RepeatedItemsCount
		{
			get { return base.GetXmlNodeIntNull("@r") ?? 0; }
			set { base.SetXmlNodeString("@r", value.ToString()); }
		}

		/// <summary>
		/// Gets or sets the item type.
		/// </summary>
		public string ItemType
		{
			get { return base.GetXmlNodeString("@t"); }
			set { base.SetXmlNodeString("@t", value, true); }
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
			if (namespaceManager == null)
				throw new ArgumentNullException(nameof(namespaceManager));
		}
		#endregion
	}
}