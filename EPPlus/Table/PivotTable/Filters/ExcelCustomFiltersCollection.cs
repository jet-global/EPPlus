/*******************************************************************************
* You may amend and distribute as you like, but don't remove this header!
*
* EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
* See http://www.codeplex.com/EPPlus for details.
*
* Copyright (C) 2011-2018 Michelle Lau and others as noted in the source history.
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

namespace OfficeOpenXml.Table.PivotTable.Filters
{
	/// <summary>
	/// A collection of custom <see cref="ExcelFilter"/>s.
	/// </summary>
	public class ExcelCustomFiltersCollection : XmlCollectionBase<ExcelFilter>
	{
		#region Properties
		/// <summary>
		/// Gets a flag indicating whether or not the two filter criterias should have an "and" relationship.
		/// Remarks: A value of '1' indicates "and", '0' indicates "or".
		/// </summary>
		public bool And
		{
			get { return base.GetXmlNodeBool("@and"); }
		}
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of an <see cref="ExcelCustomFiltersCollection"/>.
		/// </summary>
		/// <param name="namespaceManager">The namespace manager.</param>
		/// <param name="node">The xml top node.</param>
		public ExcelCustomFiltersCollection(XmlNamespaceManager namespaceManager, XmlNode node) : base (namespaceManager, node)
		{
			if (node == null)
				throw new ArgumentNullException(nameof(node));
		}
		#endregion

		#region XmlCollectionBase Overrides
		/// <summary>
		/// Loads the custom filters from the xml document.
		/// </summary>
		/// <remarks>There can be atmost two custom filters in this collection.</remarks>
		/// <returns>The collection of custom filters.</returns>
		protected override List<ExcelFilter> LoadItems()
		{
			var items = new List<ExcelFilter>();
			// Selects all possible child node types.
			foreach (XmlNode customFilter in base.TopNode.SelectNodes("d:customFilter", this.NameSpaceManager))
			{
				items.Add(new ExcelFilter(this.NameSpaceManager, customFilter));
			}
			return items;
		}
		#endregion
	}
}
