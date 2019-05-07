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

namespace OfficeOpenXml.Table.PivotTable.Formats
{
	/// <summary>
	/// The collection of pivot areas.
	/// </summary>
	public class PivotAreasCollection : XmlCollectionBase<PivotArea>
	{
		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="PivotAreasCollection"/>.
		/// </summary>
		/// <param name="namespaceManager">The namespace manager.</param>
		/// <param name="node">The xml top node.</param>
		public PivotAreasCollection(XmlNamespaceManager namespaceManager, XmlNode node) : base(namespaceManager, node)
		{
			if (node == null)
				throw new ArgumentNullException(nameof(node));
		}
		#endregion

		#region XmlCollectionBase Overrides
		/// <summary>
		/// Loads the <see cref="PivotArea"/>s from the xml document.
		/// </summary>
		/// <returns>The collection of <see cref="PivotArea"/>s.</returns>
		protected override List<PivotArea> LoadItems()
		{
			var pivotAreaCollection = new List<PivotArea>();
			foreach (XmlNode pivotArea in base.TopNode.SelectNodes("d:pivotArea", this.NameSpaceManager))
			{
				pivotAreaCollection.Add(new PivotArea(this.NameSpaceManager, pivotArea));
			}
			return pivotAreaCollection;
		}
		#endregion
	}
}
