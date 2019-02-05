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
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable
{
	/// <summary>
	/// Collection class for <see cref="AutoSortScopeReference"/>.
	/// </summary>
	public class AutoSortScopeReferencesCollection : ExcelPivotTableFieldCollectionBase<AutoSortScopeReference>
	{
		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="AutoSortScopeReferencesCollection"/>.
		/// </summary>
		/// <param name="namespaceManager">The namespace manager.</param>
		/// <param name="node">The xml node.</param>
		/// <param name="table">The existing pivot table.</param>
		public AutoSortScopeReferencesCollection(XmlNamespaceManager namespaceManager, XmlNode node, ExcelPivotTable table) : base(namespaceManager, node, table) { }
		#endregion

		#region XmlCollectionBase Overrides
		/// <summary>
		/// Loads the <see cref="AutoSortScopeReference"/>s from the xml document.
		/// </summary>
		/// <returns>The collection of <see cref="AutoSortScopeReference"/>s.</returns>
		protected override List<AutoSortScopeReference> LoadItems()
		{
			var references = new List<AutoSortScopeReference>();
			foreach (XmlNode reference in base.TopNode.SelectNodes("d:reference", this.NameSpaceManager))
			{
				references.Add(new AutoSortScopeReference(this.NameSpaceManager, reference.FirstChild));
			}
			return references;
		}
		#endregion
	}
}
