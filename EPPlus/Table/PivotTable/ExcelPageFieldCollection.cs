/*******************************************************************************
* You may amend and distribute as you like, but don't remove this header!
*
* EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
* See http://www.codeplex.com/EPPlus for details.
*
* Copyright (C) 2011-2018 Evan Schallerer, and others as noted in the source history.
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
	/// Represents an Excel Pivot Table Page Field collection XML element.
	/// </summary>
	public class ExcelPageFieldCollection : ExcelPivotTableFieldCollectionBase<ExcelPageField>
	{
		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="ExcelPageFieldCollection"/>.
		/// </summary>
		/// <param name="namespaceManager">The namespace manager.</param>
		/// <param name="node">The top xml node.</param>
		/// <param name="table">The existing pivot table.</param>
		public ExcelPageFieldCollection(XmlNamespaceManager namespaceManager, XmlNode node, ExcelPivotTable table) 
			: base(namespaceManager, node, table) { }
		#endregion

		#region ExcelPivotTableFieldCollectionBase Overrides
		/// <summary>
		/// Loads the initial collection of items from XML.
		/// </summary>
		/// <returns>A new list of page fields from the XML.</returns>
		protected override List<ExcelPageField> LoadItems()
		{
			var collection = new List<ExcelPageField>();
			var fields = base.TopNode.SelectNodes("d:pageField", base.NameSpaceManager);
			foreach (XmlNode xmlNode in fields)
			{
				collection.Add(new ExcelPageField(base.NameSpaceManager, xmlNode));
			}
			return collection;
		}
		#endregion
	}
}
