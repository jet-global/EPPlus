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
	/// Collection class for <see cref="ExcelPivotTableFieldItem"/>.
	/// </summary>
	public class ExcelPivotTableFieldItemCollection : ExcelPivotTableFieldCollectionBase<ExcelPivotTableFieldItem>
	{
		#region Properties
		/// <summary>
		/// Gets the <see cref="ExcelPivotTableField"/> this collection is a part of.
		/// </summary>
		public ExcelPivotTableField Field { get; }
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="ExcelPivotTableFieldCollection"/>.
		/// </summary>
		/// <param name="namespaceManager">The namespace manager.</param>
		/// <param name="topNode">The xml node.</param>
		/// <param name="table">The existing pivot table.</param>
		/// <param name="field">The <see cref="ExcelPivotTableField"/> of this collection.</param>
		public ExcelPivotTableFieldItemCollection(XmlNamespaceManager namespaceManager, XmlNode topNode, ExcelPivotTable table, ExcelPivotTableField field) : base(namespaceManager, topNode, table)
		{
			if (field == null)
				throw new ArgumentNullException(nameof(field));
			this.Field = field;
		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Adds a new <see cref="ExcelPivotTableFieldItem"/> to the collection.
		/// </summary>
		/// <param name="pivotFieldIndex">The item's @x attribute value.</param>
		/// <param name="defaultSubtotal">A flag indicating if there exists an item with a non-null @t attribute.</param>
		public void AddItem(int pivotFieldIndex, bool defaultSubtotal)
		{
			var item = new ExcelPivotTableFieldItem(base.NameSpaceManager, base.TopNode, this.Field, pivotFieldIndex);
			if (defaultSubtotal)
				base.InsertItem(pivotFieldIndex, item);
			else
				base.AddItem(item);
		}
		#endregion

		#region ExcelPivotTableFieldCollectionBase Overrides
		/// <summary>
		/// Loads the <see cref="ExcelPivotTableFieldItem"/>s from the xml document.
		/// </summary>
		/// <returns>The collection of <see cref="ExcelPivotTableFieldItem"/>s.</returns>
		protected override List<ExcelPivotTableFieldItem> LoadItems()
		{
			var collection = new List<ExcelPivotTableFieldItem>();
			var items = base.TopNode.SelectNodes("d:item", base.NameSpaceManager);
			foreach (XmlNode xmlNode in items)
			{
				collection.Add(new ExcelPivotTableFieldItem(base.NameSpaceManager, xmlNode, this.Field));
			}
			return collection;
		}
		#endregion
	}
}