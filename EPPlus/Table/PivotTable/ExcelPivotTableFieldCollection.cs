/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See http://www.codeplex.com/EPPlus for details.
 *
 * Copyright (C) 2011  Jan Källman
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
 * Code change notes:
 * 
 * Author							Change						Date
 * ******************************************************************************
 * Jan Källman		Added		21-MAR-2011
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using System.Collections.Generic;
using System.Xml;
using OfficeOpenXml.Extensions;

namespace OfficeOpenXml.Table.PivotTable
{
	/// <summary>
	/// Collection class for fields in a pivot table.
	/// </summary>
	public class ExcelPivotTableFieldCollection : ExcelPivotTableFieldCollectionBase<ExcelPivotTableField>
	{
		#region Properties
		/// <summary>
		/// Gets the field accessed by the name.
		/// </summary>
		/// <param name="name">The name of the field.</param>
		/// <returns>The specified field or null if it does not exist.</returns>
		public ExcelPivotTableField this[string name]
		{
			get
			{
				foreach (var field in this)
				{
					if (field.Name.IsEquivalentTo(name))
						return field;
				}
				return null;
			}
		}
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="ExcelPivotTableFieldCollection"/>.
		/// </summary>
		/// <param name="namespaceManager">The namespace manager.</param>
		/// <param name="node">The top xml node.</param>
		/// <param name="table">The existing pivot table.</param>
		internal ExcelPivotTableFieldCollection(XmlNamespaceManager namespaceManager, XmlNode node, ExcelPivotTable table) 
			: base(namespaceManager, node, table) { }
		#endregion

		#region Public Methods
		/// <summary>
		/// Returns the date group field.
		/// </summary>
		/// <param name="groupBy">The type of grouping.</param>
		/// <returns>The matching field or null if none is found.</returns>
		public ExcelPivotTableField GetDateGroupField(eDateGroupBy groupBy)
		{
			foreach (var field in this)
			{
				if (field.Grouping is ExcelPivotTableFieldDateGroup && (((ExcelPivotTableFieldDateGroup)field.Grouping).GroupBy) == groupBy)
					return field;
			}
			return null;
		}

		/// <summary>
		/// Returns the numeric group field.
		/// </summary>
		/// <returns>The matching field or null if none is found.</returns>
		public ExcelPivotTableField GetNumericGroupField()
		{
			foreach (var field in this)
			{
				if (field.Grouping is ExcelPivotTableFieldNumericGroup)
					return field;
			}
			return null;
		}

		/// <summary>
		/// Adds the specified <paramref name="field"/> to the collection.
		/// </summary>
		/// <param name="field">The field being added.</param>
		/// <remarks>This does not add the field to the XML.</remarks>
		public void Add(ExcelPivotTableField field)
		{
			base.AddItem(field);
		}
		#endregion

		#region ExcelPivotTableFieldCollectionBase Overrides
		/// <summary>
		/// Loads all the <see cref="ExcelPivotTableFieldItem"/> from the xml document.
		/// </summary>
		/// <returns>The collection of <see cref="ExcelPivotTableFieldItem"/>s.</returns>
		protected override List<ExcelPivotTableField> LoadItems()
		{
			var collection = new List<ExcelPivotTableField>();
			int index = 0;
			foreach (XmlNode fieldNode in base.TopNode.SelectNodes("d:pivotField", base.NameSpaceManager))
			{
				var field = new ExcelPivotTableField(base.NameSpaceManager, fieldNode, base.PivotTable, index, index++);
				collection.Add(field);
			}
			return collection;
		}
		#endregion
	}
}