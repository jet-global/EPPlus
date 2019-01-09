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
	/// The type of a pivot table item.
	/// </summary>
	internal enum PivotTableItemType
	{
		/// <summary>
		/// The row field.
		/// </summary>
		Row,
		/// <summary>
		/// The column field.
		/// </summary>
		Column
	}
	#endregion

	/// <summary>
	/// Collection class for row and column fields in a pivot table.
	/// </summary>
	public class ExcelPivotTableRowColumnFieldCollection : ExcelPivotTableFieldCollectionBase<ExcelPivotTableField>
	{
		#region Properties
		/// <summary>
		/// Gets the <see cref="PivotTableItemType"/>.
		/// </summary>
		internal PivotTableItemType FieldType { get; }
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="ExcelPivotTableRowColumnFieldCollection"/>.
		/// </summary>
		/// <param name="namespaceManager">The namespace manager.</param>
		/// <param name="node">The top xml node.</param>
		/// <param name="table">The existing pivot table.</param>
		/// <param name="type">The <see cref="PivotTableItemType"/>.</param>
		internal ExcelPivotTableRowColumnFieldCollection(XmlNamespaceManager namespaceManager, XmlNode node, ExcelPivotTable table, PivotTableItemType type) 
			: base(namespaceManager, node, table)
		{
			if (node == null)
			{
				if (type == PivotTableItemType.Row)
					base.TopNode = base.PivotTable.CreateNode("d:rowFields");
				else if (type == PivotTableItemType.Column)
					base.TopNode = base.PivotTable.CreateNode("d:colFields");
				else
					throw new InvalidOperationException($"The enum value '{this.FieldType}' Is not supported.");
			}
			this.FieldType = type;
		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Add a new row/column field.
		/// </summary>
		/// <param name="field">The field.</param>
		/// <returns>The new field.</returns>
		public ExcelPivotTableField Add(ExcelPivotTableField field)
		{
			this.SetFlag(field, true);
			base.AddItem(field);
			return field;
		}

		/// <summary>
		/// Remove a field from the pivot table.
		/// </summary>
		/// <param name="field">The field that is being removed.</param>
		public void Remove(ExcelPivotTableField field)
		{
			if (!base.ContainsItem(field))
				throw new ArgumentException("Field not in collection");
			this.SetFlag(field, false);
			base.RemoveItem(field);
		}

		/// <summary>
		/// Insert a new row/column field.
		/// </summary>
		/// <param name="field">The field.</param>
		/// <param name="index">The position to insert the field.</param>
		/// <returns>The new field.</returns>
		internal ExcelPivotTableField Insert(ExcelPivotTableField field, int index)
		{
			this.SetFlag(field, true);
			base.InsertItem(index, field);
			return field;
		}
		#endregion

		#region ExcelPivotTableFieldCollectionBase
		/// <summary>
		/// Loads the row/column or page fields from the xml document.
		/// </summary>
		/// <returns>The collection of fields.</returns>
		protected override List<ExcelPivotTableField> LoadItems()
		{
			if (this.FieldType == PivotTableItemType.Row || this.FieldType == PivotTableItemType.Column)
				return this.LoadRowColumnFields();
			else
				throw new InvalidOperationException($"The enum value '{this.FieldType}' is not supported as a pivot table field type.");
		}
		#endregion

		#region Private Methods
		private void SetFlag(ExcelPivotTableField field, bool value)
		{
			switch (this.FieldType)
			{
				case PivotTableItemType.Row:
					if (field.IsColumnField)
						throw (new Exception("This field is a column or page field. Can't add it to the RowFields collection"));
					field.IsRowField = value;
					field.Axis = ePivotFieldAxis.Row;
					break;
				case PivotTableItemType.Column:
					if (field.IsRowField)
						throw (new Exception("This field is a row or page field. Can't add it to the ColumnFields collection"));
					field.IsColumnField = value;
					field.Axis = ePivotFieldAxis.Column;
					break;
				default:
					throw new InvalidOperationException($"The enum value '{this.FieldType}' has not been accounted for.");
			}
		}

		private List<ExcelPivotTableField> LoadRowColumnFields()
		{
			var collection = new List<ExcelPivotTableField>();
			var fieldNodes = base.TopNode.SelectNodes("d:field", base.NameSpaceManager);
			if (fieldNodes == null)
				return collection;
			foreach (XmlElement element in fieldNodes)
			{
				if (int.TryParse(element.GetAttribute("x"), out var x)) {
					if (x >= 0)
						collection.Add(base.PivotTable.Fields[x]);
					else
					{
						// TODO (Task #8178): Figure out what the base index is. (Possibly related to grouping)
						collection.Add(new ExcelPivotTableField(base.NameSpaceManager, base.TopNode, this.PivotTable, x, 0));
					}
				}
				else
				{
					// If it doesn't have an 'x' attribute, remove the element.
					// We have no idea why this is here, but it is legacy logic that we're afraid to remove.
					base.TopNode.RemoveChild(element);
				}
			}
			return collection;
		}
		#endregion
	}
}