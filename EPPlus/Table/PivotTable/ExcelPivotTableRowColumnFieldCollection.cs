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
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable
{
	/// <summary>
	/// Collection class for row and column fields in a pivot table.
	/// </summary>
	public class ExcelPivotTableRowColumnFieldCollection : ExcelPivotTableFieldCollectionBase<ExcelPivotTableField>
	{
		#region Class Variables
		/// <summary>
		/// The top node.
		/// </summary>
		internal string myTopNode;
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="ExcelPivotTableRowColumnFieldCollection"/>.
		/// </summary>
		/// <param name="table">The existing pivot table.</param>
		/// <param name="topNode">The text of the top node in the xml.</param>
		internal ExcelPivotTableRowColumnFieldCollection(ExcelPivotTable table, string topNode) :
			 base(table)
		{
			myTopNode = topNode;
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
			myList.Add(field);
			return field;
		}

		/// <summary>
		/// Remove a field from the pivot table.
		/// </summary>
		/// <param name="field">The field that is being removed.</param>
		public void Remove(ExcelPivotTableField field)
		{
			if (!myList.Contains(field))
				throw new ArgumentException("Field not in collection");
			this.SetFlag(field, false);
			myList.Remove(field);
		}

		/// <summary>
		/// Remove a field at a specific position.
		/// </summary>
		/// <param name="index">The position of the target field.</param>
		public void RemoveAt(int index)
		{
			if (index > -1 && index < myList.Count)
				throw (new IndexOutOfRangeException());
			this.SetFlag(myList[index], false);
			myList.RemoveAt(index);
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
			myList.Insert(index, field);
			return field;
		}

		/// <summary>
		/// Populate the <see cref="ExcelPivotTableFieldCollection"/> with row and column fields.
		/// </summary>
		/// <param name="namespaceManager">The namespace manager.</param>
		/// <param name="parentNode">The top node.</param>
		/// <param name="nodePath">The path to the field nodes.</param>
		/// <param name="fields">The <see cref="ExcelPivotTableFieldCollection"/> of fields.</param>
		internal void PopulateRowColumnFields(XmlNamespaceManager namespaceManager, XmlNode parentNode, string nodePath, ExcelPivotTableFieldCollection fields)
		{
			foreach (XmlElement element in parentNode.SelectNodes(nodePath, namespaceManager))
			{
				if (int.TryParse(element.GetAttribute("x"), out var x) && x >= 0)
					this.AddInternal(fields[x]);
				else
					element.ParentNode.RemoveChild(element);
			}
		}

		internal void PopulatePageFields(XmlNamespaceManager namespaceManager, XmlNode parentNode, string nodePath, ExcelPivotTableFieldCollection fields)
		{
			foreach (XmlElement pageElem in parentNode.SelectNodes(nodePath, namespaceManager))
			{
				if (int.TryParse(pageElem.GetAttribute("fld"), out var fld) && fld >= 0)
				{
					var field = fields[fld];
					field.myPageFieldSettings = new ExcelPivotTablePageFieldSettings(namespaceManager, pageElem, field, fld);
					this.AddInternal(field);
				}
			}
		}
		#endregion

		#region Private Methods
		private void SetFlag(ExcelPivotTableField field, bool value)
		{
			switch (myTopNode)
			{
				case "rowFields":
					if (field.IsColumnField || field.IsPageField)
						throw (new Exception("This field is a column or page field. Can't add it to the RowFields collection"));
					field.IsRowField = value;
					field.Axis = ePivotFieldAxis.Row;
					break;
				case "colFields":
					if (field.IsRowField || field.IsPageField)
						throw (new Exception("This field is a row or page field. Can't add it to the ColumnFields collection"));
					field.IsColumnField = value;
					field.Axis = ePivotFieldAxis.Column;
					break;
				case "pageFields":
					if (field.IsColumnField || field.IsRowField)
						throw (new Exception("Field is a column or row field. Can't add it to the PageFields collection"));
					if (myTable.Address._fromRow < 3)
						throw (new Exception(string.Format("A pivot table with page fields must be located above row 3. Currenct location is {0}", myTable.Address.Address)));
					field.IsPageField = value;
					field.Axis = ePivotFieldAxis.Page;
					break;
				case "dataFields":
					break;
			}
		}
		#endregion
	}
}