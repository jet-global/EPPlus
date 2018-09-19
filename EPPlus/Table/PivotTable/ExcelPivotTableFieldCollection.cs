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
using System;
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable
{
	/// <summary>
	/// Base collection class for pivot table fields.
	/// </summary>
	/// <typeparam name="T"></typeparam>
	public class ExcelPivotTableFieldCollectionBase<T> : IEnumerable<T>
	{
		#region Class Variables
		/// <summary>
		/// The pivot table.
		/// </summary>
		protected ExcelPivotTable myTable;
		
		/// <summary>
		/// A list of fields.
		/// </summary>
		internal List<T> myList = new List<T>();
		#endregion

		#region Properties
		/// <summary>
		/// Gets the enumerator of the list.
		/// </summary>
		/// <returns>The enumerator.</returns>
		public IEnumerator<T> GetEnumerator()
		{
			return myList.GetEnumerator();
		}

		System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
		{
			return myList.GetEnumerator();
		}

		/// <summary>
		/// Gets the count of fields in the pivot table.
		/// </summary>
		public int Count
		{
			get
			{
				return myList.Count;
			}
		}

		/// <summary>
		/// Gets the field at the given index.
		/// </summary>
		/// <param name="Index">The position of the field in the list.</param>
		/// <returns>The pivot table field.</returns>
		public T this[int Index]
		{
			get
			{
				if (Index < 0 || Index >= myList.Count)
					throw (new ArgumentOutOfRangeException("Index out of range"));
				return myList[Index];
			}
		}
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="ExcelPivotTableFieldCollection"/>.
		/// </summary>
		/// <param name="table">The existing pivot table.</param>
		internal ExcelPivotTableFieldCollectionBase(ExcelPivotTable table)
		{
			myTable = table;
		}
		#endregion

		#region Methods
		/// <summary>
		/// Adds a field to the collection.
		/// </summary>
		/// <param name="field"></param>
		internal void AddInternal(T field)
		{
			myList.Add(field);
		}

		/// <summary>
		/// Clears the field collection.
		/// </summary>
		internal void Clear()
		{
			myList.Clear();
		}
		#endregion
	}

	/// <summary>
	/// A collection of pivot table field objects.
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
				foreach (var field in myList)
				{
					if (field.Name.Equals(name, StringComparison.InvariantCultureIgnoreCase))
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
		/// <param name="table">The existing pivot table.</param>
		/// <param name="topNode">The text of the top node in the xml.</param>
		internal ExcelPivotTableFieldCollection(ExcelPivotTable table, string topNode) :
			 base(table)
		{

		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Returns the date group field.
		/// </summary>
		/// <param name="groupBy">The type of grouping.</param>
		/// <returns>The matching field or null if none is found.</returns>
		public ExcelPivotTableField GetDateGroupField(eDateGroupBy groupBy)
		{
			foreach (var fld in myList)
			{
				if (fld.Grouping is ExcelPivotTableFieldDateGroup && (((ExcelPivotTableFieldDateGroup)fld.Grouping).GroupBy) == groupBy)
					return fld;
			}
			return null;
		}

		/// <summary>
		/// Returns the numeric group field.
		/// </summary>
		/// <returns>The matching field or null if none is found.</returns>
		public ExcelPivotTableField GetNumericGroupField()
		{
			foreach (var fld in myList)
			{
				if (fld.Grouping is ExcelPivotTableFieldNumericGroup)
					return fld;
			}
			return null;
		}
		#endregion
	}

	/// <summary>
	/// Collection class for Row and column fields in a pivot table .
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

	/// <summary>
	/// Collection class for data fields in a Pivottable 
	/// </summary>
	public class ExcelPivotTableDataFieldCollection : ExcelPivotTableFieldCollectionBase<ExcelPivotTableDataField>
	{
		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="ExcelPivotTableDataFieldCollection"/>.
		/// </summary>
		/// <param name="table">The pivot table.</param>
		internal ExcelPivotTableDataFieldCollection(ExcelPivotTable table) :
			 base(table)
		{

		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Add a new data field.
		/// </summary>
		/// <param name="field">The field being added.</param>
		/// <returns>The new data field.</returns>
		public ExcelPivotTableDataField Add(ExcelPivotTableField field)
		{
			var dataFieldsNode = field.TopNode.SelectSingleNode("../../d:dataFields", field.NameSpaceManager);
			if (dataFieldsNode == null)
			{
				myTable.CreateNode("d:dataFields");
				dataFieldsNode = field.TopNode.SelectSingleNode("../../d:dataFields", field.NameSpaceManager);
			}

			XmlElement node = myTable.PivotTableXml.CreateElement("dataField", ExcelPackage.schemaMain);
			node.SetAttribute("fld", field.Index.ToString());
			dataFieldsNode.AppendChild(node);

			//XmlElement node = field.AppendField(dataFieldsNode, field.Index, "dataField", "fld");
			field.SetXmlNodeBool("@dataField", true, false);

			var dataField = new ExcelPivotTableDataField(field.NameSpaceManager, node, field);
			ValidateDupName(dataField);

			myList.Add(dataField);
			return dataField;
		}

		/// <summary>
		/// Remove a data field.
		/// </summary>
		/// <param name="dataField">The specified data field.</param>
		public void Remove(ExcelPivotTableDataField dataField)
		{
			if (dataField.Field.TopNode.SelectSingleNode(string.Format("../../d:dataFields/d:dataField[@fld={0}]", dataField.Index), dataField.NameSpaceManager) is XmlElement node)
				node.ParentNode.RemoveChild(node);
			myList.Remove(dataField);
		}

		/// <summary>
		/// Checks if a data field exists.
		/// </summary>
		/// <param name="name">The name of the data field.</param>
		/// <param name="datafield">The pivot table data field.</param>
		/// <returns>True if the data field exists.</returns>
		internal bool ExistsDfName(string name, ExcelPivotTableDataField datafield)
		{
			foreach (var df in myList)
			{
				if (((!string.IsNullOrEmpty(df.Name) && df.Name.Equals(name, StringComparison.InvariantCultureIgnoreCase) ||
					  (string.IsNullOrEmpty(df.Name) && df.Field.Name.Equals(name, StringComparison.InvariantCultureIgnoreCase)))) && datafield != df)
					return true;
			}
			return false;
		}
		#endregion

		#region Private Methods
		private void ValidateDupName(ExcelPivotTableDataField dataField)
		{
			if (this.ExistsDfName(dataField.Field.Name, null))
			{
				var index = 2;
				string name;
				do
				{
					name = dataField.Field.Name + "_" + index++.ToString();
				}
				while (this.ExistsDfName(name, null));
				dataField.Name = name;
			}
		}
		#endregion
	}
}