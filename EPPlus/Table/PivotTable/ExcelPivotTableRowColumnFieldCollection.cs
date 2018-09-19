using System;

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