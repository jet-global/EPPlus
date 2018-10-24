using System;
using System.Collections.Generic;

namespace OfficeOpenXml.Table.PivotTable
{
	/// <summary>
	/// A row label or column header object used to calculate pivot table cell values.
	/// </summary>
	internal class PivotTableHeader
	{
		#region Properties
		/// <summary>
		/// Gets or sets the list of <see cref="ExcelPivotCacheRecords"/> indices.
		/// </summary>
		public List<Tuple<int, int>> CacheRecordIndices { get; set; }
		 
		/// <summary>
		/// Gets or sets the related <see cref="ExcelPivotTableField"/>.
		/// </summary>
		public ExcelPivotTableField PivotTableField { get; set; }
		
		/// <summary>
		/// Gets or sets the field index of the <see cref="ExcelPivotTableDataField"/>.
		/// </summary>
		public int DataFieldIndex { get; set; }

		/// <summary>
		/// Gets or sets the flag indicating if this is a grand total object.
		/// </summary>
		public bool IsGrandTotal { get; set; }

		/// <summary>
		/// Gets or sets the flag indicating if this is a row label.
		/// </summary>
		public bool IsRowHeader { get; set; }

		/// <summary>
		/// Gets or sets the value if there is a <see cref="RowColumnItem"/>  with a non-null itemType, typically 'default'.
		/// </summary>
		public string SumType { get; set; }
		#endregion

		#region Constructors
		/// <summary>
		/// Creates a new <see cref="PivotTableHeader"/> object.
		/// </summary>
		/// <param name="recordIndices">The list of cacheRecord indices.</param>
		/// <param name="field">The pivot table field.</param>
		/// <param name="dataFieldIndex">The index of the data field.</param>
		/// <param name="isGrandTotal">A value indicating if it is a grand total.</param>
		/// <param name="isRowHeader">A value indicating if it is a row header.</param>
		/// <param name="sumType">The itemType value of the <see cref="RowColumnItem"/>.</param>
		public PivotTableHeader(List<Tuple<int, int>> recordIndices, ExcelPivotTableField field, int dataFieldIndex, bool isGrandTotal, bool isRowHeader, string sumType = null)
		{
			this.CacheRecordIndices = recordIndices;
			this.PivotTableField = field;
			this.DataFieldIndex = dataFieldIndex;
			this.IsGrandTotal = isGrandTotal;
			this.IsRowHeader = isRowHeader;
			this.SumType = sumType;
		}
		#endregion
	}
}