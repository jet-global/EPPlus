using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.Table.PivotTable.DataCalculation.ShowDataAsCalculation
{
	/// <summary>
	/// Calculates the <see cref="ShowDataAs.PercentOfCol"/> value in a pivot table.
	/// </summary>
	internal class PercentOfParentColumnCalculator : PercentOfParentCalculatorBase
	{
		#region Constructors
		/// <summary>
		/// Constructs the calculator.
		/// </summary>
		/// <param name="pivotTable">The pivot table to calculate against.</param>
		/// <param name="dataFieldCollectionIndex">The index of the data field that the calculator is calculating.</param>
		/// <param name="totalsCalculator">A <see cref="TotalsFunctionHelper"/> to calculate values with.</param>
		public PercentOfParentColumnCalculator(ExcelPivotTable pivotTable, int dataFieldCollectionIndex, TotalsFunctionHelper totalsCalculator) 
			: base(pivotTable, dataFieldCollectionIndex, totalsCalculator) { }
		#endregion

		#region ShowDataAsCalculatorBase Overrides
		/// <summary>
		/// Calculates a body value in a pivot table cell.
		/// </summary>
		/// <param name="dataRow">The row in the backing body data.</param>
		/// <param name="dataColumn">The column in the backing body data.</param>
		/// <param name="backingDatas">The backing data for the pivot table body.</param>
		/// <param name="grandGrandTotalValues">The backing data for the pivot table grand grand totals.</param>
		/// <param name="rowGrandTotalsValuesLists">The backing data for the pivot table row grand totals.</param>
		/// <param name="columnGrandTotalsValuesLists">The backing data for the pivot table column grand totals.</param>
		/// <returns>An object value for the cell.</returns>
		public override object CalculateBodyValue(
		int dataRow, int dataColumn,
		PivotCellBackingData[,] backingDatas,
		PivotCellBackingData[] grandGrandTotalValues,
		List<PivotCellBackingData> rowGrandTotalsValuesLists,
		List<PivotCellBackingData> columnGrandTotalsValuesLists)
		{
			var rowHeader = base.PivotTable.RowHeaders[dataRow];
			var columnHeader = base.PivotTable.ColumnHeaders[dataColumn];
			var cellBackingData = backingDatas[dataRow, dataColumn];

			if (columnHeader.IsDataField)
				return null;
			else
			{
				var parentHeaderIndices = new List<Tuple<int, int>>();
				if (columnHeader.CacheRecordIndices.Count > 1)
					parentHeaderIndices = columnHeader.CacheRecordIndices.Take(columnHeader.CacheRecordIndices.Count - 1).ToList();
				else if (columnHeader.IsDataField)
					return null;  // Data field root nodes don't get values.
				return base.CalculateBodyValue(false, dataRow, dataColumn, parentHeaderIndices, backingDatas);
			}
		}

		/// <summary>
		/// Calculates the grand total value in a pivot table cell.
		/// </summary>
		/// <param name="index">The index into the backing data.</param>
		/// <param name="grandTotalsBackingDatas">The backing data for grand totals.</param>
		/// <param name="columnGrandGrandTotalValues">The backing data for the column grand grand totals.</param>
		/// <param name="isRowTotal">A value indicating whether or not this calculation is for row totals.</param>
		/// <returns>An object value for the cell.</returns>
		public override object CalculateGrandTotalValue(
			int index,
			List<PivotCellBackingData> grandTotalsBackingDatas,
			PivotCellBackingData[] columnGrandGrandTotalValues,
			bool isRowTotal)
		{
			if (!isRowTotal)
				return 1;
			var dataField = base.PivotTable.DataFields[base.DataFieldCollectionIndex];
			var cellBackingData = grandTotalsBackingDatas[index];
			return base.CalculateGrandTotalValue(base.PivotTable.ColumnHeaders, cellBackingData, isRowTotal);
		}

		/// <summary>
		/// Calculates a grand grand total value for a pivot table cell.
		/// </summary>
		/// <param name="backingData">The backing data for the grand total cell to calculate.</param>
		/// <returns>An object value for the cell.</returns>
		public override object CalculateGrandGrandTotalValue(PivotCellBackingData backingData) => 1;
		#endregion

		#region Private Methods
		private bool TryFindParent(int startIndex, out int index)
		{
			index = 0;
			var header = base.PivotTable.ColumnHeaders[startIndex];
			// Walk backwards up the headers until we find a parent.
			for (int i = startIndex - 1; i >= 0; i--)
			{
				var previousHeader = base.PivotTable.ColumnHeaders[i];
				if (previousHeader.CacheRecordIndices.Count < header.CacheRecordIndices.Count && previousHeader.IsDataField == false)
				{
					index = i;
					return true;
				}
			}
			index = -1;
			return false;
		}
		#endregion
	}
}
