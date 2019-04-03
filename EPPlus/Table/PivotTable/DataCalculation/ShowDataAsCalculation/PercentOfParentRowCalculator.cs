using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.Table.PivotTable.DataCalculation.ShowDataAsCalculation
{
	/// <summary>
	/// Calculates the <see cref="ShowDataAs.PercentOfParentRow"/> value in a pivot table.
	/// </summary>
	internal class PercentOfParentRowCalculator : PercentOfParentCalculatorBase
	{
		#region Constructors
		/// <summary>
		/// Constructs the calculator.
		/// </summary>
		/// <param name="pivotTable">The pivot table to calculate against.</param>
		/// <param name="dataFieldCollectionIndex">The index of the data field that the calculator is calculating.</param>
		/// <param name="totalsCalculator">A <see cref="TotalsFunctionHelper"/> to calculate values with.</param>
		public PercentOfParentRowCalculator(ExcelPivotTable pivotTable, int dataFieldCollectionIndex, TotalsFunctionHelper totalsCalculator) 
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
			var parentHeaderIndices = new List<Tuple<int, int>>();
			if (rowHeader.CacheRecordIndices.Count > 1)
				parentHeaderIndices = rowHeader.CacheRecordIndices.Take(rowHeader.CacheRecordIndices.Count - 1).ToList();
			else if (rowHeader.IsDataField)
				return null;  // Data field root nodes don't get values.
			return base.CalculateBodyValue(true, dataRow, dataColumn, parentHeaderIndices, backingDatas);
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
			var dataField = base.PivotTable.DataFields[base.DataFieldCollectionIndex];
			var cellBackingData = grandTotalsBackingDatas[index];
			if (isRowTotal)
			{
				if ((double)cellBackingData.Result == 0)
					return null;
				else
					return 1;
			}
			return base.CalculateGrandTotalValue(base.PivotTable.RowHeaders, cellBackingData, isRowTotal);
		}

		/// <summary>
		/// Calculates a grand grand total value for a pivot table cell.
		/// </summary>
		/// <param name="backingData">The backing data for the grand total cell to calculate.</param>
		/// <returns>An object value for the cell.</returns>
		public override object CalculateGrandGrandTotalValue(PivotCellBackingData backingData)
		{
			if ((double)backingData.Result == 0)
				return null;
			else
				return 1;
		}
		#endregion
	}
}
