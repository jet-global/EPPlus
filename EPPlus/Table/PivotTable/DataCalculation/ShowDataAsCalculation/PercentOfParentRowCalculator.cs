using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.Table.PivotTable.DataCalculation.ShowDataAsCalculation
{
	/// <summary>
	/// Calculates the <see cref="ShowDataAs.PercentOfParentRow"/> value in a pivot table.
	/// </summary>
	internal class PercentOfParentRowCalculator : ShowDataAsCalculatorBase
	{
		#region Constructors
		/// <summary>
		/// Constructs the calculator.
		/// </summary>
		/// <param name="pivotTable">The pivot table to calculate against.</param>
		/// <param name="dataFieldCollectionIndex">The index of the data field that the calculator is calculating.</param>
		public PercentOfParentRowCalculator(ExcelPivotTable pivotTable, int dataFieldCollectionIndex) : base(pivotTable, dataFieldCollectionIndex) { }
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

			object baseValue = null;
			if (this.TryFindParent(dataRow, out int parentIndex))
				baseValue = backingDatas[parentIndex, dataColumn]?.Result;
			else if (rowHeader.IsDataField)
				return null;  // Data field root nodes don't get values.
			else
			{
				// At a root node, go to the grand grand total for the base value.
				baseValue = rowGrandTotalsValuesLists
					.Where(d => d.MajorAxisIndex == dataColumn)
					.ElementAt(rowHeader.DataFieldCollectionIndex)
					?.Result;
			}

			if (cellBackingData?.Result == null)
			{
				// If both are null, write null.
				if (baseValue == null)
					return null;
				// If the parent has a value, write out 0.
				return 0;
			}
			return (double)cellBackingData.Result / (double)baseValue;
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
			if (isRowTotal)
				return 1;
			object baseValue = null;
			var grandTotalBackingData = grandTotalsBackingDatas[index];
			if (this.TryFindParent(grandTotalBackingData.MajorAxisIndex, out int parentIndex))
			{
				baseValue = grandTotalsBackingDatas
					.First(v => v.MajorAxisIndex == parentIndex && v.DataFieldCollectionIndex == grandTotalBackingData.DataFieldCollectionIndex)
					.Result;
			}
			else if (this.PivotTable.RowHeaders[grandTotalBackingData.MajorAxisIndex].IsDataField)
				return null;  // Data field root nodes don't get values.
			else
			{
				// If a value was not found, the grand total value is the base value.
				baseValue = (double)columnGrandGrandTotalValues[grandTotalBackingData.DataFieldCollectionIndex].Result;
			}
			return (double)grandTotalBackingData.Result / (double)baseValue;
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
			var header = base.PivotTable.RowHeaders[startIndex];
			// Walk backwards up the row headers until we find a parent.
			for (int i = startIndex - 1; i >= 0; i--)
			{
				var previousRowHeader = base.PivotTable.RowHeaders[i];
				if (previousRowHeader.CacheRecordIndices.Count < header.CacheRecordIndices.Count && previousRowHeader.IsDataField == false)
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
