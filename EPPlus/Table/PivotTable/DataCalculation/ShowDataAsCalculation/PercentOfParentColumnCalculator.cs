using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.Table.PivotTable.DataCalculation.ShowDataAsCalculation
{
	/// <summary>
	/// Calculates the <see cref="ShowDataAs.PercentOfCol"/> value in a pivot table.
	/// </summary>
	internal class PercentOfParentColumnCalculator : ShowDataAsCalculatorBase
	{
		#region Constructors
		/// <summary>
		/// Constructs the calculator.
		/// </summary>
		/// <param name="pivotTable">The pivot table to calculate against.</param>
		/// <param name="dataFieldCollectionIndex">The index of the data field that the calculator is calculating.</param>
		public PercentOfParentColumnCalculator(ExcelPivotTable pivotTable, int dataFieldCollectionIndex) : base(pivotTable, dataFieldCollectionIndex) { }
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
		/// <param name="totalsCalculator">A <see cref="TotalsFunctionHelper"/> to calculate values with.</param>
		/// <returns>An object value for the cell.</returns>
		public override object CalculateBodyValue(
			int dataRow, int dataColumn,
			PivotCellBackingData[,] backingDatas,
			PivotCellBackingData[] grandGrandTotalValues,
			List<PivotCellBackingData> rowGrandTotalsValuesLists,
			List<PivotCellBackingData> columnGrandTotalsValuesLists,
			TotalsFunctionHelper totalsCalculator)
		{
			var rowHeader = base.PivotTable.RowHeaders[dataRow];
			var columnHeader = base.PivotTable.ColumnHeaders[dataColumn];
			var cellBackingData = backingDatas[dataRow, dataColumn];

			if (columnHeader.IsDataField)
				return null;
			else
			{
				// Because columns don't always have a cell where a subtotal would go, we need to calculate
				// subtotals from scratch here.
				PivotCellBackingData subtotalBackingData = null;
				if (cellBackingData.IsCalculatedCell)
					subtotalBackingData = new PivotCellBackingData(new Dictionary<string, List<object>>(), cellBackingData.Formula);
				else
					subtotalBackingData = new PivotCellBackingData(new List<object>());

				// Find all of the cells with the same parent and merge their backing datas.
				for (int i = 0; i < base.PivotTable.ColumnHeaders.Count; i++)
				{
					var possibleSibling = this.PivotTable.ColumnHeaders[i];
					bool isSibling = this.IsSibling(columnHeader.CacheRecordIndices, possibleSibling.CacheRecordIndices);
					if (isSibling)
					{
						var backingData = backingDatas[dataRow, i];
						if (backingData != null)
							subtotalBackingData.Merge(backingData);
					}
				}

				object baseValue = null;
				var dataField = this.PivotTable.DataFields[base.DataFieldCollectionIndex];
				baseValue = totalsCalculator.CalculateCellTotal(dataField, subtotalBackingData, rowHeader.TotalType, columnHeader.TotalType);

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
			object baseValue = null;
			var grandTotalBackingData = grandTotalsBackingDatas[index];
			if (this.TryFindParent(grandTotalBackingData.MajorAxisIndex, out int parentIndex))
			{
				baseValue = grandTotalsBackingDatas
					.First(v => v.MajorAxisIndex == parentIndex && v.DataFieldCollectionIndex == grandTotalBackingData.DataFieldCollectionIndex)
					.Result;
			}
			else if (this.PivotTable.ColumnHeaders[grandTotalBackingData.MajorAxisIndex].IsDataField)
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
		private bool TryFindBodyParent(int startIndex, out int index)
		{
			index = 0;
			var header = base.PivotTable.ColumnHeaders[startIndex];
			// Walk down the headers until we find a parent.
			for (int i = startIndex + 1; i < base.PivotTable.ColumnHeaders.Count; i++)
			{
				var previousHeader = base.PivotTable.ColumnHeaders[i];
				if (previousHeader.CacheRecordIndices?.Count < header.CacheRecordIndices.Count && previousHeader.IsDataField == false)
				{
					index = i;
					return true;
				}
			}
			index = -1;
			return false;
		}

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

		private bool IsSibling(List<Tuple<int, int>> indices, List<Tuple<int, int>> possibleSiblingIndices)
		{
			if (indices == null || possibleSiblingIndices == null || indices.Count != possibleSiblingIndices.Count)
				return false;
			for (int i = 0; i < indices.Count - 1; i++)
			{
				if (indices[i] != possibleSiblingIndices[i])
					return false;
			}
			return true;
		}
		#endregion
	}
}
