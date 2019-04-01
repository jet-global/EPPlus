using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.Table.PivotTable.DataCalculation.ShowDataAsCalculation
{
	/// <summary>
	/// Calculates the <see cref="ShowDataAs.Percent"/> value in a pivot table.
	/// </summary>
	internal class PercentOfCalculator : ShowDataAsCalculatorBase
	{
		#region Constructors
		/// <summary>
		/// Constructs the calculator.
		/// </summary>
		/// <param name="pivotTable">The pivot table to calculate against.</param>
		/// <param name="dataFieldCollectionIndex">The index of the data field that the calculator is calculating.</param>
		/// <param name="totalsCalculator">A <see cref="TotalsFunctionHelper"/> to calculate values with.</param>
		public PercentOfCalculator(ExcelPivotTable pivotTable, int dataFieldCollectionIndex, TotalsFunctionHelper totalsCalculator) 
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
			var dataField = base.PivotTable.DataFields[base.DataFieldCollectionIndex];
			// TODO: Deal with "(next)" and "(previous)" options for this setting. See task #11840
			// "(next)" is stored as "1048829" and "(previous)" is "1048828".
			if (dataField.BaseItem == 1048829)
				throw new InvalidOperationException(@"'(next)' is not supported for the 'Show Data as Percent of' setting.");
			else if (dataField.BaseItem == 1048828)
				throw new InvalidOperationException(@"'(previous)' is not supported for the 'Show Data as Percent of' setting.");

			var baseFieldItemTuple = new Tuple<int, int>(dataField.BaseField, dataField.BaseItem);
			if (cellBackingData == null)
				return cellBackingData?.Result;
			else if (rowHeader.CacheRecordIndices.Any(t => t.Equals(baseFieldItemTuple)) || columnHeader.CacheRecordIndices.Any(t => t.Equals(baseFieldItemTuple)))
			{
				if (cellBackingData?.Result != null)
					return 1;  // At a row/column that contains the comparison field item which makes this 100%.
				else
					return cellBackingData?.Result;
			}
			else
			{
				// Try to find a value that matches either the current row or column header structure.
				// If a value is found, the percentage can be calculated. Otherwise, the appropriate error is written out.
				if (this.TryFindMatchingHeaderIndex(rowHeader, baseFieldItemTuple, base.PivotTable.RowHeaders, out int headerIndex))
				{
					var baseValue = backingDatas[headerIndex, dataColumn]?.Result;
					return this.GetShowDataAsPercentOfValue(baseValue, cellBackingData?.Result);
				}
				else if (this.TryFindMatchingHeaderIndex(columnHeader, baseFieldItemTuple, base.PivotTable.ColumnHeaders, out headerIndex))
				{
					var baseValue = backingDatas[dataRow, headerIndex]?.Result;
					return this.GetShowDataAsPercentOfValue(baseValue, cellBackingData?.Result);
				}
				else
				{
					if (!base.PivotTable.RowFields.Any(f => f.Index == dataField.BaseField)
						&& !base.PivotTable.ColumnFields.Any(f => f.Index == dataField.BaseField))
					{
						// If the dataField.BaseField is not a row or column field, all values #N/A!
						return ExcelErrorValue.Create(eErrorType.NA);
					}
					else if (!rowHeader.CacheRecordIndices.Any(i => i.Item1 == dataField.BaseField)
						&& !columnHeader.CacheRecordIndices.Any(i => i.Item1 == dataField.BaseField))
					{
						// Subtotals only get a value if they are at the same depth or below the show data as field.
						return null;
					}
					else
						return ExcelErrorValue.Create(eErrorType.NA);
				}
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
			var dataField = base.PivotTable.DataFields[base.DataFieldCollectionIndex];
			// TODO: Deal with "(next)" and "(previous)" options for this setting. See task #11840
			// "(next)" is stored as "1048829" and "(previous)" is "1048828".
			if (dataField.BaseItem == 1048829)
				throw new InvalidOperationException(@"'(next)' is not supported for the 'Show Data as Percent of' setting.");
			else if (dataField.BaseItem == 1048828)
				throw new InvalidOperationException(@"'(previous)' is not supported for the 'Show Data as Percent of' setting.");
			var grandTotalBackingData = grandTotalsBackingDatas[index];
			var baseFieldItemTuple = new Tuple<int, int>(dataField.BaseField, dataField.BaseItem);
			var headers = isRowTotal ? base.PivotTable.ColumnHeaders : base.PivotTable.RowHeaders;
			var header = headers[grandTotalBackingData.MajorAxisIndex];
			if (header.CacheRecordIndices.Any(t => t.Equals(baseFieldItemTuple)))
				return 1;  // At a row/column that contains the comparison field item which makes this 100%.
			else
			{
				// Try to find a value that matches either the current row or column header structure.
				// If a value is found, the percentage can be calculated. Otherwise, the appropriate error is written out.
				if (this.TryFindMatchingHeaderIndex(header, baseFieldItemTuple, headers, out int headerIndex))
				{
					// Get the correct index into grand totals backing data which is a 1d array 
					// representing [datafields.Count] number of rows/columns.
					var denominatorHeader = grandTotalsBackingDatas
						.Where(d => d.MajorAxisIndex == headerIndex)
						.ElementAt(grandTotalBackingData.DataFieldCollectionIndex);
					var baseValue = denominatorHeader?.Result;
					return this.GetShowDataAsPercentOfValue(baseValue, grandTotalBackingData?.Result);
				}
				else
				{
					if (!header.CacheRecordIndices.Any(x => x.Item1 == dataField.BaseField)
						&& (this.PivotTable.RowFields.Any(f => f.Index == dataField.BaseField)
						|| this.PivotTable.ColumnFields.Any(f => f.Index == dataField.BaseField)))
						return null;

					// If the dataField.BaseField is not a row or column field, all values #N/A!
					return ExcelErrorValue.Create(eErrorType.NA);
				}
			}
		}

		/// <summary>
		/// Calculates a grand grand total value for a pivot table cell.
		/// </summary>
		/// <param name="backingData">The backing data for the grand total cell to calculate.</param>
		/// <returns>An object value for the cell.</returns>
		public override object CalculateGrandGrandTotalValue(PivotCellBackingData backingData)
		{
			var dataField = base.PivotTable.DataFields[base.DataFieldCollectionIndex];
			if (!base.PivotTable.RowFields.Any(f => f.Index == dataField.BaseField)
				&& !base.PivotTable.ColumnFields.Any(f => f.Index == dataField.BaseField))
			{
				// If the dataField.BaseField is not a row or column field, all values #N/A!
				return ExcelErrorValue.Create(eErrorType.NA);
			}
			// The "% Of" option doesn't write in values for grand grand totals.
			return null;
		}
		#endregion

		#region Private Methods
		private object GetShowDataAsPercentOfValue(object baseValue, object value)
		{
			if (baseValue == null && value == null)
				return ExcelErrorValue.Create(eErrorType.Null);
			else if (baseValue == null)
				return null;
			else if (value == null)
				return ExcelErrorValue.Create(eErrorType.Null);
			else
				return base.CalculatePercentage(value, baseValue);
		}

		private bool TryFindMatchingHeaderIndex(PivotTableHeader header, Tuple<int, int> baseFieldItem, List<PivotTableHeader> headers, out int headerIndex)
		{
			headerIndex = -1;
			var index = header.CacheRecordIndices.FindIndex(i => i.Item1 == baseFieldItem.Item1);
			if (index >= 0)
			{
				// The value that this will be compared against is in the cell that matches this cell's 
				// row and colum indices other than the base field/item indices.
				var indicesToMatch = header.CacheRecordIndices.ToList();
				indicesToMatch[index] = baseFieldItem;
				headerIndex = headers.FindIndex(h => this.AreEquivalent(h.CacheRecordIndices, indicesToMatch));
				return headerIndex != -1;
			}
			return false;
		}

		private bool AreEquivalent(List<Tuple<int, int>> first, List<Tuple<int, int>> second)
		{
			if (first?.Count != second?.Count)
				return false;
			for (int j = 0; j < first.Count; j++)
			{
				if (!first[j].Equals(second[j]))
					return false;
			}
			return true;
		}
		#endregion
	}
}
