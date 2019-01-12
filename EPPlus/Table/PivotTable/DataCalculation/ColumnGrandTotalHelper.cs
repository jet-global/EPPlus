using System.Collections.Generic;

namespace OfficeOpenXml.Table.PivotTable.DataCalculation
{
	/// <summary>
	/// Calculate column grand totals (grand totals on the right side of the pivot table).
	/// </summary>
	internal class ColumnGrandTotalHelper : GrandTotalHelperBase
	{
		#region Constructors
		/// <summary>
		/// Create a new <see cref="ColumnGrandTotalHelper"/> object.
		/// </summary>
		/// <param name="pivotTable">The <see cref="ExcelPivotTable"/>.</param>
		/// <param name="backingData">The data backing the pivot table.</param>
		internal ColumnGrandTotalHelper(ExcelPivotTable pivotTable, List<object>[,] backingData) : base(pivotTable, backingData)
		{
			this.MajorHeaderCollection = this.PivotTable.RowHeaders;
			this.MinorHeaderCollection = this.PivotTable.ColumnHeaders;
		}
		#endregion

		#region GrandTotalHelperBase Overrides
		/// <summary>
		/// Adds matching values to the <paramref name="grandTotalValueLists"/> and <paramref name="grandGrandTotalValueLists"/>.
		/// </summary>
		/// <param name="majorHeader">The major axis header.</param>
		/// <param name="majorIndex">The current major axis index.</param>
		/// <param name="minorIndex">The current minor axis index.</param>
		/// <param name="dataFieldCollectionIndex">The index of the data field in the data field collection.</param>
		/// <param name="grandTotalValueLists">The list of values used to calculate grand totals for a row or column.</param>
		/// <param name="grandGrandTotalValueLists">The list of values used to calcluate the grand-grand total values.</param>
		/// <returns>The index of the data field in the data field collection.</returns>
		protected override int AddMatchingValues(
			PivotTableHeader majorHeader,
			int majorIndex,
			int minorIndex,
			int dataFieldCollectionIndex,
			List<object>[] grandTotalValueLists,
			List<object>[] grandGrandTotalValueLists)
		{
			if (this.BackingData[majorIndex, minorIndex] == null)
				return dataFieldCollectionIndex;
			var minorHeader = this.MinorHeaderCollection[minorIndex];
			dataFieldCollectionIndex = this.PivotTable.HasRowDataFields ? majorHeader.DataFieldCollectionIndex : minorHeader.DataFieldCollectionIndex;
			if (minorHeader.IsLeafNode)
			{
				if (grandTotalValueLists[dataFieldCollectionIndex] == null)
					grandTotalValueLists[dataFieldCollectionIndex] = new List<object>();
				grandTotalValueLists[dataFieldCollectionIndex].AddRange(base.BackingData[majorIndex, minorIndex]);
				// Only add row header leaf node values for grand-grand totals.
				if (majorHeader.IsLeafNode)
				{
					if (grandGrandTotalValueLists[dataFieldCollectionIndex] == null)
						grandGrandTotalValueLists[dataFieldCollectionIndex] = new List<object>();
					grandGrandTotalValueLists[dataFieldCollectionIndex].AddRange(base.BackingData[majorIndex, minorIndex]);
				}
			}
			return dataFieldCollectionIndex;
		}

		/// <summary>
		/// Calculates and writes the grand total values to the worksheet.
		/// </summary>
		/// <param name="majorIndex">The current major axis index.</param>
		/// <param name="totalsCalculator">The grand totals calculation helper class.</param>
		/// <param name="grandTotalValueLists">The values used to calculate grand totals.</param>
		protected override void WriteGrandTotal(
			int majorIndex, 
			TotalsFunctionHelper totalsCalculator, 
			List<object>[] grandTotalValueLists)
		{
			var row = this.PivotTable.Address.Start.Row + this.PivotTable.FirstDataRow + majorIndex;
			var column = this.PivotTable.Address.End.Column;
			if (this.PivotTable.HasColumnDataFields)
				column -= this.PivotTable.DataFields.Count - 1;
			for (int i = 0; i < grandTotalValueLists.Length; i++)
			{
				if (grandTotalValueLists[i] != null)
					base.WriteCellTotal(row, column++, this.PivotTable.DataFields[i], grandTotalValueLists[i], totalsCalculator);
			}
		}

		/// <summary>
		/// Gets the start cell index for the grand-grand total values.
		/// </summary>
		/// <returns>The start cell index for the grand-grand total values.</returns>
		protected override int GetStartIndex()
		{
			return this.PivotTable.Address.End.Column - this.PivotTable.DataFields.Count + 1;
		}

		/// <summary>
		/// Writes the grand total for the specified <paramref name="values"/> in the cell at the specified <paramref name="index"/>.
		/// </summary>
		/// <param name="index">The major index of the cell to write the total to.</param>
		/// <param name="dataField">The data field to use the number format of.</param>
		/// <param name="values">The values to use to calculate the total.</param>
		/// <param name="totalsFunctionHelper">The totals calculation helper class.</param>
		protected override void WriteCellTotal(int index, ExcelPivotTableDataField dataField, List<object> values, TotalsFunctionHelper totalsFunctionHelper)
		{
			string cacheFieldFormula = this.PivotTable.CacheDefinition.CacheFields[dataField.Index].Formula;
			if (string.IsNullOrEmpty(cacheFieldFormula))
				this.WriteCellTotal(this.PivotTable.Address.End.Row, index, dataField, values, totalsFunctionHelper);
			else
			{
				// TODO: Calculating grand totals for calculated fields requires all of the values
				// that were used to calculate the subtotals. 
				// Instead of passing in the list of values, consider adding the full backing pivot table data
				// to the totals function helper and using that data to calculate the calculated field.
				throw new System.NotImplementedException();
			}

		}
		#endregion
	}
}