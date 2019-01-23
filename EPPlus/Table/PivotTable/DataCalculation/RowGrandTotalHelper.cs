namespace OfficeOpenXml.Table.PivotTable.DataCalculation
{
	/// <summary>
	/// Calculate row grand totals (grand totals at the bottom of a pivot table).
	/// </summary>
	internal class RowGrandTotalHelper : GrandTotalHelperBase
	{
		#region Constructors
		/// <summary>
		/// Create a new <see cref="RowGrandTotalHelper"/> object.
		/// </summary>
		/// <param name="pivotTable">The <see cref="ExcelPivotTable"/>.</param>
		/// <param name="backingData">The data backing the pivot table.</param>
		/// <param name="totalsCalculator">The calculation helper.</param>
		internal RowGrandTotalHelper(ExcelPivotTable pivotTable, PivotCellBackingData[,] backingData, TotalsFunctionHelper totalsCalculator) 
			: base(pivotTable, backingData, totalsCalculator)
		{
			this.MajorHeaderCollection = this.PivotTable.ColumnHeaders;
			this.MinorHeaderCollection = this.PivotTable.RowHeaders;
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
			PivotCellBackingData[] grandTotalValueLists, 
			PivotCellBackingData[] grandGrandTotalValueLists)
		{
			if (this.BackingData[minorIndex, majorIndex] == null)
				return dataFieldCollectionIndex;
			var minorHeader = this.MinorHeaderCollection[minorIndex];
			dataFieldCollectionIndex = this.PivotTable.HasRowDataFields ? minorHeader.DataFieldCollectionIndex : majorHeader.DataFieldCollectionIndex;
			if (minorHeader.IsLeafNode)
				base.AddGrandTotalsBackingData(minorIndex, majorIndex, dataFieldCollectionIndex, grandTotalValueLists);
			return dataFieldCollectionIndex;
		}

		/// <summary>
		/// Calculates and writes the grand total values to the worksheet.
		/// </summary>
		/// <param name="majorIndex">The current major axis index.</param>
		/// <param name="grandTotalValueLists">The values used to calculate grand totals.</param>
		protected override void WriteGrandTotal(int majorIndex, PivotCellBackingData[] grandTotalValueLists)
		{
			var row = this.PivotTable.Address.End.Row;
			if (this.PivotTable.HasRowDataFields)
				row -= this.PivotTable.DataFields.Count - 1;
			var column = this.PivotTable.Address.Start.Column + this.PivotTable.FirstDataCol + majorIndex;
			for (int i = 0; i < grandTotalValueLists.Length; i++)
			{
				if (grandTotalValueLists[i] != null)
				{
					var cell = this.PivotTable.Worksheet.Cells[row++, column];
					var dataField = this.PivotTable.DataFields[i];
					var cacheField = this.PivotTable.CacheDefinition.CacheFields[dataField.Index];
					var styles = this.PivotTable.Worksheet.Workbook.Styles;
					base.TotalsCalculator.WriteCellTotal(cell, dataField, grandTotalValueLists[i], styles);
				}
			}
		}

		/// <summary>
		/// Gets the start cell index for the grand-grand total values.
		/// </summary>
		/// <returns>The start cell index for the grand-grand total values.</returns>
		protected override int GetStartIndex()
		{
			return this.PivotTable.Address.End.Row - this.PivotTable.DataFields.Count + 1;
		}

		/// <summary>
		/// Writes the grand total for the specified <paramref name="backingData"/> in the cell at the specified <paramref name="index"/>.
		/// </summary>
		/// <param name="index">The major index of the cell to write the total to.</param>
		/// <param name="dataField">The data field to use the number format of.</param>
		/// <param name="backingData">The data to use to calculate the total.</param>
		protected override void WriteCellTotal(int index, ExcelPivotTableDataField dataField, PivotCellBackingData backingData)
		{
			var cell = this.PivotTable.Worksheet.Cells[index, this.PivotTable.Address.End.Column];
			var cacheField = this.PivotTable.CacheDefinition.CacheFields[dataField.Index];
			var styles = this.PivotTable.Worksheet.Workbook.Styles;
			base.TotalsCalculator.WriteCellTotal(cell, dataField, backingData, styles);
		}
		#endregion
	}
}