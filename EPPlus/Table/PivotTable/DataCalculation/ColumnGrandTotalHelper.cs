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
		/// <param name="totalsCalculator">The calculation helper.</param>
		internal ColumnGrandTotalHelper(ExcelPivotTable pivotTable, PivotCellBackingData[,] backingData, TotalsFunctionHelper totalsCalculator) 
			: base(pivotTable, backingData, totalsCalculator)
		{
			this.MajorHeaderCollection = this.PivotTable.RowHeaders;
			this.MinorHeaderCollection = this.PivotTable.ColumnHeaders;
		}
		#endregion

		#region GrandTotalHelperBase Overrides
		/// <summary>
		/// Adds matching values to the <paramref name="grandTotalValueList"/> and <paramref name="grandGrandTotalValueList"/>.
		/// </summary>
		/// <param name="majorHeader">The major axis header.</param>
		/// <param name="majorIndex">The current major axis index.</param>
		/// <param name="minorIndex">The current minor axis index.</param>
		/// <param name="dataFieldCollectionIndex">The index of the data field in the data field collection.</param>
		/// <param name="grandTotalValueList">The list of values used to calculate grand totals for a row or column.</param>
		/// <param name="grandGrandTotalValueList">The list of values used to calcluate the grand-grand total values.</param>
		/// <returns>The index of the data field in the data field collection.</returns>
		protected override int AddMatchingValues(
			PivotTableHeader majorHeader,
			int majorIndex,
			int minorIndex,
			int dataFieldCollectionIndex,
			PivotCellBackingData[] grandTotalValueList,
			PivotCellBackingData[] grandGrandTotalValueList)
		{
			if (this.BackingData[majorIndex, minorIndex] == null)
				return dataFieldCollectionIndex;
			var minorHeader = this.MinorHeaderCollection[minorIndex];
			dataFieldCollectionIndex = this.PivotTable.HasRowDataFields ? majorHeader.DataFieldCollectionIndex : minorHeader.DataFieldCollectionIndex;
			if (minorHeader.IsLeafNode)
			{
				base.AddGrandTotalsBackingData(majorIndex, minorIndex, dataFieldCollectionIndex, grandTotalValueList);
				// Only add row header leaf node values for grand-grand totals.
				if (majorHeader.IsLeafNode)
					base.AddGrandTotalsBackingData(majorIndex, minorIndex, dataFieldCollectionIndex, grandGrandTotalValueList);
			}
			return dataFieldCollectionIndex;
		}

		/// <summary>
		/// Calculates and writes the grand total values to the worksheet.
		/// </summary>
		/// <param name="majorIndex">The current major axis index.</param>
		/// <param name="grandTotalValueLists">The values used to calculate grand totals.</param>
		protected override void WriteGrandTotal(
			int majorIndex, 
			PivotCellBackingData[] grandTotalValueLists)
		{
			var row = this.PivotTable.Address.Start.Row + this.PivotTable.FirstDataRow + majorIndex;
			var column = this.PivotTable.Address.End.Column;
			if (this.PivotTable.HasColumnDataFields)
				column -= this.PivotTable.DataFields.Count - 1;
			for (int i = 0; i < grandTotalValueLists.Length; i++)
			{
				if (grandTotalValueLists[i] != null)
				{
					var cell = this.PivotTable.Worksheet.Cells[row, column++];
					var dataField = this.PivotTable.DataFields[i];
					var cacheField = this.PivotTable.CacheDefinition.CacheFields[dataField.Index];
					base.TotalsCalculator.WriteCellTotal(cell, dataField, grandTotalValueLists[i], this.PivotTable.Worksheet.Workbook.Styles);
				}
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
		/// Writes the grand total for the specified <paramref name="backingData"/> in the cell at the specified <paramref name="index"/>.
		/// </summary>
		/// <param name="index">The major index of the cell to write the total to.</param>
		/// <param name="dataField">The data field to use the number format of.</param>
		/// <param name="backingData">The values to use to calculate the total.</param>
		protected override void WriteCellTotal(int index, ExcelPivotTableDataField dataField, PivotCellBackingData backingData)
		{
			var cell = this.PivotTable.Worksheet.Cells[this.PivotTable.Address.End.Row, index];
			var styles = this.PivotTable.Worksheet.Workbook.Styles;
			base.TotalsCalculator.WriteCellTotal(cell, dataField, backingData, styles);
		}
		#endregion
	}
}