using System.Collections.Generic;

namespace OfficeOpenXml.Table.PivotTable.DataCalculation
{
	/// <summary>
	/// Calculate column grand totals.
	/// </summary>
	internal class ColumnGrandTotalHelper : GrandTotalHelperBase
	{
		#region Constructors
		/// <summary>
		/// Create a new <see cref="ColumnGrandTotalHelper"/> object.
		/// </summary>
		/// <param name="pivotTable">The <see cref="ExcelPivotTable"/>.</param>
		/// <param name="backingData">The data backing the pivot table.</param>
		public ColumnGrandTotalHelper(ExcelPivotTable pivotTable, List<object>[,] backingData) : base(pivotTable, backingData)
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
			List<object>[] grandTotalValueLists, 
			List<object>[] grandGrandTotalValueLists)
		{
			if (this.BackingData[minorIndex, majorIndex] == null)
				return dataFieldCollectionIndex;
			var minorHeader = this.MinorHeaderCollection[minorIndex];
			dataFieldCollectionIndex = this.PivotTable.HasRowDataFields ? minorHeader.DataFieldCollectionIndex : majorHeader.DataFieldCollectionIndex;
			if (minorHeader.IsLeafNode)
			{
				if (grandTotalValueLists[dataFieldCollectionIndex] == null)
					grandTotalValueLists[dataFieldCollectionIndex] = new List<object>();
				grandTotalValueLists[dataFieldCollectionIndex].AddRange(this.BackingData[minorIndex, majorIndex]);
			}
			return dataFieldCollectionIndex;
		}

		/// <summary>
		/// Calculates and writes the grand total values to the worksheet.
		/// </summary>
		/// <param name="dataFieldCollectionIndex">The index of the data field in the data field collection.</param>
		/// <param name="majorIndex">The current major axis index.</param>
		/// <param name="totalsCalculator">The grand totals calculation helper class.</param>
		/// <param name="grandTotalValueLists">The values used to calculate grand totals.</param>
		protected override void WriteGrandTotal(
			int dataFieldCollectionIndex,
			int majorIndex,
			TotalsFunctionHelper totalsCalculator,
			List<object>[] grandTotalValueLists)
		{
			var row = this.PivotTable.Address.End.Row;
			if (this.PivotTable.HasRowDataFields)
				row -= this.PivotTable.DataFields.Count - 1;
			var column = this.PivotTable.Address.Start.Column + this.PivotTable.FirstDataCol + majorIndex;
			foreach (var valuesList in grandTotalValueLists)
			{
				if (valuesList != null)
					this.WriteCellTotal(row++, column, this.PivotTable.DataFields[dataFieldCollectionIndex], valuesList, totalsCalculator);
			}
		}
		#endregion
	}
}