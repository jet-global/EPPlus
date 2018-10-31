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
		public ColumnGrandTotalHelper(ExcelPivotTable pivotTable) : base(pivotTable)
		{
			base.OuterLoop = base.PivotTable.ColumnHeaders;
			base.InnerLoop = base.PivotTable.RowHeaders;
			base.GrandTotals = new double?[base.InnerLoop.Count, base.PivotTable.DataFields.Count];
			base.HasOuterGrandTotals = base.PivotTable.ColumnGrandTotals;
			base.HasInnerGrandTotals = base.PivotTable.RowGrandTotals;
			base.HasOuterDataFields = base.PivotTable.HasColumnDataFields;
			base.OuterCellIndex = base.DataStartColumn;
			base.StartIndex = this.InnerCellIndex = base.DataStartRow;
			base.ShouldWriteInnerGrandTotals = this.OuterLoop.Count > 0 
				&& base.PivotTable.ColumnFields.Count > 0 
				&& base.PivotTable.ColumnFields[0].Index != -2;
		}
		#endregion

		#region GrandTotalHelperBase Overrides
		/// <summary>
		/// Gets the <see cref="ExcelRange"/> of the worksheet to write to.
		/// </summary>
		/// <returns>The worksheet cell to write to.</returns>
		protected override ExcelRange GetCell()
		{
			return base.PivotTable.WorkSheet.Cells[base.InnerCellIndex, base.OuterCellIndex];
		}

		/// <summary>
		/// Calculate the column grand total.
		/// </summary>
		/// <param name="innerHeaderIndex">The index of the inner 'for' loop.</param>
		/// <param name="value">TThe value in the current cell.</param>
		/// <param name="outerHeader">The <see cref="PivotTableHeader"/>.</param>
		/// <param name="grandTotal">The current grand total value that needs to be updated.</param>
		/// <returns>The updated grand total value.</returns>
		protected override double? CalculateTotal(int innerHeaderIndex, double value, PivotTableHeader outerHeader, double? grandTotal)
		{
			// If the subtotals are displayed at the top, then add all the root nodes' values to the column grand totals.
			// Otherwise, add all leaf node values.
			var pivotTableField = this.PivotTable.RowHeaders[innerHeaderIndex].PivotTableField;
			if (pivotTableField.SubtotalTop && pivotTableField.DefaultSubtotal)
			{
				if (string.IsNullOrEmpty(this.PivotTable.RowHeaders[innerHeaderIndex].SumType) && this.PivotTable.RowItems[innerHeaderIndex].RepeatedItemsCount == 0)
					grandTotal = (grandTotal ?? 0) + value;
			}
			else if (string.IsNullOrEmpty(this.PivotTable.RowHeaders[innerHeaderIndex].SumType))
				grandTotal = (grandTotal ?? 0) + value;

			// Add the value to the corresponding row grand total if it is not a subtotal node.
			if (string.IsNullOrEmpty(outerHeader.SumType))
			{
				var totalValue = base.GrandTotals[innerHeaderIndex, outerHeader.DataFieldCollectionIndex];
				base.GrandTotals[innerHeaderIndex, outerHeader.DataFieldCollectionIndex] = (totalValue ?? 0) + value;
			}
			return grandTotal;
		}

		/// <summary>
		/// Write the grand total value to the worksheet.
		/// </summary>
		/// <param name="dataField">The data field index.</param>
		protected override void WriteGrandTotals(int dataField)
		{
			for (int row = 0; row < base.GrandTotals.GetLength(0); row++)
			{
				this.PivotTable.WorkSheet.Cells[this.DataStartRow + row, this.OuterCellIndex + dataField].Value = base.GrandTotals[row, dataField];
			}
		}

		/// <summary>
		/// Store the grand total for rows to write out at the end.
		/// </summary>
		/// <param name="innerHeaderIndex">The index of the inner 'for' loop.</param>
		/// <param name="dataFieldCollectionIndex">The collection index of the data field.</param>
		/// <param name="total">The value in the current cell.</param>
		protected override void StoreTotal(int innerHeaderIndex, int dataFieldCollectionIndex, double? total)
		{
			var totalValue = this.GrandTotals[innerHeaderIndex, dataFieldCollectionIndex];
			this.GrandTotals[innerHeaderIndex, dataFieldCollectionIndex] = (totalValue ?? 0) + total;
		}
		#endregion
	}
}