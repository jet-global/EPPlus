namespace OfficeOpenXml.Table.PivotTable.DataCalculation
{
	/// <summary>
	/// Calculate row grand totals.
	/// </summary>
	internal class RowGrandTotalHelper : GrandTotalHelperBase
	{
		#region Constructors
		/// <summary>
		/// Create a new <see cref="RowGrandTotalHelper"/> object.
		/// </summary>
		/// <param name="pivotTable">The <see cref="ExcelPivotTable"/>.</param>
		internal RowGrandTotalHelper(ExcelPivotTable pivotTable) : base(pivotTable)
		{
			base.OuterLoop = base.PivotTable.RowHeaders;
			base.InnerLoop = base.PivotTable.ColumnHeaders;
			base.GrandTotals = new double?[base.PivotTable.DataFields.Count, base.InnerLoop.Count];
			base.HasOuterGrandTotals = base.PivotTable.RowGrandTotals;
			base.HasOuterDataFields = base.PivotTable.HasRowDataFields;
			base.HasInnerGrandTotals = base.PivotTable.ColumnGrandTotals;
			base.OuterCellIndex = base.DataStartRow;
			base.StartIndex = base.InnerCellIndex = base.DataStartColumn;
			base.ShouldWriteInnerGrandTotals = this.OuterLoop.Count > 0 && base.PivotTable.RowFields[0].Index != -2;
		}
		#endregion

		#region GrandTotalHelperBase Overrides
		/// <summary>
		/// Gets the <see cref="ExcelRange"/> of the worksheet to write to.
		/// </summary>
		/// <returns>The worksheet cell to write to.</returns>
		protected override ExcelRange GetCell()
		{
			return this.PivotTable.WorkSheet.Cells[base.OuterCellIndex, base.InnerCellIndex];
		}

		/// <summary>
		/// Calculate the row grand total.
		/// </summary>
		/// <param name="innerHeaderIndex">The index of the inner 'for' loop.</param>
		/// <param name="value">TThe value in the current cell.</param>
		/// <param name="outerHeader">The <see cref="PivotTableHeader"/>.</param>
		/// <param name="grandTotal">The current grand total value that needs to be updated.</param>
		/// <returns>The updated grand total value.</returns>
		protected override double? CalculateTotal(int innerHeaderIndex, double value, PivotTableHeader outerHeader, double? grandTotal)
		{
			grandTotal = (grandTotal ?? 0) + value;
			if (outerHeader.IsLeafNode || string.IsNullOrEmpty(outerHeader.SumType))
			{
				var total = base.GrandTotals[outerHeader.DataFieldCollectionIndex, innerHeaderIndex];
				base.GrandTotals[outerHeader.DataFieldCollectionIndex, innerHeaderIndex] = (total ?? 0) + value;
			}
			return grandTotal;
		}

		/// <summary>
		/// Write the grand total value to the worksheet.
		/// </summary>
		/// <param name="dataField">The data field index.</param>
		protected override void WriteGrandTotals(int dataField)
		{
			int totalColumn = base.DataStartColumn;
			for (int column = 0; column < base.GrandTotals.GetLength(1); column++)
			{
				base.PivotTable.WorkSheet.Cells[base.OuterCellIndex + dataField, totalColumn + column].Value = base.GrandTotals[dataField, column];
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
			var totalValue = this.GrandTotals[dataFieldCollectionIndex, innerHeaderIndex];
			this.GrandTotals[dataFieldCollectionIndex, innerHeaderIndex] = (totalValue ?? 0) + total;
		}
		#endregion
	}
}