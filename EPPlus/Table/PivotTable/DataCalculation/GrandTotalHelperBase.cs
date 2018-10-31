using System;
using System.Collections.Generic;

using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.Table.PivotTable.DataCalculation
{
	/// <summary>
	/// Base class to calculate and write the grand totals.
	/// </summary>
	internal abstract class GrandTotalHelperBase
	{
		#region Properties
		/// <summary>
		/// Gets the <see cref="ExcelPivotTable"/>.
		/// </summary>
		protected ExcelPivotTable PivotTable { get; }

		/// <summary>
		/// Gets or sets the array of grand totals for rows/columns.
		/// </summary>
		protected double?[,] GrandTotals { get; set; }

		/// <summary>
		/// Gets or sets the header for the outer 'for' loop.
		/// </summary>
		protected List<PivotTableHeader> OuterLoop { get; set; }

		/// <summary>
		/// Gets or sets the header for the inner 'for' loop.
		/// </summary>
		protected List<PivotTableHeader> InnerLoop { get; set; }

		/// <summary>
		/// Gets or sets the value indicating if the outer grand totals is on.
		/// </summary>
		protected bool HasOuterGrandTotals { get; set; }

		/// <summary>
		/// Gets or sets the value indicating if the inner grand totals is on.
		/// </summary>
		protected bool HasInnerGrandTotals { get; set; }

		/// <summary>
		/// Gets or sets the value indicating if the outer header has multiple data fields.
		/// </summary>
		protected bool HasOuterDataFields { get; set; }

		/// <summary>
		/// Gets or sets the value indicating if the grand totals should be written on the worksheet.
		/// </summary>
		protected bool ShouldWriteInnerGrandTotals { get; set; }

		/// <summary>
		/// Gets the start row of the data on the worksheet.
		/// </summary>
		protected int DataStartRow { get; }

		/// <summary>
		/// Gets the start column of the data on the worksheet.
		/// </summary>
		protected int DataStartColumn { get; }

		/// <summary>
		/// Gets or sets the start index for row/column.
		/// </summary>
		protected int StartIndex { get; set; }

		/// <summary>
		/// Gets or sets the worksheet cell index for the outer 'for' loop.
		/// </summary>
		protected int OuterCellIndex { get; set; }

		/// <summary>
		/// Gets or sets the worksheet cell index for the inner 'for' loop.
		/// </summary>
		protected int InnerCellIndex { get; set; }
		#endregion

		#region Constructors
		/// <summary>
		/// Create a new <see cref="GrandTotalHelperBase"/> object.
		/// </summary>
		/// <param name="pivotTable">The <see cref="ExcelPivotTable"/>.</param>
		protected GrandTotalHelperBase(ExcelPivotTable pivotTable)
		{
			if (pivotTable == null)
				throw new ArgumentNullException(nameof(pivotTable));
			this.PivotTable = pivotTable;
			this.DataStartColumn = this.PivotTable.Address.Start.Column + this.PivotTable.FirstDataCol;
			this.DataStartRow = this.PivotTable.Address.Start.Row + this.PivotTable.FirstDataRow;
		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Calculate and update the grand totals in the <see cref="ExcelPivotTable"/>.
		/// </summary>
		public void UpdateGrandTotals()
		{
			int outerLoopCount = this.OuterLoop.Count;
			if (this.HasOuterGrandTotals && this.OuterLoop.Count > this.PivotTable.DataFields.Count)
				outerLoopCount = this.OuterLoop.Count - this.PivotTable.DataFields.Count;

			int innerLoopCount = this.HasInnerGrandTotals ? this.InnerLoop.Count - 1 : this.InnerLoop.Count;
			for (int i = 0; i < outerLoopCount; i++)
			{
				this.InnerCellIndex = this.StartIndex;
				double? total = null;
				var header = this.OuterLoop[i];
				int j = 0;
				// Calculate grand totals.
				for (; j < innerLoopCount; j++)
				{
					if (this.GetCell().Value.IsNumeric())
					{
						double value = this.GetCell().GetValue<double>();
						total = this.CalculateTotal(j, value, header, total);
					}
					this.InnerCellIndex++;
				}

				// Write in the grand totals for each outer axis.
				if (this.HasInnerGrandTotals)
				{
					this.GetCell().Value = total;
					// Only sum up the non-subtotal grand total values.
					if (total != null && string.IsNullOrEmpty(header.SumType))
						this.StoreTotal(j, header.DataFieldCollectionIndex, total);
				}
				this.OuterCellIndex++;
			}

			// Write in the grand totals for each inner axis.
			if (this.HasOuterGrandTotals && this.ShouldWriteInnerGrandTotals)
			{
				for (int dataField = 0; dataField < this.PivotTable.DataFields.Count; dataField++)
				{
					this.WriteGrandTotals(dataField);
				}
			}
		}
		#endregion

		#region Protected Abstract Methods
		/// <summary>
		/// Gets the <see cref="ExcelRange"/> of the worksheet to write to.
		/// </summary>
		/// <returns>The worksheet cell to write to.</returns>
		protected abstract ExcelRange GetCell();

		/// <summary>
		/// Calculate the grand total.
		/// </summary>
		/// <param name="innerHeaderIndex">The index of the inner 'for' loop.</param>
		/// <param name="value">TThe value in the current cell.</param>
		/// <param name="outerHeader">The <see cref="PivotTableHeader"/>.</param>
		/// <param name="grandTotal">The current grand total value that needs to be updated.</param>
		/// <returns>The updated grand total value.</returns>
		protected abstract double? CalculateTotal(int innerHeaderIndex, double value, PivotTableHeader outerHeader, double? grandTotal);

		/// <summary>
		/// Write the grand total value to the worksheet.
		/// </summary>
		/// <param name="dataField">The data field index.</param>
		protected abstract void WriteGrandTotals(int dataField);

		/// <summary>
		/// Store the grand total for rows/columns to write out at the end.
		/// </summary>
		/// <param name="innerHeaderIndex">The index of the inner 'for' loop.</param>
		/// <param name="dataFieldCollectionIndex">The collection index of the data field.</param>
		/// <param name="total">The value in the current cell.</param>
		protected abstract void StoreTotal(int innerHeaderIndex, int dataFieldCollectionIndex, double? total);
		#endregion
	}
}