using System;
using System.Collections.Generic;
using System.Linq;

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
		/// Gets the data that backs the pivot table values.
		/// </summary>
		protected List<object>[,] BackingData { get; }

		/// <summary>
		/// Gets or sets the major axis <see cref="PivotTableHeader"/> collection.
		/// </summary>
		protected List<PivotTableHeader> MajorHeaderCollection { get; set; }

		/// <summary>
		/// Gets or sets the minor axis <see cref="PivotTableHeader"/> collection.
		/// </summary>
		protected List<PivotTableHeader> MinorHeaderCollection { get; set; }
		#endregion

		#region Constructors
		/// <summary>
		/// Create a new <see cref="GrandTotalHelperBase"/> object.
		/// </summary>
		/// <param name="pivotTable">The <see cref="ExcelPivotTable"/>.</param>
		/// <param name="backingData">The data backing the pivot table.</param>
		protected GrandTotalHelperBase(ExcelPivotTable pivotTable, List<object>[,] backingData)
		{
			if (pivotTable == null)
				throw new ArgumentNullException(nameof(pivotTable));
			this.PivotTable = pivotTable;
			this.BackingData = backingData;
		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Updates the grand totals in the pivot table.
		/// </summary>
		/// <returns>The list of values used to calcluate grand-grand totals.</returns>
		public List<object>[] UpdateGrandTotals()
		{
			var grandTotalValueLists = new List<object>[this.PivotTable.DataFields.Count];
			var grandGrandTotalValueLists = new List<object>[this.PivotTable.DataFields.Count];
			using (var totalsCalculator = new TotalsFunctionHelper(this.PivotTable))
			{
				for (int majorIndex = 0; majorIndex < this.MajorHeaderCollection.Count; majorIndex++)
				{
					// Reset values lists.
					for (int i = 0; i < grandTotalValueLists.Count(); i++)
					{
						grandTotalValueLists[i] = null;
					}
					var majorHeader = this.MajorHeaderCollection[majorIndex];
					int dataFieldCollectionIndex = -1;
					int minorIndex = 0;
					for (; minorIndex < this.MinorHeaderCollection.Count; minorIndex++)
					{
						dataFieldCollectionIndex = this.AddMatchingValues(majorHeader, majorIndex, minorIndex, dataFieldCollectionIndex, grandTotalValueLists, grandGrandTotalValueLists);
					}
					if (dataFieldCollectionIndex != -1)
						this.WriteGrandTotal(dataFieldCollectionIndex, majorIndex, totalsCalculator, grandTotalValueLists);
				}
			}
			return grandGrandTotalValueLists;
		}
		#endregion

		#region Protected Abstract Methods
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
		protected abstract int AddMatchingValues(
			PivotTableHeader majorHeader, 
			int majorIndex, 
			int minorIndex, 
			int dataFieldCollectionIndex,
			List<object>[] grandTotalValueLists,
			List<object>[] grandGrandTotalValueLists);

		/// <summary>
		/// Calculates and writes the grand total values to the worksheet.
		/// </summary>
		/// <param name="dataFieldCollectionIndex">The index of the data field in the data field collection.</param>
		/// <param name="majorIndex">The current major axis index.</param>
		/// <param name="totalsCalculator">The grand totals calculation helper class.</param>
		/// <param name="grandTotalValueLists">The values used to calculate grand totals.</param>
		protected abstract void WriteGrandTotal(
			int dataFieldCollectionIndex, 
			int majorIndex, 
			TotalsFunctionHelper totalsCalculator, 
			List<object>[] grandTotalValueLists);
		#endregion

		#region Protected Methods
		/// <summary>
		/// Calculates and writes a grand total value to a cell at the specified <paramref name="row"/> and <paramref name="column"/>.
		/// </summary>
		/// <param name="row">The row to write the total value to.</param>
		/// <param name="column">The column to write the total value to.</param>
		/// <param name="dataField">The dataField to get the number format from.</param>
		/// <param name="values">The values to use to calculate the grand total.</param>
		/// <param name="functionCalculator">The totals calcluation helper class.</param>
		protected void WriteCellTotal(int row, int column, ExcelPivotTableDataField dataField, List<object> values, TotalsFunctionHelper functionCalculator)
		{
			// TODO: This method was copy/pasted from pivot table class, possibly share?
			var cell = this.PivotTable.WorkSheet.Cells[row, column];
			cell.Value = functionCalculator.Calculate(dataField, values);
			var style = this.PivotTable.WorkSheet.Workbook.Styles.NumberFormats.FirstOrDefault(n => n.NumFmtId == dataField.NumFmtId);
			if (style != null)
				cell.Style.Numberformat.Format = style.Format;
		}
		#endregion
	}
}