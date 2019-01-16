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
		protected PivotCellBackingData[,] BackingData { get; }

		/// <summary>
		/// Gets or sets the major axis <see cref="PivotTableHeader"/> collection.
		/// </summary>
		protected List<PivotTableHeader> MajorHeaderCollection { get; set; }

		/// <summary>
		/// Gets or sets the minor axis <see cref="PivotTableHeader"/> collection.
		/// </summary>
		protected List<PivotTableHeader> MinorHeaderCollection { get; set; }

		/// <summary>
		/// Gets the totals calculation helper class.
		/// </summary>
		protected TotalsFunctionHelper TotalsCalculator { get; }
		#endregion

		#region Constructors
		/// <summary>
		/// Create a new <see cref="GrandTotalHelperBase"/> object.
		/// </summary>
		/// <param name="pivotTable">The <see cref="ExcelPivotTable"/>.</param>
		/// <param name="backingData">The data backing the pivot table.</param>
		/// <param name="totalsCalculator">The calculation helper.</param>
		protected GrandTotalHelperBase(ExcelPivotTable pivotTable, PivotCellBackingData[,] backingData, TotalsFunctionHelper totalsCalculator)
		{
			if (pivotTable == null)
				throw new ArgumentNullException(nameof(pivotTable));
			this.PivotTable = pivotTable;
			this.BackingData = backingData;
			this.TotalsCalculator = totalsCalculator;
		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Updates the grand totals in the pivot table.
		/// </summary>
		/// <returns>The list of values used to calcluate grand-grand totals.</returns>
		public PivotCellBackingData[] UpdateGrandTotals()
		{
			var grandTotalValueLists = new PivotCellBackingData[this.PivotTable.DataFields.Count];
			var grandGrandTotalValueLists = new PivotCellBackingData[this.PivotTable.DataFields.Count];
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
					dataFieldCollectionIndex = this.AddMatchingValues(majorHeader, majorIndex, minorIndex, 
						dataFieldCollectionIndex, grandTotalValueLists, grandGrandTotalValueLists);
				}
				if (dataFieldCollectionIndex != -1)
					this.WriteGrandTotal(majorIndex, grandTotalValueLists);
			}
			return grandGrandTotalValueLists;
		}

		/// <summary>
		/// Updates the grand-grand totals in a pivot table (bottom right corner totals).
		/// </summary>
		/// <param name="grandTotalsBackingData">The data backing the grand totals.</param>
		public void UpdateGrandGrandTotals(PivotCellBackingData[] grandTotalsBackingData)
		{
			int startIndex = this.GetStartIndex();
			for (int i = 0; i < grandTotalsBackingData.Length; i++)
			{
				var totalsBackingDatas = grandTotalsBackingData[i];
				this.WriteCellTotal(startIndex + i, this.PivotTable.DataFields[i], totalsBackingDatas);
			}
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
			PivotCellBackingData[] grandTotalValueLists,
			PivotCellBackingData[] grandGrandTotalValueLists);

		/// <summary>
		/// Calculates and writes the grand total values to the worksheet.
		/// </summary>
		/// <param name="majorIndex">The current major axis index.</param>
		/// <param name="grandTotalValueLists">The values used to calculate grand totals.</param>
		protected abstract void WriteGrandTotal(
			int majorIndex, 
			PivotCellBackingData[] grandTotalValueLists);

		/// <summary>
		/// Gets the start cell index for the grand-grand total values.
		/// </summary>
		/// <returns>The start cell index for the grand-grand total values.</returns>
		protected abstract int GetStartIndex();

		/// <summary>
		/// Writes the grand total for the specified <paramref name="backingData"/> in the cell at the specified <paramref name="index"/>.
		/// </summary>
		/// <param name="index">The major index of the cell to write the total to.</param>
		/// <param name="dataField">The data field to use the number format of.</param>
		/// <param name="backingData">The values to use to calculate the total.</param>
		protected abstract void WriteCellTotal(int index, ExcelPivotTableDataField dataField, PivotCellBackingData backingData);
		#endregion

		#region Protected Methods
		/// <summary>
		/// Adds backing data to the grand totals collection.
		/// </summary>
		/// <param name="index1">The major index of the backing data.</param>
		/// <param name="index2">The minor index of the backing data.</param>
		/// <param name="dataFieldCollectionIndex">The index into the data field collection.</param>
		/// <param name="backingDatas">The pivot table backing data.</param>
		protected void AddGrandTotalsBackingData(int index1, int index2, int dataFieldCollectionIndex, PivotCellBackingData[] backingDatas)
		{
			if (backingDatas[dataFieldCollectionIndex] == null)
				backingDatas[dataFieldCollectionIndex] = this.BackingData[index1, index2].Clone();
			else
				backingDatas[dataFieldCollectionIndex].Merge(this.BackingData[index1, index2]);
		}
		#endregion
	}
}