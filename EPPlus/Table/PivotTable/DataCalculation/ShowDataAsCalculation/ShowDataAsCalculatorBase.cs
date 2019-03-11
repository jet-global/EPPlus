using System.Collections.Generic;

namespace OfficeOpenXml.Table.PivotTable.DataCalculation.ShowDataAsCalculation
{
	/// <summary>
	/// Base class for calculating a <see cref="ShowDataAs"/> value in a pivot table.
	/// </summary>
	internal abstract class ShowDataAsCalculatorBase
	{
		#region Properties
		/// <summary>
		/// Gets the pivot table that this calculator is calculating against.
		/// </summary>
		protected ExcelPivotTable PivotTable { get; }
		
		/// <summary>
		/// Gets the index of the data field that is being calculated.
		/// </summary>
		protected int DataFieldCollectionIndex { get; }
		#endregion

		#region Constructors
		/// <summary>
		/// Base calculator constructor.
		/// </summary>
		/// <param name="pivotTable">The pivot table to calculate against.</param>
		/// <param name="dataFieldCollectionIndex">The index of the data field that the calculator is calculating.</param>
		public ShowDataAsCalculatorBase(ExcelPivotTable pivotTable, int dataFieldCollectionIndex)
		{
			this.PivotTable = pivotTable;
			this.DataFieldCollectionIndex = dataFieldCollectionIndex;
		}
		#endregion

		#region Abstract Methods
		/// <summary>
		/// Calculates a body value in a pivot table cell.
		/// </summary>
		/// <param name="dataRow">The row in the backing body data.</param>
		/// <param name="dataColumn">The column in the backing body data.</param>
		/// <param name="backingDatas">The backing data for the pivot table body.</param>
		/// <param name="grandGrandTotalValues">The backing data for the pivot table grand grand totals.</param>
		/// <param name="rowGrandTotalsValuesLists">The backing data for the pivot table row grand totals.</param>
		/// <param name="columnGrandTotalsValuesLists">The backing data for the pivot table column grand totals.</param>
		/// <param name="totalsCalculator">A <see cref="TotalsFunctionHelper"/> to calculate values with.</param>
		/// <returns>An object value for the cell.</returns>
		public abstract object CalculateBodyValue(
			int dataRow, int dataColumn,
			PivotCellBackingData[,] backingDatas,
			PivotCellBackingData[] grandGrandTotalValues,
			List<PivotCellBackingData> rowGrandTotalsValuesLists,
			List<PivotCellBackingData> columnGrandTotalsValuesLists,
			TotalsFunctionHelper totalsCalculator);

		/// <summary>
		/// Calculates the grand total value in a pivot table cell.
		/// </summary>
		/// <param name="index">The index into the backing data.</param>
		/// <param name="grandTotalsBackingDatas">The backing data for grand totals.</param>
		/// <param name="columnGrandGrandTotalValues">The backing data for the column grand grand totals.</param>
		/// <param name="isRowTotal">A value indicating whether or not this calculation is for row totals.</param>
		/// <returns>An object value for the cell.</returns>
		public abstract object CalculateGrandTotalValue(
			int index, 
			List<PivotCellBackingData> grandTotalsBackingDatas, 
			PivotCellBackingData[] columnGrandGrandTotalValues, 
			bool isRowTotal);

		/// <summary>
		/// Calculates a grand grand total value for a pivot table cell.
		/// </summary>
		/// <param name="backingData">The backing data for the grand total cell to calculate.</param>
		/// <returns>An object value for the cell.</returns>
		public abstract object CalculateGrandGrandTotalValue(PivotCellBackingData backingData);
		#endregion
	}
}
