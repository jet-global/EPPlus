using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.Table.PivotTable.DataCalculation.ShowDataAsCalculation
{
	/// <summary>
	/// Calculates the <see cref="ShowDataAs.PercentOfCol"/> value in a pivot table.
	/// </summary>
	internal class PercentOfColCalculator : ShowDataAsCalculatorBase
	{
		#region Constructors
		/// <summary>
		/// Constructs the calculator.
		/// </summary>
		/// <param name="pivotTable">The pivot table to calculate against.</param>
		/// <param name="dataFieldCollectionIndex">The index of the data field that the calculator is calculating.</param>
		public PercentOfColCalculator(ExcelPivotTable pivotTable, int dataFieldCollectionIndex) : base(pivotTable, dataFieldCollectionIndex) { }
		#endregion

		#region ShowDataAsCalculatorBase Overrides
		/// <summary>
		/// Calculates a body value in a pivot table cell.
		/// </summary>
		/// <param name="dataRow">The row in the backing body data.</param>
		/// <param name="dataColumn">The column in the backing body data.</param>
		/// <param name="backingDatas">The backing data for the pivot table body.</param>
		/// <param name="grandGrandTotalValues">The backing data for the pivot table grand grand totals.</param>
		/// <param name="rowGrandTotalsValuesLists">The backing data for the pivot table row grand totals.</param>
		/// <param name="columnGrandTotalsValuesLists">The backing data for the pivot table column grand totals.</param>
		/// <returns>An object value for the cell.</returns>
		public override object CalculateBodyValue(
			int dataRow, int dataColumn,
			PivotCellBackingData[,] backingDatas,
			PivotCellBackingData[] grandGrandTotalValues,
			List<PivotCellBackingData> rowGrandTotalsValuesLists,
			List<PivotCellBackingData> columnGrandTotalsValuesLists)
		{
			var cellBackingData = backingDatas[dataRow, dataColumn];
			if (cellBackingData == null)
				return null;
			else if (cellBackingData.Result == null)
				return 0;
			else
			{
			var sheetColumn = this.PivotTable.Address.Start.Column + this.PivotTable.FirstDataCol + dataColumn;
				double denominator = (double)rowGrandTotalsValuesLists
					.First(v => v.SheetColumn == sheetColumn && v.DataFieldCollectionIndex == base.DataFieldCollectionIndex)
					.Result;
				return (double)cellBackingData.Result / denominator;
			}
		}

		/// <summary>
		/// Calculates the grand total value in a pivot table cell.
		/// </summary>
		/// <param name="index">The index into the backing data.</param>
		/// <param name="grandTotalsBackingDatas">The backing data for grand totals.</param>
		/// <param name="columnGrandGrandTotalValues">The backing data for the column grand grand totals.</param>
		/// <param name="isRowTotal">A value indicating whether or not this calculation is for row totals.</param>
		/// <returns>An object value for the cell.</returns>
		public override object CalculateGrandTotalValue(
			int index, 
			List<PivotCellBackingData> grandTotalsBackingDatas, 
			PivotCellBackingData[] columnGrandGrandTotalValues, 
			bool isRowTotal)
		{
			var grandTotalBackingData = grandTotalsBackingDatas[index];
			if (columnGrandGrandTotalValues.Length > grandTotalBackingData.DataFieldCollectionIndex)
			{
				if (grandTotalBackingData?.Result == null)
					return 0;
				else if (!isRowTotal)
				{
					double grandGrandTotalValue = (double)columnGrandGrandTotalValues[grandTotalBackingData.DataFieldCollectionIndex].Result;
					return (double)grandTotalBackingData.Result / grandGrandTotalValue;
				}
			}
			return 1;
		}


		/// <summary>
		/// Calculates a grand grand total value for a pivot table cell.
		/// </summary>
		/// <param name="backingData">The backing data for the grand total cell to calculate.</param>
		/// <returns>An object value for the cell.</returns>
		public override object CalculateGrandGrandTotalValue(PivotCellBackingData backingData) => 1;
		#endregion
	}
}
