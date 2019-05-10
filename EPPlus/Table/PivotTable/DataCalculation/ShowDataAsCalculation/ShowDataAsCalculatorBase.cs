using System;
using System.Collections.Generic;
using System.Linq;

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

		protected TotalsFunctionHelper TotalsCalculator { get; }
		#endregion

		#region Constructors
		/// <summary>
		/// Base calculator constructor.
		/// </summary>
		/// <param name="pivotTable">The pivot table to calculate against.</param>
		/// <param name="dataFieldCollectionIndex">The index of the data field that the calculator is calculating.</param>
		public ShowDataAsCalculatorBase(ExcelPivotTable pivotTable, int dataFieldCollectionIndex, TotalsFunctionHelper totalsCalculator)
		{
			this.PivotTable = pivotTable;
			this.DataFieldCollectionIndex = dataFieldCollectionIndex;
			this.TotalsCalculator = totalsCalculator;
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
		/// <returns>An object value for the cell.</returns>
		public abstract object CalculateBodyValue(
			int dataRow, int dataColumn,
			PivotCellBackingData[,] backingDatas,
			PivotCellBackingData[] grandGrandTotalValues,
			List<PivotCellBackingData> rowGrandTotalsValuesLists,
			List<PivotCellBackingData> columnGrandTotalsValuesLists);

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

		#region Protected Methods
		/// <summary>
		/// Calculates a grand total value for a cell.
		/// </summary>
		/// <param name="headers">The headers for the grand totals.</param>
		/// <param name="cellBackingData">The backing data for the current grand total cell.</param>
		/// <param name="isRowTotal">A value indicating if this is a row grand total.</param>
		/// <param name="isParentColumnTotal">Bool indicating if % of Parent Total is used.</param>
		/// <returns>An object value for a cell.</returns>
		protected object CalculateGrandTotalValue(List<PivotTableHeader> headers, PivotCellBackingData cellBackingData, bool isRowTotal, bool isParentColumnTotal = false)
		{
			var currentHeader = headers[cellBackingData.MajorAxisIndex];
			List<Tuple<int, int>> parentRowHeaderIndices, parentColumnHeaderIndices;
			string rowTotalType = string.Empty, columnTotalType = string.Empty;

			// Data fields do not get a value.
			if (currentHeader.CacheRecordIndices.Count == 1 && currentHeader.CacheRecordIndices.First().Item1 == -2)
				return null;

			// Find the parent indices in order to calculate the parent total value.
			var headerIndices = currentHeader.CacheRecordIndices.Take(currentHeader.CacheRecordIndices.Count - 1).ToList();
			parentRowHeaderIndices = new List<Tuple<int, int>>();
			parentColumnHeaderIndices = isParentColumnTotal ? new List<Tuple<int, int>>() : headerIndices;
			if (!isRowTotal)
			{
				// Use the datafield tuple if it exists (for parent row calculator), otherwise, use the tuple list passed in.
				if (currentHeader.CacheRecordIndices.Any(i => i.Item1 == -2) && currentHeader.CacheRecordIndices.First().Item1 != -2)
					parentColumnHeaderIndices = currentHeader.CacheRecordIndices.Where(i => i.Item1 == -2).ToList();
			}

			var parentBackingData = PivotTableDataManager.GetParentBackingCellValues(
				this.PivotTable,
				this.DataFieldCollectionIndex,
				parentRowHeaderIndices,
				parentColumnHeaderIndices,
				rowTotalType,
				columnTotalType,
				this.TotalsCalculator,
				includeHiddenValues: true);
			var baseValue = parentBackingData.Result;

			if (cellBackingData?.Result == null)
			{
				// If both are null, write null.
				if (baseValue == null)
					return null;
				// If the parent has a value, write out 0.
				return 0;
			}
			else if (baseValue == null)
				return 1;
			return this.CalculatePercentage(Convert.ToDouble(cellBackingData.Result), Convert.ToDouble(baseValue), true);
		}

		/// <summary>
		/// Calculates the percentage between two values.
		/// </summary>
		/// <param name="numerator">The numerator of the percentage.</param>
		/// <param name="denominator">The denominator of the percentage.</param>
		/// <param name="isGrandTotal">A flag indicating if this is a grand total value (optional).</param>
		/// <returns>The percentage between two values.</returns>
		protected object CalculatePercentage(object numerator, object denominator, bool isGrandTotal = false)
		{
			var denomiatorDouble = Convert.ToDouble(denominator);
			if (denomiatorDouble == 0d)
				return isGrandTotal ? null : ExcelErrorValue.Create(eErrorType.Div0); // If this is a grand total value, print null instead of an error value.
			return Convert.ToDouble(numerator) / denomiatorDouble;
		}
		#endregion
	}
}
