using System;
using System.Collections.Generic;
using System.Linq;
using static OfficeOpenXml.ExcelErrorValue;

namespace OfficeOpenXml.Table.PivotTable.DataCalculation.ShowDataAsCalculation
{
	internal abstract class PercentOfParentCalculatorBase : ShowDataAsCalculatorBase
	{
		#region Constructors
		/// <summary>
		/// Constructs the calculator.
		/// </summary>
		/// <param name="pivotTable">The pivot table to calculate against.</param>
		/// <param name="dataFieldCollectionIndex">The index of the data field that the calculator is calculating.</param>
		/// <param name="totalsCalculator">A <see cref="TotalsFunctionHelper"/> to calculate values with.</param>
		public PercentOfParentCalculatorBase(ExcelPivotTable pivotTable, int dataFieldCollectionIndex, TotalsFunctionHelper totalsCalculator) 
			: base(pivotTable, dataFieldCollectionIndex, totalsCalculator) { }
		#endregion

		#region Protected Methods
		/// <summary>
		/// Calculates a pivot table body value for show data as percent of parent [row|column|total].
		/// </summary>
		/// <param name="isParentRow">A value indicating if the parent is a row.</param>
		/// <param name="dataRow">The current row in the data.</param>
		/// <param name="dataColumn">The current column in the data.</param>
		/// <param name="parentHeaderIndices">The cache record indices of the parent header.</param>
		/// <param name="backingDatas">The backing body data for the pivot table.</param>
		/// <param name="isParentColumnTotal">Bool indicating if % of Parent Total is used.</param>
		/// <returns>An object value for a cell.</returns>
		protected object CalculateBodyValue(bool isParentRow, int dataRow, int dataColumn, List<Tuple<int, int>> parentHeaderIndices, PivotCellBackingData[,] backingDatas, bool isParentColumnTotal = false)
		{
			var rowHeader = base.PivotTable.RowHeaders[dataRow];
			var columnHeader = base.PivotTable.ColumnHeaders[dataColumn];
			var cellBackingData = backingDatas[dataRow, dataColumn];
			var dataField = base.PivotTable.DataFields[base.DataFieldCollectionIndex];

			List<Tuple<int, int>> parentRowHeaderIndices, parentColumnHeaderIndices;
			string rowTotalType, columnTotalType;

			if (isParentRow)
			{
				parentRowHeaderIndices = parentHeaderIndices;
				parentColumnHeaderIndices = columnHeader.CacheRecordIndices;
				var parentRowHeader = this.FindHeader(base.PivotTable.RowHeaders, parentRowHeaderIndices, out _);
				// If the parent is a datafield no value is written out.
				if (base.PivotTable.DataFields.Any(d => d.Field == parentRowHeader.PivotTableField))
					return null;

				rowTotalType = parentRowHeader.TotalType;
				columnTotalType = columnHeader.TotalType;

				// Use the datafield tuple if it exists (for parent row calculator), otherwise, use the tuple list passed in.
				if (rowHeader.CacheRecordIndices.Any(i => i.Item1 == -2) && rowHeader.CacheRecordIndices.First().Item1 != -2)
					parentRowHeaderIndices = rowHeader.CacheRecordIndices.Where(i => i.Item1 == -2).ToList();
			}
			else
			{
				parentColumnHeaderIndices = parentHeaderIndices;
				parentRowHeaderIndices = rowHeader.CacheRecordIndices;
				var parentColumnHeader = this.FindHeader(base.PivotTable.ColumnHeaders, parentColumnHeaderIndices, out _);
				// If the parent is a datafield no value is written out.
				if (base.PivotTable.DataFields.Any(d => d.Field == parentColumnHeader.PivotTableField))
					return null;

				rowTotalType = rowHeader.TotalType;
				columnTotalType = parentColumnHeader.TotalType;
				if (isParentColumnTotal)
					parentColumnHeaderIndices = null;
			}

			var parentBackingData = PivotTableDataManager.GetBackingCellValues(
				base.PivotTable,
				base.DataFieldCollectionIndex,
				parentRowHeaderIndices,
				parentColumnHeaderIndices,
				rowTotalType,
				columnTotalType,
				base.TotalsCalculator);

			// Don't calculate the percentage if the cell contains an error value.
			if (cellBackingData?.Result != null && Values.TryGetErrorType(cellBackingData?.Result.ToString(), out _)
				|| parentBackingData.Result != null && Values.TryGetErrorType(parentBackingData.Result.ToString(), out _))
				return null;

			if (cellBackingData?.Result == null || Convert.ToDouble(cellBackingData.Result) == 0 ||
				parentBackingData.Result == null || Convert.ToDouble(parentBackingData.Result) == 0)
			{
				// If both are null, write null.
				if (parentBackingData.Result == null || Convert.ToDouble(parentBackingData.Result) == 0)
					return null;
				// If the parent has a value, write out 0.
				return 0;
			}
			return base.CalculatePercentage(Convert.ToDouble(cellBackingData.Result), Convert.ToDouble(parentBackingData.Result));
		}
		#endregion

		#region Private Methods
		private PivotTableHeader FindHeader(List<PivotTableHeader> headers, List<Tuple<int, int>> indices, out int index)
		{
			for (index = 0; index < headers.Count; index++)
			{
				var header = headers[index];
				if (this.IndexMatch(header.CacheRecordIndices, indices))
					return header;
			}
			throw new InvalidOperationException("No header was found matching the specified indices.");
		}

		private bool IndexMatch(List<Tuple<int, int>> parentIndices, List<Tuple<int, int>> containsIndices)
		{
			if (parentIndices.Count < containsIndices.Count)
				return false;
			for (int i = 0; i < containsIndices.Count; i++)
			{
				if (parentIndices[i] != containsIndices[i])
					return false;
			}
			return true;
		}
		#endregion
	}
}
