using System;
using System.Collections.Generic;
using System.Linq;

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
		/// <returns>An object value for a cell.</returns>
		protected object CalculateBodyValue(bool isParentRow, int dataRow, int dataColumn, List<Tuple<int, int>> parentHeaderIndices, PivotCellBackingData[,] backingDatas)
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
			}

			var parentBackingData = PivotTableDataManager.GetBackingCellValues(
				base.PivotTable,
				base.DataFieldCollectionIndex,
				parentRowHeaderIndices,
				parentColumnHeaderIndices,
				rowTotalType,
				columnTotalType,
				base.TotalsCalculator);

			if (cellBackingData?.Result == null || Convert.ToDouble(cellBackingData.Result) == 0)
			{
				// If both are null, write null.
				if (parentBackingData.Result == null || Convert.ToDouble(parentBackingData.Result) == 0)
					return null;
				// If the parent has a value, write out 0.
				return 0;
			}
			return Convert.ToDouble(cellBackingData.Result) / Convert.ToDouble(parentBackingData.Result);
		}

		/// <summary>
		/// Calculates a grand total value for a cell.
		/// </summary>
		/// <param name="headers">The headers for the grand totals.</param>
		/// <param name="grandTotalsBackingDatas">The backing data objects for the grand totals.</param>
		/// <param name="columnGrandGrandTotalValues">The grand grand total backing data objects.</param>
		/// <param name="cellBackingData">The backing data for the current grand total cell.</param>
		/// <param name="dataField">The data field that the grand total will be calculated for.</param>
		/// <param name="isRowTotal">A value indicating if this is a row grand total.</param>
		/// <returns>An object value for a cell.</returns>
		protected object CalculateGrandTotalValue(List<PivotTableHeader> headers, List<PivotCellBackingData> grandTotalsBackingDatas,
			PivotCellBackingData[] columnGrandGrandTotalValues, PivotCellBackingData cellBackingData, ExcelPivotTableDataField dataField, bool isRowTotal)
		{
			var currentHeader = headers[cellBackingData.MajorAxisIndex];
			List<Tuple<int, int>> parentRowHeaderIndices, parentColumnHeaderIndices;
			string rowTotalType = string.Empty, columnTotalType = string.Empty;
			
			// Data fields do not get a value.
			if (currentHeader.CacheRecordIndices.Count == 1 && currentHeader.CacheRecordIndices.First().Item1 == -2)
				return null;

			// Find the parent indices in order to calculate the parent total value.
			var headerIndices = currentHeader.CacheRecordIndices.Take(currentHeader.CacheRecordIndices.Count - 1).ToList();
			if (isRowTotal)
			{
				parentColumnHeaderIndices = new List<Tuple<int, int>>();
				parentRowHeaderIndices = headerIndices;
			}
			else
			{
				parentRowHeaderIndices = new List<Tuple<int, int>>();
				parentColumnHeaderIndices = headerIndices;
			}

			var parentBackingData = PivotTableDataManager.GetBackingCellValues(
				base.PivotTable,
				base.DataFieldCollectionIndex,
				parentRowHeaderIndices,
				parentColumnHeaderIndices,
				rowTotalType,
				columnTotalType,
				base.TotalsCalculator);
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
			var result = Convert.ToDouble(cellBackingData.Result) / Convert.ToDouble(baseValue);
			return result;
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

		private List<int> FindSiblings(List<PivotTableHeader> headers, List<Tuple<int, int>> childIndices)
		{
			var siblingHeaderIndices = new List<int>();
			for (int i = 0; i < headers.Count; i++)
			{
				if (this.IsSibling(headers[i], childIndices))
					siblingHeaderIndices.Add(i);
			}
			return siblingHeaderIndices;
		}

		private bool IsSibling(PivotTableHeader header, List<Tuple<int, int>> childIndices)
		{
			var possibleSiblingIndices = header.CacheRecordIndices;
			if (childIndices == null || possibleSiblingIndices == null || childIndices.Count <= 1 || childIndices.Count != possibleSiblingIndices.Count)
				return false;
			for (int i = 0; i < childIndices.Count - 1; i++)
			{
				if (childIndices[i] != possibleSiblingIndices[i])
					return false;
			}
			return true;
		}

		#endregion
	}
}
