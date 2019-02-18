/*******************************************************************************
* You may amend and distribute as you like, but don't remove this header!
*
* EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
* See http://www.codeplex.com/EPPlus for details.
*
* Copyright (C) 2011-2019 Michelle Lau and others as noted in the source history.
*
* This library is free software; you can redistribute it and/or
* modify it under the terms of the GNU Lesser General Public
* License as published by the Free Software Foundation; either
* version 2.1 of the License, or (at your option) any later version.
* This library is distributed in the hope that it will be useful,
* but WITHOUT ANY WARRANTY; without even the implied warranty of
* MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
* See the GNU Lesser General Public License for more details.
*
* The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
* If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
*
* All code and executables are provided "as is" with no warranty either express or implied. 
* The author accepts no liability for any damage or loss of business that this product may cause.
*
* For code change notes, see the source control history.
*******************************************************************************/
namespace OfficeOpenXml.Table.PivotTable.DataCalculation
{
	/// <summary>
	/// Calculate column grand totals (grand totals on the right side of the pivot table).
	/// </summary>
	internal class ColumnGrandTotalHelper : GrandTotalHelperBase
	{
		#region Constructors
		/// <summary>
		/// Create a new <see cref="ColumnGrandTotalHelper"/> object.
		/// </summary>
		/// <param name="pivotTable">The <see cref="ExcelPivotTable"/>.</param>
		/// <param name="backingData">The data backing the pivot table.</param>
		/// <param name="totalsCalculator">The calculation helper.</param>
		internal ColumnGrandTotalHelper(ExcelPivotTable pivotTable, PivotCellBackingData[,] backingData, TotalsFunctionHelper totalsCalculator) 
			: base(pivotTable, backingData, totalsCalculator)
		{
			this.MajorHeaderCollection = this.PivotTable.RowHeaders;
			this.MinorHeaderCollection = this.PivotTable.ColumnHeaders;
		}
		#endregion

		#region GrandTotalHelperBase Overrides
		/// <summary>
		/// Adds matching values to the <paramref name="grandTotalValueList"/> and <paramref name="grandGrandTotalValueList"/>.
		/// </summary>
		/// <param name="majorHeader">The major axis header.</param>
		/// <param name="majorIndex">The current major axis index.</param>
		/// <param name="minorIndex">The current minor axis index.</param>
		/// <param name="dataFieldCollectionIndex">The index of the data field in the data field collection.</param>
		/// <param name="grandTotalValueList">The list of values used to calculate grand totals for a row or column.</param>
		/// <param name="grandGrandTotalValueList">The list of values used to calcluate the grand-grand total values.</param>
		/// <returns>The index of the data field in the data field collection.</returns>
		protected override int AddMatchingValues(
			PivotTableHeader majorHeader,
			int majorIndex,
			int minorIndex,
			int dataFieldCollectionIndex,
			PivotCellBackingData[] grandTotalValueList,
			PivotCellBackingData[] grandGrandTotalValueList)
		{
			if (this.BackingData[majorIndex, minorIndex] == null)
				return dataFieldCollectionIndex;
			var minorHeader = this.MinorHeaderCollection[minorIndex];
			dataFieldCollectionIndex = this.PivotTable.HasRowDataFields ? majorHeader.DataFieldCollectionIndex : minorHeader.DataFieldCollectionIndex;
			if (minorHeader.IsLeafNode)
			{
				base.AddGrandTotalsBackingData(majorIndex, minorIndex, dataFieldCollectionIndex, grandTotalValueList);
				// Only add row header leaf node values for grand-grand totals.
				if (majorHeader.IsLeafNode)
					base.AddGrandTotalsBackingData(majorIndex, minorIndex, dataFieldCollectionIndex, grandGrandTotalValueList);
			}
			return dataFieldCollectionIndex;
		}

		/// <summary>
		/// Calculates the grand total values and updates the backing data.
		/// </summary>
		/// <param name="majorIndex">The current major axis index.</param>
		/// <param name="grandTotalValueLists">The values used to calculate grand totals.</param>
		/// <param name="totalFunctionType">The type of function that the subtotal should be calculated with.</param>
		protected override void CalculateBackingDataTotal(
			int majorIndex, 
			PivotCellBackingData[] grandTotalValueLists,
			string totalFunctionType)
		{
			var row = this.PivotTable.Address.Start.Row + this.PivotTable.FirstDataRow + majorIndex;
			var column = this.PivotTable.Address.End.Column;
			if (this.PivotTable.HasColumnDataFields)
				column -= this.PivotTable.DataFields.Count - 1;
			for (int i = 0; i < grandTotalValueLists.Length; i++)
			{
				if (grandTotalValueLists[i] != null)
				{
					var dataField = this.PivotTable.DataFields[i];
					var result = base.TotalsCalculator.CalculateCellTotal(dataField, grandTotalValueLists[i], rowTotalType: totalFunctionType);
					var backingData = grandTotalValueLists[i];
					backingData.Result = result;
					backingData.SheetRow = row;
					backingData.SheetColumn = column++;
					backingData.DataFieldCollectionIndex = i;
				}
			}
		}

		/// <summary>
		/// Gets the start cell index for the grand-grand total values.
		/// </summary>
		/// <returns>The start cell index for the grand-grand total values.</returns>
		protected override int GetStartIndex()
		{
			return this.PivotTable.Address.End.Column - this.PivotTable.DataFields.Count + 1;
		}

		/// <summary>
		/// Updates the specified <paramref name="backingData"/> with the grand total result and corresponding cell location.
		/// </summary>
		/// <param name="index">The major index of the corresponding cell.</param>
		/// <param name="dataFieldCollectionIndex">The index of the data field to use the formula of.</param>
		/// <param name="backingData">The values to use to calculate the total.</param>
		protected override void UpdateGrandGrandTotalBackingDataTotal(int index, int dataFieldCollectionIndex, PivotCellBackingData backingData)
		{
			var dataField = this.PivotTable.DataFields[dataFieldCollectionIndex];
			var value = base.TotalsCalculator.CalculateCellTotal(dataField, backingData);
			backingData.SheetRow = this.PivotTable.Address.End.Row;
			backingData.SheetColumn = index;
			backingData.DataFieldCollectionIndex = dataFieldCollectionIndex;
			backingData.Result = value;
		}
		#endregion
	}
}