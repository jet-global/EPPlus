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
	/// Calculate row grand totals (grand totals at the bottom of a pivot table).
	/// </summary>
	internal class RowGrandTotalHelper : GrandTotalHelperBase
	{
		#region Constructors
		/// <summary>
		/// Create a new <see cref="RowGrandTotalHelper"/> object.
		/// </summary>
		/// <param name="pivotTable">The <see cref="ExcelPivotTable"/>.</param>
		/// <param name="backingData">The data backing the pivot table.</param>
		/// <param name="totalsCalculator">The calculation helper.</param>
		internal RowGrandTotalHelper(ExcelPivotTable pivotTable, PivotCellBackingData[,] backingData, TotalsFunctionHelper totalsCalculator) 
			: base(pivotTable, backingData, totalsCalculator)
		{
			this.MajorHeaderCollection = this.PivotTable.ColumnHeaders;
			this.MinorHeaderCollection = this.PivotTable.RowHeaders;
		}
		#endregion

		#region GrandTotalHelperBase Overrides
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
		protected override int AddMatchingValues(
			PivotTableHeader majorHeader, 
			int majorIndex, 
			int minorIndex, 
			int dataFieldCollectionIndex,
			PivotCellBackingData[] grandTotalValueLists, 
			PivotCellBackingData[] grandGrandTotalValueLists)
		{
			if (this.BackingData[minorIndex, majorIndex] == null)
				return dataFieldCollectionIndex;
			var minorHeader = this.MinorHeaderCollection[minorIndex];
			dataFieldCollectionIndex = this.PivotTable.HasRowDataFields ? minorHeader.DataFieldCollectionIndex : majorHeader.DataFieldCollectionIndex;
			if (minorHeader.IsLeafNode)
				base.AddGrandTotalsBackingData(minorIndex, majorIndex, dataFieldCollectionIndex, grandTotalValueLists);
			return dataFieldCollectionIndex;
		}

		/// <summary>
		/// Calculates the grand total values and updates the backing data.
		/// </summary>
		/// <param name="majorIndex">The current major axis index.</param>
		/// <param name="grandTotalValueLists">The values used to calculate grand totals.</param>
		/// <param name="totalType">The type of function that the subtotal should be calculated with.</param>
		protected override void CalculateBackingDataTotal(int majorIndex, PivotCellBackingData[] grandTotalValueLists, string totalType)
		{
			var row = this.PivotTable.Address.End.Row;
			if (this.PivotTable.HasRowDataFields)
				row -= this.PivotTable.DataFields.Count - 1;
			var column = this.PivotTable.Address.Start.Column + this.PivotTable.FirstDataCol + majorIndex;
			for (int i = 0; i < grandTotalValueLists.Length; i++)
			{
				if (grandTotalValueLists[i] != null)
				{
					var dataField = this.PivotTable.DataFields[i];
					var result = base.TotalsCalculator.CalculateCellTotal(dataField, grandTotalValueLists[i], columnTotalType: totalType);
					var backingData = grandTotalValueLists[i];
					backingData.Result = result;
					backingData.SheetRow = row++;
					backingData.SheetColumn = column;
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
			return this.PivotTable.Address.End.Row - this.PivotTable.DataFields.Count + 1;
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
			backingData.SheetRow = index;
			backingData.SheetColumn = this.PivotTable.Address.End.Column;
			backingData.DataFieldCollectionIndex = dataFieldCollectionIndex;
			backingData.Result = value;
		}
		#endregion
	}
}