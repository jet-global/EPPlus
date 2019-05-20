/*******************************************************************************
* You may amend and distribute as you like, but don't remove this header!
*
* EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
* See http://www.codeplex.com/EPPlus for details.
*
* Copyright (C) 2011-2019 Evan Schallerer and others as noted in the source history.
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
using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.Extensions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.Table.PivotTable.DataCalculation.ShowDataAsCalculation;

namespace OfficeOpenXml.Table.PivotTable.DataCalculation
{
	/// <summary>
	/// Updates a worksheet with a pivot table's data.
	/// </summary>
	internal class PivotTableDataManager
	{
		#region Properties
		private ExcelPivotTable PivotTable { get; }

		private TotalsFunctionHelper TotalsCalculator { get; set; }
		#endregion

		#region Constructors
		/// <summary>
		/// Constructor.
		/// </summary>
		/// <param name="pivotTable">The pivot table.</param>
		public PivotTableDataManager(ExcelPivotTable pivotTable)
		{
			this.PivotTable = pivotTable;
		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Updates the pivot table's worksheet with the latest calculated data.
		/// </summary>
		public void UpdateWorksheet(List<Tuple<int, List<string>>> conditionalFormattingAddress)
		{
			using (var totalsCalculator = new TotalsFunctionHelper())
			{
				this.TotalsCalculator = totalsCalculator;
				// If the workbook has calculated fields, configure the calculation helper and cache fields appropriately.
				var calculatedFields = this.PivotTable.CacheDefinition.CacheFields.Where(c => !string.IsNullOrEmpty(c.Formula));
				if (calculatedFields.Any())
					PivotTableDataManager.ConfigureCalculatedFields(calculatedFields, totalsCalculator, this.PivotTable);

				// Generate backing body data.
				var backingBodyData = this.GetPivotTableBodyBackingData();

				// Calculate grand (and grand-grand) totals, but don't write out the values yet.
				var columnGrandTotalHelper = new ColumnGrandTotalHelper(this.PivotTable, backingBodyData, totalsCalculator);
				var columnGrandGrandTotalsLists = columnGrandTotalHelper.UpdateGrandTotals(out var columnGrandTotalBackingData);
				var rowGrandTotalHelper = new RowGrandTotalHelper(this.PivotTable, backingBodyData, totalsCalculator);
				rowGrandTotalHelper.UpdateGrandTotals(out var rowGrandTotalBackingData);
				if (this.PivotTable.HasRowDataFields)
					rowGrandTotalHelper.CalculateGrandGrandTotals(columnGrandGrandTotalsLists);
				else
					columnGrandTotalHelper.CalculateGrandGrandTotals(columnGrandGrandTotalsLists);

				// Generate row and column grand grand totals backing data
				if (this.PivotTable.ColumnGrandTotals && this.PivotTable.RowGrandTotals && this.PivotTable.ColumnFields.Any())
				{
					// Write grand-grand totals to worksheet (grand totals at bottom right corner of pivot table).
					this.WriteGrandGrandTotals(columnGrandGrandTotalsLists);
				}

				// Write out row and column grand grand totals.
				if (this.PivotTable.ColumnGrandTotals)
					this.WriteGrandTotalValues(false, columnGrandTotalBackingData, columnGrandGrandTotalsLists);
				if (this.PivotTable.RowGrandTotals)
					this.WriteGrandTotalValues(true, rowGrandTotalBackingData, columnGrandGrandTotalsLists);

				// Write out body data applying "Show Data As" and other settings as necessary.
				this.WritePivotTableBodyData(backingBodyData, columnGrandTotalBackingData, 
					rowGrandTotalBackingData, columnGrandGrandTotalsLists, totalsCalculator, conditionalFormattingAddress);
			}
		}
		#endregion

		#region Public Static Methods
		/// <summary>
		/// Gets the backing cell values for a given set of row header and column header indices and a data field.
		/// </summary>
		/// <param name="pivotTable">The pivot table to get backing values for.</param>
		/// <param name="dataFieldCollectionIndex">The index of the data field to get backing values for.</param>
		/// <param name="rowHeaderIndices">The row indices to filter values down by.</param>
		/// <param name="columnHeaderIndices">The column indices to filter values down by.</param>
		/// <param name="rowHeaderTotalType">The row function type to calculate values with.</param>
		/// <param name="columnHeaderTotalType">The column function type to calculate values with.</param>
		/// <param name="functionCalculator">The <see cref="TotalsFunctionHelper"/> to perform calculations with.</param>
		/// <returns>A <see cref="PivotCellBackingData"/> containing the backing values and a calculated result.</returns>
		public static PivotCellBackingData GetParentBackingCellValues(
			ExcelPivotTable pivotTable,
			int dataFieldCollectionIndex,
			List<Tuple<int, int>> rowHeaderIndices,
			List<Tuple<int, int>> columnHeaderIndices,
			string rowHeaderTotalType,
			string columnHeaderTotalType,
			TotalsFunctionHelper functionCalculator)
		{
			var dataField = pivotTable.DataFields[dataFieldCollectionIndex];
			var cacheField = pivotTable.CacheDefinition.CacheFields[dataField.Index];
			PivotCellBackingData backingData = null;
			if (string.IsNullOrEmpty(cacheField.Formula))
			{
				var matchingValues = pivotTable.CacheDefinition.CacheRecords.FindMatchingValues(
					rowHeaderIndices,
					columnHeaderIndices,
					pivotTable.GetPageFieldIndices(),
					dataField.Index,
					pivotTable,
					true);
				backingData = new PivotCellBackingData(matchingValues);
			}
			else
			{
				// If a formula is present, it is a calculated field which needs to be evaluated.
				var fieldNameToValues = new Dictionary<string, List<object>>();
				foreach (var cacheFieldName in cacheField.ReferencedCacheFieldsToIndex.Keys)
				{
					var values = pivotTable.CacheDefinition.CacheRecords.FindMatchingValues(
						rowHeaderIndices,
						columnHeaderIndices,
						pivotTable.GetPageFieldIndices(),
						cacheField.ReferencedCacheFieldsToIndex[cacheFieldName],
						pivotTable,
						true);
					fieldNameToValues.Add(cacheFieldName, values);
				}
				backingData = new PivotCellBackingData(fieldNameToValues, cacheField.ResolvedFormula);
			}
			var value = functionCalculator.CalculateCellTotal(dataField, backingData, rowHeaderTotalType, columnHeaderTotalType);
			if (backingData != null)
				backingData.Result = value;
			return backingData;
		}

		/// <summary>
		/// Resolve the name references and other formulas contained in a formula.
		/// </summary>
		/// <param name="calculatedFields">The list of calculated fields in the pivot table.</param>
		/// <param name="totalsCalculator">The function helper calculator.</param>
		/// <param name="pivotTable">The pivot table the fields are on.</param>
		public static void ConfigureCalculatedFields(IEnumerable<CacheFieldNode> calculatedFields, TotalsFunctionHelper totalsCalculator, ExcelPivotTable pivotTable)
		{
			// Add all of the cache field names to the calculation helper.
			var cacheFieldNames = new HashSet<string>(pivotTable.CacheDefinition.CacheFields.Select(c => c.Name));
			totalsCalculator.AddNames(cacheFieldNames);

			// Resolve any calclulated fields that may be referencing each other to forumlas composed of regular ol' cache fields.
			foreach (var calculatedField in calculatedFields)
			{
				var resolvedFormulaTokens = PivotTableDataManager.ResolveFormulaReferences(calculatedField.Formula, totalsCalculator, calculatedFields);
				foreach (var token in resolvedFormulaTokens.Where(t => t.TokenType == TokenType.NameValue))
				{
					if (!calculatedField.ReferencedCacheFieldsToIndex.ContainsKey(token.Value))
					{
						var referencedFieldIndex = pivotTable.CacheDefinition.GetCacheFieldIndex(token.Value);
						calculatedField.ReferencedCacheFieldsToIndex.Add(token.Value, referencedFieldIndex);
					}
				}
				// Reconstruct the formula and wrap all field names in single ticks.
				string resolvedFormula = string.Empty;
				foreach (var token in resolvedFormulaTokens)
				{
					string tokenValue = token.Value;
					if (token.TokenType == TokenType.NameValue)
						tokenValue = $"'{tokenValue}'";
					resolvedFormula += tokenValue;
				}
				calculatedField.ResolvedFormula = resolvedFormula;
			}
		}
		#endregion

		#region Private Methods
		private void WritePivotTableBodyData(PivotCellBackingData[,] backingDatas,
			List<PivotCellBackingData> columnGrandTotalsValuesLists, List<PivotCellBackingData> rowGrandTotalsValuesLists,
			PivotCellBackingData[] grandGrandTotalValues, TotalsFunctionHelper totalsCalculator, List<Tuple<int, List<string>>> conditionalFormattingAddress)
		{
			int sheetColumn = this.PivotTable.Address.Start.Column + this.PivotTable.FirstDataCol;
			for (int column = 0; column < this.PivotTable.ColumnHeaders.Count; column++)
			{
				var columnHeader = this.PivotTable.ColumnHeaders[column];
				int sheetRow = this.PivotTable.Address.Start.Row + this.PivotTable.FirstDataRow - 1;
				for (int row = 0; row < this.PivotTable.RowHeaders.Count; row++)
				{
					sheetRow++;
					var rowHeader = this.PivotTable.RowHeaders[row];
					if (rowHeader.IsGrandTotal || columnHeader.IsGrandTotal || !backingDatas[row, column].ShowValue)
						continue;

					var dataFieldCollectionIndex = this.PivotTable.HasRowDataFields ? rowHeader.DataFieldCollectionIndex : columnHeader.DataFieldCollectionIndex;
					var dataField = this.PivotTable.DataFields[dataFieldCollectionIndex];
					var showDataAsCalculator = ShowDataAsFactory.GetShowDataAsCalculator(dataField.ShowDataAs, this.PivotTable, dataFieldCollectionIndex, this.TotalsCalculator);
					var value = showDataAsCalculator.CalculateBodyValue(
						row, column, 
						backingDatas, 
						grandGrandTotalValues, 
						rowGrandTotalsValuesLists,
						columnGrandTotalsValuesLists);

					var cell = this.PivotTable.Worksheet.Cells[sheetRow, sheetColumn];
					this.WriteCellValue(value, cell, dataField, this.PivotTable.Workbook.Styles);

					if (rowHeader.ConditionalFormattingTupleList.Count > 0)
					{
						var currentTuples = rowHeader.ConditionalFormattingTupleList.Where(i => i.Item2 == dataFieldCollectionIndex);
						foreach (var tuple in currentTuples)
						{
							int index = conditionalFormattingAddress.FindIndex(i => i.Item1 == tuple.Item1);
							conditionalFormattingAddress[index].Item2.Add(cell.AddressSpaceSeparated + " ");
						}
					}
				}
				sheetColumn++;
			}
		}

		private void WriteGrandTotalValues(bool isRowTotal, List<PivotCellBackingData> grandTotalsBackingDatas, PivotCellBackingData[] columnGrandGrandTotalValues)
		{
			for (int i = 0; i < grandTotalsBackingDatas.Count; i++)
			{
				var grandTotalBackingData = grandTotalsBackingDatas[i];
				if (grandTotalBackingData == null || !grandTotalBackingData.ShowValue)
					continue;
				var dataField = this.PivotTable.DataFields[grandTotalBackingData.DataFieldCollectionIndex];

				var showDataAsCalculator = ShowDataAsFactory.GetShowDataAsCalculator(dataField.ShowDataAs, this.PivotTable, grandTotalBackingData.DataFieldCollectionIndex, this.TotalsCalculator);
				var value = showDataAsCalculator.CalculateGrandTotalValue(i, grandTotalsBackingDatas, columnGrandGrandTotalValues, isRowTotal);

				var cell = this.PivotTable.Worksheet.Cells[grandTotalBackingData.SheetRow, grandTotalBackingData.SheetColumn];
				this.WriteCellValue(value, cell, dataField, this.PivotTable.Workbook.Styles);
			}
		}

		private void WriteGrandGrandTotals(PivotCellBackingData[] columnGrandGrandTotalsLists)
		{
			var styles = this.PivotTable.Workbook.Styles;
			foreach (var backingData in columnGrandGrandTotalsLists)
			{
				if (backingData == null)
					continue;
				var dataField = this.PivotTable.DataFields[backingData.DataFieldCollectionIndex];

				var showDataAsCalculator = ShowDataAsFactory.GetShowDataAsCalculator(dataField.ShowDataAs, this.PivotTable, backingData.DataFieldCollectionIndex, this.TotalsCalculator);
				var value = showDataAsCalculator.CalculateGrandGrandTotalValue(backingData);

				var cell = this.PivotTable.Worksheet.Cells[backingData.SheetRow, backingData.SheetColumn];
				this.WriteCellValue(value, cell, dataField, styles);
			}
		}

		private PivotCellBackingData GetBodyBackingCellValues(
			ExcelPivotTable pivotTable,
			int dataFieldCollectionIndex,
			List<int> rowCacheRecordIndices,
			List<int> columnCacheRecordIndices,
			string rowHeaderTotalType,
			string columnHeaderTotalType,
			TotalsFunctionHelper functionCalculator)
		{
			var dataField = pivotTable.DataFields[dataFieldCollectionIndex];
			var cacheField = pivotTable.CacheDefinition.CacheFields[dataField.Index];
			PivotCellBackingData backingData = null;
			if (string.IsNullOrEmpty(cacheField.Formula))
			{
				var matchingValues = pivotTable.CacheDefinition.CacheRecords.GetDataFieldValues(
					rowCacheRecordIndices,
					columnCacheRecordIndices,
					dataField.Index,
					pivotTable.ItemsMatcher);
				backingData = new PivotCellBackingData(matchingValues);
			}
			else
			{
				// If a formula is present, it is a calculated field which needs to be evaluated.
				var fieldNameToValues = new Dictionary<string, List<object>>();
				foreach (var cacheFieldName in cacheField.ReferencedCacheFieldsToIndex.Keys)
				{
					var values = pivotTable.CacheDefinition.CacheRecords.GetDataFieldValues(
						rowCacheRecordIndices,
						columnCacheRecordIndices,
						cacheField.ReferencedCacheFieldsToIndex[cacheFieldName],
						pivotTable.ItemsMatcher);
					fieldNameToValues.Add(cacheFieldName, values);
				}
				backingData = new PivotCellBackingData(fieldNameToValues, cacheField.ResolvedFormula);
			}
			var value = functionCalculator.CalculateCellTotal(dataField, backingData, rowHeaderTotalType, columnHeaderTotalType);
			if (backingData != null)
				backingData.Result = value;
			return backingData;
		}

		private PivotCellBackingData[,] GetPivotTableBodyBackingData()
		{
			var backingData = new PivotCellBackingData[this.PivotTable.RowHeaders.Count(), this.PivotTable.ColumnHeaders.Count()];
			int dataColumn = this.PivotTable.Address.Start.Column + this.PivotTable.FirstDataCol;
			for (int column = 0; column < this.PivotTable.ColumnHeaders.Count; column++)
			{
				var columnHeader = this.PivotTable.ColumnHeaders[column];
				int dataRow = this.PivotTable.Address.Start.Row + this.PivotTable.FirstDataRow - 1;
				for (int row = 0; row < this.PivotTable.RowHeaders.Count; row++)
				{
					dataRow++;
					var rowHeader = this.PivotTable.RowHeaders[row];
					if (rowHeader.IsGrandTotal || columnHeader.IsGrandTotal)
						continue;

					var dataFieldCollectionIndex = this.PivotTable.HasRowDataFields ? rowHeader.DataFieldCollectionIndex : columnHeader.DataFieldCollectionIndex;
					backingData[row, column] = this.GetBodyBackingCellValues(
						this.PivotTable,
						dataFieldCollectionIndex,
						rowHeader.UsedCacheRecordIndices,
						columnHeader.UsedCacheRecordIndices,
						rowHeader.TotalType,
						columnHeader.TotalType,
						this.TotalsCalculator);

					if (rowHeader.IsPlaceHolder)
						backingData[row, column].ShowValue = true;
					else if ((rowHeader.CacheRecordIndices == null && columnHeader.CacheRecordIndices.Count == this.PivotTable.ColumnFields.Count)
						|| rowHeader.CacheRecordIndices.Count == this.PivotTable.RowFields.Count)
					{
						// At a leaf node.
						backingData[row, column].ShowValue = true;
					}
					else if (this.PivotTable.HasRowDataFields)
					{
						if (rowHeader.PivotTableField != null && rowHeader.PivotTableField.DefaultSubtotal && !rowHeader.TotalType.IsEquivalentTo("none"))
						{
							if ((rowHeader.PivotTableField != null && rowHeader.PivotTableField.SubtotalTop && !rowHeader.IsAboveDataField)
								|| !string.IsNullOrEmpty(rowHeader.TotalType))
							{
								backingData[row, column].ShowValue = true;
							}
						}
					}
					else if (rowHeader.PivotTableField.DefaultSubtotal && !rowHeader.TotalType.IsEquivalentTo("none")
						&& (rowHeader.TotalType != null || rowHeader.PivotTableField.SubtotalTop))
					{
						backingData[row, column].ShowValue = true;
					}
				}
				dataColumn++;
			}
			return backingData;
		}

		private void WriteCellValue(object value, ExcelRange cell, ExcelPivotTableDataField dataField, ExcelStyles styles)
		{
			cell.Value = value;
			var style = styles.NumberFormats.FirstOrDefault(n => n.NumFmtId == dataField.NumFmtId);
			if (style != null)
				cell.Style.Numberformat.Format = style.Format;
		}
		#endregion

		#region Private Static Methods
		private static List<Token> ResolveFormulaReferences(string formula, TotalsFunctionHelper totalsCalculator, IEnumerable<CacheFieldNode> calculatedFields)
		{
			var resolvedFormulaTokens = new List<Token>();
			var tokens = totalsCalculator.Tokenize(formula);
			foreach (var token in tokens)
			{
				if (token.TokenType == TokenType.NameValue)
				{
					// If a token references another calculated field, resolve the chain of formulas.
					var field = calculatedFields.FirstOrDefault(f => f.Name.IsEquivalentTo(token.Value));
					if (field != null)
					{
						var resolvedReferences = PivotTableDataManager.ResolveFormulaReferences(field.Formula, totalsCalculator, calculatedFields);
						resolvedFormulaTokens.Add(new Token("(", TokenType.OpeningParenthesis));
						resolvedFormulaTokens.AddRange(resolvedReferences);
						resolvedFormulaTokens.Add(new Token(")", TokenType.ClosingParenthesis));
					}
					else
						resolvedFormulaTokens.Add(token);
				}
				else
					resolvedFormulaTokens.Add(token);
			}
			return resolvedFormulaTokens;
		}
		#endregion
	}
}
