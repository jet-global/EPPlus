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
		public void UpdateWorksheet()
		{
			using (var totalsCalculator = new TotalsFunctionHelper())
			{
				this.TotalsCalculator = totalsCalculator;
				// If the workbook has calculated fields, configure the calculation helper and cache fields appropriately.
				var calculatedFields = this.PivotTable.CacheDefinition.CacheFields.Where(c => !string.IsNullOrEmpty(c.Formula));
				if (calculatedFields.Any())
					this.ConfigureCalculatedFields(calculatedFields, totalsCalculator);

				// Generate backing body data.
				var backingBodyData = this.GetPivotTableBodyBackingData();

				// Generate row and column grand totals backing data.
				PivotCellBackingData[] columnGrandGrandTotalsLists = null;
				List<PivotCellBackingData> rowGrandTotalBackingData = null, columnGrandTotalBackingData = null;
				RowGrandTotalHelper rowGrandTotalHelper = null;
				ColumnGrandTotalHelper columnGrandTotalHelper = null;
				// Calculate grand totals, but don't write out the values yet.
				if (this.PivotTable.ColumnGrandTotals)
				{
					columnGrandTotalHelper = new ColumnGrandTotalHelper(this.PivotTable, backingBodyData, totalsCalculator);
					columnGrandGrandTotalsLists = columnGrandTotalHelper.UpdateGrandTotals(out columnGrandTotalBackingData);
				}
				if (this.PivotTable.RowGrandTotals)
				{
					rowGrandTotalHelper = new RowGrandTotalHelper(this.PivotTable, backingBodyData, totalsCalculator);
					rowGrandTotalHelper.UpdateGrandTotals(out rowGrandTotalBackingData);
				}

				// Generate row and column grand grand totals backing data
				if (this.PivotTable.ColumnGrandTotals && this.PivotTable.RowGrandTotals && this.PivotTable.ColumnFields.Any())
				{
					if (this.PivotTable.HasRowDataFields)
						rowGrandTotalHelper.CalculateGrandGrandTotals(columnGrandGrandTotalsLists);
					else
						columnGrandTotalHelper.CalculateGrandGrandTotals(columnGrandGrandTotalsLists);

					// Write grand-grand totals to worksheet (grand totals at bottom right corner of pivot table).
					this.WriteGrandGrandTotals(columnGrandGrandTotalsLists);
				}

				// Write out row and column grand grand totals.
				if (this.PivotTable.ColumnGrandTotals)
					this.WriteGrandTotalValues(false, columnGrandTotalBackingData, columnGrandGrandTotalsLists);
				if (this.PivotTable.RowGrandTotals)
					this.WriteGrandTotalValues(true, rowGrandTotalBackingData, columnGrandGrandTotalsLists);

				// Write out body data applying "Show Data As" and other settings as necessary.
				this.WritePivotTableBodyData(backingBodyData, columnGrandTotalBackingData, rowGrandTotalBackingData, columnGrandGrandTotalsLists);
			}
		}
		#endregion

		#region Private Methods
		private void WritePivotTableBodyData(PivotCellBackingData[,] backingDatas,
			List<PivotCellBackingData> columnGrandTotalsValuesLists, List<PivotCellBackingData> rowGrandTotalsValuesLists,
			PivotCellBackingData[] grandGrandTotalValues)
		{
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

					var cell = this.PivotTable.Worksheet.Cells[dataRow, dataColumn];
					var dataFieldCollectionIndex = this.PivotTable.HasRowDataFields ? rowHeader.DataFieldCollectionIndex : columnHeader.DataFieldCollectionIndex;
					var dataField = this.PivotTable.DataFields[dataFieldCollectionIndex];
					var cellBackingData = backingDatas[row, column];
					var value = cellBackingData?.Result;
					
					if (dataField.ShowDataAs == ShowDataAs.NoCalculation)
					{
						// If no ShowDataAs value is selected, the "For empty cells show: [missingCaption]" setting can be applied.
						value = this.GetCellNoCalculationValue(value, cellBackingData);
					}
					else if (dataField.ShowDataAs == ShowDataAs.PercentOfTotal
						|| dataField.ShowDataAs == ShowDataAs.PercentOfCol
						|| dataField.ShowDataAs == ShowDataAs.PercentOfRow)
					{
						if (cellBackingData == null)
							value = null;
						else if (value == null)
							value = 0;
						else if (value != null)
						{
							double denominator;
							if (dataField.ShowDataAs == ShowDataAs.PercentOfTotal)
								denominator = (double)grandGrandTotalValues[dataFieldCollectionIndex].Result;
							else if (dataField.ShowDataAs == ShowDataAs.PercentOfCol)
								denominator = (double)rowGrandTotalsValuesLists.First(v => v.SheetColumn == dataColumn && v.DataFieldCollectionIndex == dataFieldCollectionIndex).Result;
							else
								denominator = (double)columnGrandTotalsValuesLists.First(v => v.SheetRow == dataRow && v.DataFieldCollectionIndex == dataFieldCollectionIndex).Result;
							value = (double)value / denominator;
						}
					}
					else if (dataField.ShowDataAs == ShowDataAs.Percent)
					{
						// TODO: Deal with "(next)" and "(previous)" options for this setting. See task #11840
						// "(next)" is stored as "1048829" and "(previous)" is "1048828".
						if (dataField.BaseItem == 1048829)
							throw new InvalidOperationException(@"'(next)' is not supported for the 'Show Data as Percent of' setting.");
						else if (dataField.BaseItem == 1048828)
							throw new InvalidOperationException(@"'(previous)' is not supported for the 'Show Data as Percent of' setting.");

						var baseFieldItemTuple = new Tuple<int, int>(dataField.BaseField, dataField.BaseItem);
						if (cellBackingData == null) { /* noop */ }
						else if (rowHeader.CacheRecordIndices.Any(t => t.Equals(baseFieldItemTuple)) || columnHeader.CacheRecordIndices.Any(t => t.Equals(baseFieldItemTuple)))
						{
							if (value != null)
								value = 1;  // At a row/column that contains the comparison field item which makes this 100%.
						}
						else
						{
							// Try to find a value that matches either the current row or column header structure.
							// If a value is found, the percentage can be calculated. Otherwise, the appropriate error is written out.
							if (this.TryFindMatchingHeaderIndex(rowHeader, baseFieldItemTuple, this.PivotTable.RowHeaders, out int headerIndex))
							{
								var baseValue = backingDatas[headerIndex, column]?.Result;
								value = this.GetShowDataAsPercentOfValue(baseValue, value);
							}
							else if (this.TryFindMatchingHeaderIndex(columnHeader, baseFieldItemTuple, this.PivotTable.ColumnHeaders, out headerIndex))
							{
								var baseValue = backingDatas[row, headerIndex]?.Result;
								value = this.GetShowDataAsPercentOfValue(baseValue, value);
							}
							else
							{
								if (!this.PivotTable.RowFields.Any(f => f.Index == dataField.BaseField) 
									&& !this.PivotTable.ColumnFields.Any(f => f.Index == dataField.BaseField))
								{
									// If the dataField.BaseField is not a row or column field, all values #N/A!
									value = ExcelErrorValue.Create(eErrorType.NA);
								}
								else if (!rowHeader.CacheRecordIndices.Any(i => i.Item1 == dataField.BaseField) 
									&& !columnHeader.CacheRecordIndices.Any(i => i.Item1 == dataField.BaseField))
									{
										// Subtotals only get a value if they are at the same depth or below the show data as field.
										value = null;
									}
								else
									value = ExcelErrorValue.Create(eErrorType.NA);
							}
						}
					}
					else
					{
						// TODO: Implement the rest of these settings. See user story 11453.
						throw new InvalidOperationException($"Unsupported dataField ShowDataAs setting '{dataField.ShowDataAs}'");
					}

					this.WriteCellValue(value, cell, dataField, this.PivotTable.Workbook.Styles);
				}
				dataColumn++;
			}
		}

		private object GetShowDataAsPercentOfValue(object baseValue, object value)
		{
			if (baseValue == null && value == null)
				return ExcelErrorValue.Create(eErrorType.Null);
			else if (baseValue == null)
				return null;
			else if (value == null)
				return ExcelErrorValue.Create(eErrorType.Null);
			else
				return (double)value / (double)baseValue;
		}

		private bool TryFindMatchingHeaderIndex(PivotTableHeader header, Tuple<int, int> baseFieldItem, List<PivotTableHeader> headers, out int headerIndex)
		{
			headerIndex = -1;
			var index = header.CacheRecordIndices.FindIndex(i => i.Item1 == baseFieldItem.Item1);
			if (index >= 0)
			{
				// The value that this will be compared against is in the cell that matches this cell's 
				// row and colum indices other than the base field/item indices.
				var indicesToMatch = header.CacheRecordIndices.ToList();
				indicesToMatch[index] = baseFieldItem;
				headerIndex = headers.FindIndex(h => this.AreEquivalent(h.CacheRecordIndices, indicesToMatch));
				return headerIndex != -1;
			}
			return false;
		}

		private bool AreEquivalent(List<Tuple<int, int>> first, List<Tuple<int, int>> second)
		{
			if (first?.Count != second?.Count)
				return false;
			for (int j = 0; j < first.Count; j++)
			{
				if (!first[j].Equals(second[j]))
					return false;
			}
			return true;
		}

		private void WriteGrandTotalValues(bool isRowTotal, List<PivotCellBackingData> grandTotalsBackingDatas, PivotCellBackingData[] columnGrandGrandTotalValues)
		{
			for (int i = 0; i < grandTotalsBackingDatas.Count; i++)
			{
				var grandTotalBackingData = grandTotalsBackingDatas[i];
				if (grandTotalBackingData == null)
					continue;
				var dataField = this.PivotTable.DataFields[grandTotalBackingData.DataFieldCollectionIndex];
				var cell = this.PivotTable.Worksheet.Cells[grandTotalBackingData.SheetRow, grandTotalBackingData.SheetColumn];
				object value = grandTotalBackingData.Result;

				if (dataField.ShowDataAs == ShowDataAs.NoCalculation)
				{
					// If no ShowDataAs value is selected, the "For empty cells show: [missingCaption]" setting can be applied.
					value = this.GetCellNoCalculationValue(value, grandTotalBackingData);
				}
				else if (dataField.ShowDataAs == ShowDataAs.PercentOfTotal 
					|| dataField.ShowDataAs == ShowDataAs.PercentOfCol 
					|| dataField.ShowDataAs == ShowDataAs.PercentOfRow)
				{
					if (columnGrandGrandTotalValues.Length > grandTotalBackingData.DataFieldCollectionIndex)
					{
						if (value == null)
							value = 0;
						else if ((dataField.ShowDataAs == ShowDataAs.PercentOfCol && isRowTotal) || (dataField.ShowDataAs == ShowDataAs.PercentOfRow && !isRowTotal))
							value = 1;
						else
						{
							double grandGrandTotalValue = (double)columnGrandGrandTotalValues[grandTotalBackingData.DataFieldCollectionIndex].Result;
							value = (double)grandTotalBackingData.Result / grandGrandTotalValue;
						}
					}
					else
						value = 1;
				}
				else if (dataField.ShowDataAs == ShowDataAs.Percent)
				{
					// TODO: Deal with "(next)" and "(previous)" options for this setting. See task #11840
					// "(next)" is stored as "1048829" and "(previous)" is "1048828".
					if (dataField.BaseItem == 1048829)
						throw new InvalidOperationException(@"'(next)' is not supported for the 'Show Data as Percent of' setting.");
					else if (dataField.BaseItem == 1048828)
						throw new InvalidOperationException(@"'(previous)' is not supported for the 'Show Data as Percent of' setting.");

					var baseFieldItemTuple = new Tuple<int, int>(dataField.BaseField, dataField.BaseItem);
					var headers = isRowTotal ? this.PivotTable.ColumnHeaders : this.PivotTable.RowHeaders;
					var header = headers[grandTotalBackingData.MajorAxisIndex];
					if (header.CacheRecordIndices.Any(t => t.Equals(baseFieldItemTuple)))
						value = 1;  // At a row/column that contains the comparison field item which makes this 100%.
					else
					{
						// Try to find a value that matches either the current row or column header structure.
						// If a value is found, the percentage can be calculated. Otherwise, the appropriate error is written out.
						if (this.TryFindMatchingHeaderIndex(header, baseFieldItemTuple, headers, out int headerIndex))
						{
							// Get the correct index into grandTotalsBackingDatas which are a 1d array 
							// representing [datafields.Count] number of rows/columns.
							var denominatorHeader = grandTotalsBackingDatas
								.Where(d => d.MajorAxisIndex == headerIndex)
								.ElementAt(grandTotalBackingData.DataFieldCollectionIndex);
							var baseValue = denominatorHeader?.Result;
							value = this.GetShowDataAsPercentOfValue(baseValue, value);
						}
						else
						{
							if (!this.PivotTable.RowFields.Any(f => f.Index == dataField.BaseField)
									&& !this.PivotTable.ColumnFields.Any(f => f.Index == dataField.BaseField))
							{
								// If the dataField.BaseField is not a row or column field, all values #N/A!
								value = ExcelErrorValue.Create(eErrorType.NA);
							}
							else if (!header.CacheRecordIndices.Any(x => x.Item1 == dataField.BaseField))
								value = null;
							else
								value = ExcelErrorValue.Create(eErrorType.NA);
						}
					}
				}
				else
				{
					// TODO: Implement the rest of these settings. See user story 11453.
					throw new InvalidOperationException($"Unsupported dataField ShowDataAs setting '{dataField.ShowDataAs}'");
				}

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
				var cell = this.PivotTable.Worksheet.Cells[backingData.SheetRow, backingData.SheetColumn];
				object value = backingData.Result;
				var dataField = this.PivotTable.DataFields[backingData.DataFieldCollectionIndex];
				if (dataField.ShowDataAs == ShowDataAs.NoCalculation)
				{
					// If no ShowDataAs value is selected, the "For empty cells show: [missingCaption]" setting can be applied.
					value = this.GetCellNoCalculationValue(value, backingData);
				}
				else if (dataField.ShowDataAs == ShowDataAs.PercentOfTotal)
					value = 1;
				else if (dataField.ShowDataAs == ShowDataAs.PercentOfCol)
					value = 1;
				else if (dataField.ShowDataAs == ShowDataAs.PercentOfRow)
					value = 1;
				else if (dataField.ShowDataAs == ShowDataAs.Percent)
				{
					if (!this.PivotTable.RowFields.Any(f => f.Index == dataField.BaseField)
						&& !this.PivotTable.ColumnFields.Any(f => f.Index == dataField.BaseField))
					{
						// If the dataField.BaseField is not a row or column field, all values #N/A!
						value = ExcelErrorValue.Create(eErrorType.NA);
					}
					else
						value = null;  // The "% Of" option doesn't write in values for grand grand totals.
				}
				else
				{
					// TODO: Implement the rest of these settings. See user story 11453.
					throw new InvalidOperationException($"Unsupported dataField ShowDataAs setting '{dataField.ShowDataAs}'");
				}

				this.WriteCellValue(value, cell, dataField, styles);
			}
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
					if (rowHeader.IsPlaceHolder)
						backingData[row, column] = this.GetBackingCellValues(rowHeader, columnHeader, this.TotalsCalculator);
					else if ((rowHeader.CacheRecordIndices == null && columnHeader.CacheRecordIndices.Count == this.PivotTable.ColumnFields.Count)
						|| rowHeader.CacheRecordIndices.Count == this.PivotTable.RowFields.Count)
					{
						// At a leaf node.
						backingData[row, column] = this.GetBackingCellValues(rowHeader, columnHeader, this.TotalsCalculator);
					}
					else if (this.PivotTable.HasRowDataFields)
					{
						if (rowHeader.PivotTableField != null && rowHeader.PivotTableField.DefaultSubtotal && !rowHeader.TotalType.IsEquivalentTo("none"))
						{
							if ((rowHeader.PivotTableField != null && rowHeader.PivotTableField.SubtotalTop && !rowHeader.IsAboveDataField)
								|| !string.IsNullOrEmpty(rowHeader.TotalType))
							{
								backingData[row, column] = this.GetBackingCellValues(rowHeader, columnHeader, this.TotalsCalculator);
							}
						}
					}
					else if (rowHeader.PivotTableField.DefaultSubtotal && !rowHeader.TotalType.IsEquivalentTo("none")
						&& (rowHeader.TotalType != null || rowHeader.PivotTableField.SubtotalTop))
						backingData[row, column] = this.GetBackingCellValues(rowHeader, columnHeader, this.TotalsCalculator);
				}
				dataColumn++;
			}
			return backingData;
		}

		private object GetCellNoCalculationValue(object value, PivotCellBackingData backingData)
		{
			// Non-null backing data indicates that this cell is eligible for a value.
			if (backingData != null && value == null)
				value = this.PivotTable.ShowMissing ? this.PivotTable.MissingCaption : "0";
			return value;
		}

		private PivotCellBackingData GetBackingCellValues(PivotTableHeader rowHeader, PivotTableHeader columnHeader, TotalsFunctionHelper functionCalculator)
		{
			var dataFieldCollectionIndex = this.PivotTable.HasRowDataFields ? rowHeader.DataFieldCollectionIndex : columnHeader.DataFieldCollectionIndex;
			var dataField = this.PivotTable.DataFields[dataFieldCollectionIndex];
			var cacheField = this.PivotTable.CacheDefinition.CacheFields[dataField.Index];
			PivotCellBackingData backingData = null;
			if (string.IsNullOrEmpty(cacheField.Formula))
			{
				var matchingValues = this.PivotTable.CacheDefinition.CacheRecords.FindMatchingValues(
					rowHeader.CacheRecordIndices,
					columnHeader.CacheRecordIndices,
					this.PivotTable.GetPageFieldIndices(),
					dataField.Index,
					this.PivotTable);
				 backingData = new PivotCellBackingData(matchingValues);
			}
			else
			{
				// If a formula is present, it is a calculated field which needs to be evaluated.
				var fieldNameToValues = new Dictionary<string, List<object>>();
				foreach (var cacheFieldName in cacheField.ReferencedCacheFieldsToIndex.Keys)
				{
					var values = this.PivotTable.CacheDefinition.CacheRecords.FindMatchingValues(
						rowHeader.CacheRecordIndices,
						columnHeader.CacheRecordIndices,
						this.PivotTable.GetPageFieldIndices(),
						cacheField.ReferencedCacheFieldsToIndex[cacheFieldName],
						this.PivotTable);
					fieldNameToValues.Add(cacheFieldName, values);
				}
				backingData = new PivotCellBackingData(fieldNameToValues, cacheField.ResolvedFormula);
			}
			var value = this.TotalsCalculator.CalculateCellTotal(dataField, backingData, rowHeader.TotalType, columnHeader.TotalType);
			if (backingData != null)
				backingData.Result = value;
			return backingData;
		}

		private void ConfigureCalculatedFields(IEnumerable<CacheFieldNode> calculatedFields, TotalsFunctionHelper totalsCalculator)
		{
			// Add all of the cache field names to the calculation helper.
			var cacheFieldNames = new HashSet<string>(this.PivotTable.CacheDefinition.CacheFields.Select(c => c.Name));
			totalsCalculator.AddNames(cacheFieldNames);

			// Resolve any calclulated fields that may be referencing each other to forumlas composed of regular ol' cache fields.
			foreach (var calculatedField in calculatedFields)
			{
				var resolvedFormulaTokens = this.ResolveFormulaReferences(calculatedField.Formula, totalsCalculator, calculatedFields);
				foreach (var token in resolvedFormulaTokens.Where(t => t.TokenType == TokenType.NameValue))
				{
					if (!calculatedField.ReferencedCacheFieldsToIndex.ContainsKey(token.Value))
					{
						var referencedFieldIndex = this.PivotTable.CacheDefinition.GetCacheFieldIndex(token.Value);
						calculatedField.ReferencedCacheFieldsToIndex.Add(token.Value, referencedFieldIndex);
					}
				}
				calculatedField.ResolvedFormula = string.Join(string.Empty, resolvedFormulaTokens.Select(t => t.Value));
			}
		}

		private List<Token> ResolveFormulaReferences(string formula, TotalsFunctionHelper totalsCalculator, IEnumerable<CacheFieldNode> calculatedFields)
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
						var resolvedReferences = this.ResolveFormulaReferences(field.Formula, totalsCalculator, calculatedFields);
						resolvedFormulaTokens.AddRange(resolvedReferences);
					}
					else
						resolvedFormulaTokens.Add(token);
				}
				else
					resolvedFormulaTokens.Add(token);
			}
			return resolvedFormulaTokens;
		}

		private void WriteCellValue(object value, ExcelRange cell, ExcelPivotTableDataField dataField, ExcelStyles styles)
		{
			cell.Value = value;
			var style = styles.NumberFormats.FirstOrDefault(n => n.NumFmtId == dataField.NumFmtId);
			if (style != null)
				cell.Style.Numberformat.Format = style.Format;
		}
		#endregion
	}
}
